# src/translator/visual_engine.py
"""
RTL Visual Engine - Robust PPTX Slide XML Transformer

Transforms LTR slides to RTL by:
1. Mirroring X coordinates for all positionable elements
2. Setting RTL text direction flags
3. Flipping text alignment (L <-> R)
4. Optionally flipping large structural shapes (arrows/banners)

Key improvements over basic ElementTree approach:
- Uses lxml for proper namespace prefix preservation (fixes "repair needed")
- Handles ALL element types: sp, pic, grpSp, cxnSp, graphicFrame
- Robust coordinate space handling for nested groups
- Smart logo detection to avoid flipping brand assets
"""

import os
from typing import Dict, Optional, Set
from dataclasses import dataclass
from lxml import etree

# ============================================================================
# NAMESPACE CONFIGURATION
# ============================================================================
# Complete namespace map for OOXML/PPTX
NAMESPACES: Dict[str, str] = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
}

# Shorthand for xpath queries
NS = NAMESPACES


def qn(tag: str) -> str:
    """Convert prefixed tag to Clark notation: 'a:off' -> '{uri}off'"""
    if ":" not in tag:
        return tag
    prefix, local = tag.split(":", 1)
    if prefix not in NS:
        raise ValueError(f"Unknown namespace prefix: {prefix}")
    return f"{{{NS[prefix]}}}{local}"


# ============================================================================
# GEOMETRY HELPERS
# ============================================================================
@dataclass
class BoundingBox:
    """Represents a shape's position and size in EMUs."""
    x: int
    y: int
    width: int
    height: int

    @property
    def right_edge(self) -> int:
        return self.x + self.width

    @property
    def aspect_ratio(self) -> float:
        return self.width / self.height if self.height > 0 else 0


def mirror_x(x: int, width: int, space_width: int, space_offset: int = 0) -> int:
    """
    Mirror an X coordinate within a coordinate space.

    Formula: new_x = space_offset + (space_width - ((x - space_offset) + width))

    This places the RIGHT edge of the shape where the LEFT edge was (mirrored).
    """
    relative_x = x - space_offset
    new_relative_x = space_width - (relative_x + width)
    return space_offset + new_relative_x


# ============================================================================
# MAIN ENGINE CLASS
# ============================================================================
class RTLVisualEngine:
    """
    Transforms PPTX slide XML from LTR to RTL layout.

    Features:
    - Preserves all namespace prefixes (critical for PowerPoint compatibility)
    - Handles nested groups with proper coordinate space transformation
    - Smart detection of logos vs structural shapes
    - Comprehensive element type support
    """

    # Shape types that should have their geometry horizontally flipped
    ARROW_GEOMETRIES: Set[str] = {
        "chevron", "homePlate", "leftArrow", "rightArrow",
        "stripedRightArrow", "bentArrow", "curvedRightArrow",
        "curvedLeftArrow", "notchedRightArrow", "leftRightArrow",
        "flowChartOffpageConnector"
    }

    # Keywords that indicate a shape is a logo (should NOT be flipped)
    LOGO_KEYWORDS: Set[str] = {
        "logo", "watermark", "brand", "trademark", "icon",
        "emblem", "badge", "seal", "copyright"
    }

    # Default slide width in EMUs (12192000 = 16:9 widescreen at 96 DPI)
    DEFAULT_SLIDE_WIDTH: int = 12192000

    def __init__(
        self,
        presentation_xml_path: str,
        slide_xml_path: str,
        layout_flip_ratio: float = 0.4,
        flip_connectors: bool = True,
        verbose: bool = True
    ):
        """
        Initialize the RTL Visual Engine.

        Args:
            presentation_xml_path: Path to presentation.xml (for slide dimensions)
            slide_xml_path: Path to the slide XML to transform
            layout_flip_ratio: Shapes wider than this ratio of slide width get flipH
            flip_connectors: Whether to mirror connector shapes (lines)
            verbose: Print transformation details
        """
        self.presentation_xml_path = presentation_xml_path
        self.slide_xml_path = slide_xml_path
        self.layout_flip_ratio = layout_flip_ratio
        self.flip_connectors = flip_connectors
        self.verbose = verbose

        # Statistics for reporting
        self.stats = {
            "shapes_mirrored": 0,
            "pictures_mirrored": 0,
            "groups_mirrored": 0,
            "connectors_mirrored": 0,
            "text_bodies_processed": 0,
            "shapes_flipped": 0,
            "logos_preserved": 0,
        }

        # Parse the slide XML using lxml (preserves namespace prefixes!)
        self.parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
        self.tree = etree.parse(slide_xml_path, self.parser)
        self.root = self.tree.getroot()

        # Extract slide dimensions from presentation.xml
        self.slide_width = self._extract_slide_width()
        self.flip_threshold = int(self.slide_width * layout_flip_ratio)

        if self.verbose:
            print(f"[RTLVisualEngine] Initialized")
            print(f"  Slide width: {self.slide_width} EMUs ({self.slide_width / 914400:.1f} inches)")
            print(f"  Flip threshold: {self.flip_threshold} EMUs ({layout_flip_ratio*100:.0f}% of width)")

    def _extract_slide_width(self) -> int:
        """Extract slide width from presentation.xml."""
        try:
            pres_tree = etree.parse(self.presentation_xml_path, self.parser)
            pres_root = pres_tree.getroot()

            # Try different possible paths for sldSz
            sld_sz = pres_root.find(".//p:sldSz", NS)
            if sld_sz is not None:
                cx = sld_sz.get("cx")
                if cx:
                    return int(cx)
        except Exception as e:
            if self.verbose:
                print(f"  Warning: Could not read slide width: {e}")

        return self.DEFAULT_SLIDE_WIDTH

    # ========================================================================
    # MAIN TRANSFORMATION ENTRY POINT
    # ========================================================================
    def transform(self) -> Dict[str, int]:
        """
        Execute the full RTL transformation on the slide.

        Returns:
            Dictionary of transformation statistics
        """
        if self.verbose:
            print(f"\n[RTLVisualEngine] Starting transformation...")

        # Find the shape tree (p:spTree) which contains all slide content
        sp_tree = self.root.find(".//p:cSld/p:spTree", NS)
        if sp_tree is None:
            raise ValueError("Could not find p:spTree in slide XML")

        # Process the entire tree with slide-level coordinate space
        self._process_container(sp_tree, self.slide_width, 0)

        if self.verbose:
            print(f"\n[RTLVisualEngine] Transformation complete:")
            for key, value in self.stats.items():
                print(f"  {key}: {value}")

        return self.stats

    def save(self, output_path: str) -> None:
        """
        Save the transformed XML to a file.

        Uses lxml's serialization which preserves namespace prefixes.
        """
        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        # Write with XML declaration, preserving original encoding
        self.tree.write(
            output_path,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True
        )

        if self.verbose:
            print(f"\n[RTLVisualEngine] Saved to: {output_path}")

    # ========================================================================
    # CONTAINER PROCESSING (Recursive)
    # ========================================================================
    def _process_container(
        self,
        container: etree._Element,
        space_width: int,
        space_offset: int
    ) -> None:
        """
        Process all child elements within a container (spTree or grpSp).

        Args:
            container: The parent element containing shapes
            space_width: Width of the coordinate space (slide or group)
            space_offset: X offset of the coordinate space
        """
        # Process regular shapes (p:sp)
        for sp in container.findall("p:sp", NS):
            try:
                self._process_shape(sp, space_width, space_offset)
            except Exception as e:
                if self.verbose:
                    name = self._get_element_name(sp)
                    print(f"  Warning: Failed to process shape '{name}': {e}")
                self.stats.setdefault("errors", 0)
                self.stats["errors"] += 1

        # Process pictures (p:pic)
        for pic in container.findall("p:pic", NS):
            try:
                self._process_picture(pic, space_width, space_offset)
            except Exception as e:
                if self.verbose:
                    name = self._get_element_name(pic)
                    print(f"  Warning: Failed to process picture '{name}': {e}")
                self.stats.setdefault("errors", 0)
                self.stats["errors"] += 1

        # Process connector shapes (p:cxnSp) - lines, arrows between shapes
        if self.flip_connectors:
            for cxn in container.findall("p:cxnSp", NS):
                try:
                    self._process_connector(cxn, space_width, space_offset)
                except Exception as e:
                    if self.verbose:
                        print(f"  Warning: Failed to process connector: {e}")
                    self.stats.setdefault("errors", 0)
                    self.stats["errors"] += 1

        # Process groups (p:grpSp) - recursively!
        for grp in container.findall("p:grpSp", NS):
            try:
                self._process_group(grp, space_width, space_offset)
            except Exception as e:
                if self.verbose:
                    print(f"  Warning: Failed to process group: {e}")
                self.stats.setdefault("errors", 0)
                self.stats["errors"] += 1

        # Process graphicFrames (charts, tables, SmartArt)
        for gfx in container.findall("p:graphicFrame", NS):
            try:
                self._process_graphicFrame(gfx, space_width, space_offset)
            except Exception as e:
                if self.verbose:
                    name = self._get_element_name(gfx)
                    print(f"  Warning: Failed to process graphicFrame '{name}': {e}")
                self.stats.setdefault("errors", 0)
                self.stats["errors"] += 1

    # ========================================================================
    # SHAPE PROCESSING
    # ========================================================================
    def _process_shape(
        self,
        element: etree._Element,
        space_width: int,
        space_offset: int
    ) -> None:
        """Process a regular shape (p:sp)."""
        bbox = self._get_bounding_box(element, "p:spPr")
        if bbox is None:
            return

        name = self._get_element_name(element)

        # Mirror the X coordinate
        new_x = mirror_x(bbox.x, bbox.width, space_width, space_offset)
        self._set_offset_x(element, "p:spPr", new_x)
        self.stats["shapes_mirrored"] += 1

        # Check if shape should be horizontally flipped
        if self._should_flip_shape(element, bbox, name):
            self._set_flip_h(element, "p:spPr")
            self.stats["shapes_flipped"] += 1

        # Process text content
        tx_body = element.find("p:txBody", NS)
        if tx_body is not None:
            self._process_text_body(tx_body)

    def _process_picture(
        self,
        element: etree._Element,
        space_width: int,
        space_offset: int
    ) -> None:
        """Process a picture (p:pic)."""
        bbox = self._get_bounding_box(element, "p:spPr")
        if bbox is None:
            return

        # Mirror the X coordinate
        new_x = mirror_x(bbox.x, bbox.width, space_width, space_offset)
        self._set_offset_x(element, "p:spPr", new_x)
        self.stats["pictures_mirrored"] += 1

        # NEVER flip images - they are decorative elements (icons, logos, photos)
        # Only their position is mirrored, not their visual content
        # This prevents icons and decorative elements from being flipped incorrectly
        self.stats["logos_preserved"] += 1

    def _process_connector(
        self,
        element: etree._Element,
        space_width: int,
        space_offset: int
    ) -> None:
        """Process a connector shape (p:cxnSp) - lines."""
        bbox = self._get_bounding_box(element, "p:spPr")
        if bbox is None:
            return

        # Mirror the X coordinate
        new_x = mirror_x(bbox.x, bbox.width, space_width, space_offset)
        self._set_offset_x(element, "p:spPr", new_x)
        self.stats["connectors_mirrored"] += 1

        # Flip the connector itself if it's directional
        # This ensures arrows point the correct direction in RTL
        xfrm = element.find("p:spPr/a:xfrm", NS)
        if xfrm is not None:
            # Flip horizontal orientation
            current_flip = xfrm.get("flipH", "0")
            new_flip = "0" if current_flip == "1" else "1"
            xfrm.set("flipH", new_flip)

    def _process_group(
        self,
        element: etree._Element,
        space_width: int,
        space_offset: int
    ) -> None:
        """Process a group (p:grpSp) and its children."""
        # Get the group's position in parent coordinate space
        grp_sp_pr = element.find("p:grpSpPr", NS)
        if grp_sp_pr is None:
            return

        xfrm = grp_sp_pr.find("a:xfrm", NS)
        if xfrm is None:
            return

        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is None or ext is None:
            return

        # Get group's bounding box in parent space
        group_x = int(off.get("x", "0"))
        group_width = int(ext.get("cx", "0"))

        # Mirror the group's position
        new_x = mirror_x(group_x, group_width, space_width, space_offset)
        off.set("x", str(new_x))
        self.stats["groups_mirrored"] += 1

        # Get the child coordinate space (chOff/chExt)
        ch_off = xfrm.find("a:chOff", NS)
        ch_ext = xfrm.find("a:chExt", NS)

        if ch_off is not None and ch_ext is not None:
            child_offset = int(ch_off.get("x", "0"))
            child_width = int(ch_ext.get("cx", group_width))
        else:
            # Fallback: use group's own dimensions
            child_offset = 0
            child_width = group_width

        # Recursively process children in the group's coordinate space
        self._process_container(element, child_width, child_offset)

    def _process_graphicFrame(
        self,
        element: etree._Element,
        space_width: int,
        space_offset: int
    ) -> None:
        """
        Process a graphicFrame (charts, tables, SmartArt, OLE objects).

        For charts and tables:
        - Mirror the position of the frame
        - For tables: also process table structure (reverse columns, translate text)
        - For charts: mirror position only (chart internals are complex)
        """
        try:
            # Get bounding box from xfrm
            xfrm = element.find(".//p:xfrm", NS)
            if xfrm is None:
                return

            off = xfrm.find("a:off", NS)
            ext = xfrm.find("a:ext", NS)
            if off is None or ext is None:
                return

            frame_x = int(off.get("x", "0"))
            frame_width = int(ext.get("cx", "0"))

            # Mirror the graphicFrame's position
            new_x = mirror_x(frame_x, frame_width, space_width, space_offset)
            off.set("x", str(new_x))

            # Determine type of graphic (chart, table, SmartArt, etc.)
            graphic = element.find(".//a:graphic", NS)
            if graphic is None:
                return

            graphic_data = graphic.find("a:graphicData", NS)
            if graphic_data is None:
                return

            uri = graphic_data.get("uri", "")

            # Process based on type
            if "chart" in uri:
                # Chart: position mirrored, but chart internals are complex
                # Note: Full chart mirroring would require parsing chart XML files
                self.stats.setdefault("charts_mirrored", 0)
                self.stats["charts_mirrored"] += 1
                if self.verbose:
                    print(f"    [Chart] Mirrored position (internal chart layout not modified)")

            elif "table" in uri:
                # Table: mirror position and process table structure
                table = graphic_data.find(".//a:tbl", NS)
                if table is not None:
                    self._process_table(table)
                self.stats.setdefault("tables_mirrored", 0)
                self.stats["tables_mirrored"] += 1

            elif "smartArt" in uri or "diagram" in uri:
                # SmartArt/Diagram: mirror position only
                self.stats.setdefault("smartart_mirrored", 0)
                self.stats["smartart_mirrored"] += 1
                if self.verbose:
                    print(f"    [SmartArt] Mirrored position")

            else:
                # Other graphic types (OLE objects, etc.)
                self.stats.setdefault("other_graphics_mirrored", 0)
                self.stats["other_graphics_mirrored"] += 1

        except Exception as e:
            # Robust error handling: log but continue processing
            if self.verbose:
                print(f"    Warning: Error processing graphicFrame: {e}")
            self.stats.setdefault("errors", 0)
            self.stats["errors"] += 1

    def _process_table(self, table: etree._Element) -> None:
        """
        Process table structure for RTL with column reversal.

        For consulting presentations, tables often have:
        - Headers in the first row/column
        - Data flowing left-to-right

        For RTL, we reverse column order to make data flow right-to-left.
        This is critical for maintaining visual hierarchy in Arabic.
        """
        try:
            # Get table grid (column definitions)
            tbl_grid = table.find("a:tblGrid", NS)

            # Reverse column order in all rows
            for tr in table.findall(".//a:tr", NS):
                # Get all cells in this row
                cells = tr.findall("a:tc", NS)

                if len(cells) <= 1:
                    # Single column or empty row, just process text
                    for tc in cells:
                        tx_body = tc.find("a:txBody", NS)
                        if tx_body is not None:
                            self._process_text_body(tx_body)
                    continue

                # Reverse cell order for multi-column rows
                # Remove all cells from row
                for tc in cells:
                    tr.remove(tc)

                # Re-add cells in reversed order
                for tc in reversed(cells):
                    tr.append(tc)

                    # Process text in each cell
                    tx_body = tc.find("a:txBody", NS)
                    if tx_body is not None:
                        self._process_text_body(tx_body)

            # Reverse column grid definitions if present
            if tbl_grid is not None:
                grid_cols = tbl_grid.findall("a:gridCol", NS)
                if len(grid_cols) > 1:
                    # Remove all columns
                    for col in grid_cols:
                        tbl_grid.remove(col)

                    # Re-add in reversed order
                    for col in reversed(grid_cols):
                        tbl_grid.append(col)

            self.stats.setdefault("table_cells_processed", 0)
            self.stats["table_cells_processed"] += len(table.findall(".//a:tc", NS))

            if self.verbose:
                print(f"    [Table] Reversed {len(table.findall('.//a:tr', NS))} rows")

        except Exception as e:
            if self.verbose:
                print(f"    Warning: Error processing table: {e}")
            self.stats.setdefault("errors", 0)
            self.stats["errors"] += 1

    # ========================================================================
    # TEXT BODY PROCESSING
    # ========================================================================
    def _process_text_body(self, tx_body: etree._Element) -> None:
        """
        Process a text body (p:txBody) for RTL text direction.

        Sets:
        - Body-level RTL flag
        - Paragraph alignment flip (L <-> R)
        - Paragraph RTL flag
        - Run-level RTL and language
        """
        self.stats["text_bodies_processed"] += 1

        # 1. Set body-level RTL (a:bodyPr)
        body_pr = tx_body.find("a:bodyPr", NS)
        if body_pr is None:
            body_pr = etree.SubElement(tx_body, qn("a:bodyPr"))
        # Note: rtlCol="1" is the correct attribute for body-level RTL
        body_pr.set("rtlCol", "1")

        # 2. Process each paragraph
        for p in tx_body.findall("a:p", NS):
            self._process_paragraph(p)

    def _process_paragraph(self, paragraph: etree._Element) -> None:
        """Process a single paragraph for RTL."""
        # Get or create paragraph properties
        p_pr = paragraph.find("a:pPr", NS)
        if p_pr is None:
            # Insert pPr at the beginning of the paragraph
            p_pr = etree.Element(qn("a:pPr"))
            paragraph.insert(0, p_pr)

        # Flip alignment: l <-> r, keep center/justified
        current_align = p_pr.get("algn")
        if current_align == "l" or current_align is None:
            p_pr.set("algn", "r")
        elif current_align == "r":
            p_pr.set("algn", "l")
        # "ctr" and "just" remain unchanged

        # Set paragraph RTL flag
        p_pr.set("rtl", "1")

        # Process runs
        for r in paragraph.findall("a:r", NS):
            self._process_run(r)

        # Also set default run properties for any new runs
        def_r_pr = p_pr.find("a:defRPr", NS)
        if def_r_pr is None:
            def_r_pr = etree.SubElement(p_pr, qn("a:defRPr"))
        def_r_pr.set("lang", "ar-SA")

    def _process_run(self, run: etree._Element) -> None:
        """
        Process a text run for RTL with Arabic font fallback.

        Sets:
        - Language to ar-SA for proper text shaping
        - Arabic font if not already specified
        - Font fallback for complex script support
        """
        r_pr = run.find("a:rPr", NS)
        if r_pr is None:
            # Insert rPr at the beginning of the run
            r_pr = etree.Element(qn("a:rPr"))
            run.insert(0, r_pr)

        # Set language to Arabic (affects font fallback and text shaping)
        r_pr.set("lang", "ar-SA")

        # Add Arabic font fallback
        # Check if there's already a latin or complex script font defined
        latin_font = r_pr.find("a:latin", NS)
        cs_font = r_pr.find("a:cs", NS)  # Complex Script font

        # If no fonts are defined, add Arabic font
        if latin_font is None and cs_font is None:
            # Add Simplified Arabic as default for Arabic text
            # This is a safe, widely-available Arabic font
            cs_font = etree.SubElement(r_pr, qn("a:cs"))
            cs_font.set("typeface", "Simplified Arabic")
            cs_font.set("pitchFamily", "34")
            cs_font.set("charset", "178")  # Arabic charset

            # Also set latin font to match (for mixed content)
            latin_font = etree.SubElement(r_pr, qn("a:latin"))
            latin_font.set("typeface", "Arial")
            latin_font.set("pitchFamily", "34")
            latin_font.set("charset", "0")

        elif cs_font is None and latin_font is not None:
            # Latin font exists but no CS font - add Arabic CS font
            cs_font = etree.SubElement(r_pr, qn("a:cs"))
            cs_font.set("typeface", "Simplified Arabic")
            cs_font.set("pitchFamily", "34")
            cs_font.set("charset", "178")

        # If CS font already exists, update it to use a good Arabic font if needed
        elif cs_font is not None:
            current_typeface = cs_font.get("typeface", "")
            # List of good Arabic fonts (in order of preference)
            arabic_fonts = ["Simplified Arabic", "Traditional Arabic", "Arabic Typesetting", "Arial"]

            # If current font is not in our list, change it to Simplified Arabic
            if current_typeface.lower() not in [f.lower() for f in arabic_fonts]:
                cs_font.set("typeface", "Simplified Arabic")
                cs_font.set("charset", "178")

    # ========================================================================
    # GEOMETRY HELPERS
    # ========================================================================
    def _get_bounding_box(
        self,
        element: etree._Element,
        sp_pr_tag: str
    ) -> Optional[BoundingBox]:
        """Extract bounding box from an element's shape properties."""
        sp_pr = element.find(sp_pr_tag, NS)
        if sp_pr is None:
            return None

        xfrm = sp_pr.find("a:xfrm", NS)
        if xfrm is None:
            return None

        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is None or ext is None:
            return None

        return BoundingBox(
            x=int(off.get("x", "0")),
            y=int(off.get("y", "0")),
            width=int(ext.get("cx", "0")),
            height=int(ext.get("cy", "0"))
        )

    def _set_offset_x(
        self,
        element: etree._Element,
        sp_pr_tag: str,
        new_x: int
    ) -> None:
        """Set the X offset of an element."""
        sp_pr = element.find(sp_pr_tag, NS)
        if sp_pr is None:
            return

        xfrm = sp_pr.find("a:xfrm", NS)
        if xfrm is None:
            return

        off = xfrm.find("a:off", NS)
        if off is not None:
            off.set("x", str(new_x))

    def _set_flip_h(self, element: etree._Element, sp_pr_tag: str) -> None:
        """Set horizontal flip on an element's transform."""
        sp_pr = element.find(sp_pr_tag, NS)
        if sp_pr is None:
            return

        xfrm = sp_pr.find("a:xfrm", NS)
        if xfrm is not None:
            xfrm.set("flipH", "1")

    def _get_element_name(self, element: etree._Element) -> str:
        """Get the name of an element from cNvPr."""
        # Try different paths for cNvPr
        for path in ["p:nvSpPr/p:cNvPr", "p:nvPicPr/p:cNvPr", ".//p:cNvPr"]:
            c_nv_pr = element.find(path, NS)
            if c_nv_pr is not None:
                return c_nv_pr.get("name", "")
        return ""

    # ========================================================================
    # SMART DETECTION LOGIC
    # ========================================================================
    def _should_flip_shape(
        self,
        element: etree._Element,
        bbox: BoundingBox,
        name: str
    ) -> bool:
        """
        Determine if a shape should be horizontally flipped (not just mirrored).

        We flip:
        - Arrow-like shapes (chevrons, arrows) WITHOUT text content
        - Very wide shapes that span significant slide width WITHOUT text

        We DON'T flip:
        - Logos
        - Small shapes
        - ANY shape that contains text (flipH would reverse the characters!)
        """
        # CRITICAL: Never flip shapes that contain text!
        # flipH on a text-containing shape causes characters to appear mirrored
        tx_body = element.find("p:txBody", NS)
        if tx_body is not None:
            # Check if there's actual text content (not just empty body)
            has_text = any(
                t.text and t.text.strip()
                for t in tx_body.findall(".//a:t", NS)
            )
            if has_text:
                return False

        # Don't flip if it's a logo
        if self._is_likely_logo(element, name, bbox):
            self.stats["logos_preserved"] += 1
            return False

        # Check for arrow-like geometry (only flip if no text)
        prst_geom = element.find(".//a:prstGeom", NS)
        if prst_geom is not None:
            geom_type = prst_geom.get("prst", "")
            if geom_type in self.ARROW_GEOMETRIES:
                return True

        # Check for wide structural shapes (banners, bars) - only if no text
        is_very_wide = bbox.width >= self.flip_threshold
        is_wide_aspect = bbox.aspect_ratio > 1.5  # Width > 1.5x height

        return is_very_wide and is_wide_aspect

    def _is_likely_logo(
        self,
        element: etree._Element,
        name: str,
        bbox: BoundingBox
    ) -> bool:
        """
        Heuristic detection of logos and brand assets.

        Logos typically:
        - Have "logo" in the name
        - Are small relative to slide
        - Are positioned in corners
        - Are roughly square
        """
        name_lower = name.lower()

        # Check name for logo keywords
        if any(keyword in name_lower for keyword in self.LOGO_KEYWORDS):
            return True

        # Check if it's a picture (pictures in corners are often logos)
        is_picture = element.tag.endswith("}pic")

        # Size heuristic: logos are usually small (< 15% of slide width)
        is_small = bbox.width < self.slide_width * 0.15

        # Position heuristic: logos are often in corners
        margin = self.slide_width * 0.1  # 10% margin
        in_left_margin = bbox.x < margin
        in_right_margin = bbox.right_edge > (self.slide_width - margin)
        in_corner = in_left_margin or in_right_margin

        # Aspect ratio heuristic: logos are often roughly square (0.5 - 2.0)
        is_squarish = 0.5 <= bbox.aspect_ratio <= 2.0 if bbox.height > 0 else True

        # Combined heuristic: pictures in corners are very likely logos
        if is_picture and is_small and in_corner:
            return True

        # For non-pictures, require all three conditions
        return is_small and in_corner and is_squarish


# ============================================================================
# MAIN EXECUTION
# ============================================================================
if __name__ == "__main__":
    import datetime
    import glob

    OUTPUT_DIR = "./output_xmls"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # Find the latest slide XML file
    slide_patterns = [
        "output_xmls/slide*_structure*.xml",
        "output_xmls/slide[0-9]*.xml",
        "slide*_structure*.xml",
        "slide[0-9]*.xml",
    ]

    slide_files = []
    for pattern in slide_patterns:
        # Exclude RTL and Final files
        for f in glob.glob(pattern):
            if "_RTL" not in f and "_Final" not in f:
                slide_files.append(f)

    if not slide_files:
        print("ERROR: No slide XML files found. Run the extractor first.")
        print(f"Searched patterns: {slide_patterns}")
        exit(1)

    INPUT_SLIDE_XML = max(slide_files, key=os.path.getmtime)

    # Find the latest presentation XML file
    pres_patterns = [
        "output_xmls/presentation*.xml",
        "presentation*.xml",
    ]

    pres_files = []
    for pattern in pres_patterns:
        pres_files.extend(glob.glob(pattern))

    if not pres_files:
        print("ERROR: No presentation XML files found. Run the extractor first.")
        print(f"Searched patterns: {pres_patterns}")
        exit(1)

    INPUT_PRES_XML = max(pres_files, key=os.path.getmtime)

    OUTPUT_FILENAME = f"slide_RTL_{timestamp}.xml"

    print(f"\n{'='*50}")
    print("RTL VISUAL TRANSFORMATION")
    print(f"{'='*50}")
    print(f"Slide XML:        {INPUT_SLIDE_XML}")
    print(f"Presentation XML: {INPUT_PRES_XML}")
    print(f"Output:           {os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)}")

    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Run transformation
    engine = RTLVisualEngine(
        presentation_xml_path=INPUT_PRES_XML,
        slide_xml_path=INPUT_SLIDE_XML,
        layout_flip_ratio=0.4,
        flip_connectors=True,
        verbose=True
    )

    stats = engine.transform()

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    engine.save(output_path)

    print(f"\n{'='*50}")
    print("TRANSFORMATION SUMMARY")
    print(f"{'='*50}")
    print(f"Input:  {INPUT_SLIDE_XML}")
    print(f"Output: {output_path}")
    print(f"Slide Width: {engine.slide_width} EMUs")
    print(f"\nStatistics:")
    for key, value in stats.items():
        print(f"  {key.replace('_', ' ').title()}: {value}")
