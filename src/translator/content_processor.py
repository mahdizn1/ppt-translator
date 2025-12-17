# src/translator/content_processor.py
"""
Content Processor for Slide Translator

Handles extraction and injection of text content for LLM translation.
Acts as a bridge between XML (structure) and JSON (LLM interface).

Key features:
- Uses lxml for proper namespace handling
- Extracts text with hierarchy information (title, header, bullets, sub-bullets)
- Preserves formatting attributes during injection
- Maintains paragraph/run structure for complex text
"""

import json
import os
from typing import Dict, List, Optional, Any
from dataclasses import dataclass
from lxml import etree


# ============================================================================
# NAMESPACE CONFIGURATION
# ============================================================================
NAMESPACES: Dict[str, str] = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

NS = NAMESPACES


def qn(tag: str) -> str:
    """Convert prefixed tag to Clark notation."""
    if ":" not in tag:
        return tag
    prefix, local = tag.split(":", 1)
    if prefix not in NS:
        raise ValueError(f"Unknown namespace prefix: {prefix}")
    return f"{{{NS[prefix]}}}{local}"


# ============================================================================
# DATA STRUCTURES
# ============================================================================
@dataclass
class TextElement:
    """Represents a text element extracted from a slide."""
    id: str                          # Shape ID (for mapping back)
    name: str                        # Shape name (for debugging)
    role: str                        # title, subtitle, header, content
    level: int                       # Hierarchy level (0=title, 1=main point, 2=sub-point)
    y_position: int                  # Y coordinate (for ordering)
    paragraphs: List[Dict[str, Any]]  # List of {text, level, is_bold} dicts
    original_text: str               # Combined text for LLM


@dataclass
class SlideContent:
    """Complete content structure for a slide."""
    slide_context: str
    elements: List[TextElement]


# ============================================================================
# CONTENT PROCESSOR
# ============================================================================
class ContentProcessor:
    """
    Extracts and injects text content for slide translation.

    Workflow:
    1. extract_content_for_llm() -> JSON structure for LLM translation
    2. LLM translates the JSON
    3. inject_translated_content() -> Updates XML with translated text
    """

    # Placeholder types that indicate title/subtitle
    TITLE_TYPES = {"title", "ctrTitle"}
    SUBTITLE_TYPES = {"subTitle"}
    BODY_TYPES = {"body", "obj", "dt", "ftr", "sldNum"}

    def __init__(self, verbose: bool = True):
        self.verbose = verbose
        self.parser = etree.XMLParser(remove_blank_text=False)

    # ========================================================================
    # EXTRACTION
    # ========================================================================
    def extract_content_for_llm(self, slide_xml_path: str) -> Dict[str, Any]:
        """
        Extract text content from a slide XML file.

        Returns a JSON-serializable structure suitable for LLM translation:
        {
            "slide_context": "Consulting Slide",
            "elements": [
                {
                    "id": "5",
                    "role": "title",
                    "text": "Strategic Framework",
                    "paragraphs": [
                        {"text": "Strategic Framework", "level": 0, "is_bold": true}
                    ]
                },
                ...
            ]
        }
        """
        tree = etree.parse(slide_xml_path, self.parser)
        root = tree.getroot()

        extracted_elements: List[TextElement] = []

        # Find all shapes with text bodies
        shapes = root.findall(".//p:sp", NS)

        for shape in shapes:
            element = self._extract_shape_content(shape)
            if element and element.original_text.strip():
                extracted_elements.append(element)

        # Sort by Y position (top to bottom reading order)
        extracted_elements.sort(key=lambda e: e.y_position)

        # Build output JSON
        output = {
            "slide_context": "Consulting Slide - Professional business presentation",
            "elements": []
        }

        for element in extracted_elements:
            output["elements"].append({
                "id": element.id,
                "role": element.role,
                "text": element.original_text,
                "paragraphs": element.paragraphs
            })

        if self.verbose:
            print(f"[ContentProcessor] Extracted {len(extracted_elements)} text elements")
            for elem in extracted_elements:
                preview = elem.original_text[:50].replace('\n', ' ')
                print(f"  - [{elem.role}] {preview}...")

        return output

    def _extract_shape_content(self, shape: etree._Element) -> Optional[TextElement]:
        """Extract content from a single shape."""
        # Get shape ID and name
        nv_sp_pr = shape.find("p:nvSpPr", NS)
        if nv_sp_pr is None:
            return None

        c_nv_pr = nv_sp_pr.find("p:cNvPr", NS)
        if c_nv_pr is None:
            return None

        shape_id = c_nv_pr.get("id", "")
        shape_name = c_nv_pr.get("name", "")

        # Get text body
        tx_body = shape.find("p:txBody", NS)
        if tx_body is None:
            return None

        # Determine role from placeholder type
        role = self._determine_role(nv_sp_pr)

        # Get Y position for sorting
        y_pos = self._get_y_position(shape)

        # Extract paragraphs with hierarchy info
        paragraphs = self._extract_paragraphs(tx_body)

        if not paragraphs:
            return None

        # Combine text for LLM (preserving newlines between paragraphs)
        original_text = "\n".join(p["text"] for p in paragraphs if p["text"].strip())

        # Determine overall level
        level = 0 if role == "title" else (1 if role == "subtitle" else 2)

        return TextElement(
            id=shape_id,
            name=shape_name,
            role=role,
            level=level,
            y_position=y_pos,
            paragraphs=paragraphs,
            original_text=original_text
        )

    def _determine_role(self, nv_sp_pr: etree._Element) -> str:
        """Determine the semantic role of a shape."""
        # Check for placeholder type
        nv_pr = nv_sp_pr.find("p:nvPr", NS)
        if nv_pr is not None:
            ph = nv_pr.find("p:ph", NS)
            if ph is not None:
                ph_type = ph.get("type", "")
                if ph_type in self.TITLE_TYPES:
                    return "title"
                elif ph_type in self.SUBTITLE_TYPES:
                    return "subtitle"
                elif ph_type in self.BODY_TYPES:
                    return "body"

        return "content"

    def _get_y_position(self, shape: etree._Element) -> int:
        """Get the Y coordinate of a shape for sorting."""
        sp_pr = shape.find("p:spPr", NS)
        if sp_pr is None:
            return 999999

        xfrm = sp_pr.find("a:xfrm", NS)
        if xfrm is None:
            return 999999

        off = xfrm.find("a:off", NS)
        if off is None:
            return 999999

        return int(off.get("y", "999999"))

    def _extract_paragraphs(self, tx_body: etree._Element) -> List[Dict[str, Any]]:
        """Extract all paragraphs with formatting info."""
        paragraphs = []

        for p in tx_body.findall("a:p", NS):
            para_info = self._extract_paragraph(p)
            if para_info:
                paragraphs.append(para_info)

        return paragraphs

    def _extract_paragraph(self, p: etree._Element) -> Optional[Dict[str, Any]]:
        """Extract a single paragraph's content and formatting."""
        # Get paragraph properties
        p_pr = p.find("a:pPr", NS)
        level = 0
        if p_pr is not None:
            level = int(p_pr.get("lvl", "0"))

        # Extract text from all runs
        text_parts = []
        is_bold = False

        for r in p.findall("a:r", NS):
            # Check run properties for bold
            r_pr = r.find("a:rPr", NS)
            if r_pr is not None:
                b = r_pr.get("b")
                if b == "1" or b == "true":
                    is_bold = True

            # Get text content
            t = r.find("a:t", NS)
            if t is not None and t.text:
                text_parts.append(t.text)

        # Also check for text fields (a:fld)
        for fld in p.findall("a:fld", NS):
            t = fld.find("a:t", NS)
            if t is not None and t.text:
                text_parts.append(t.text)

        text = "".join(text_parts)

        if not text.strip():
            return None

        return {
            "text": text,
            "level": level,
            "is_bold": is_bold
        }

    # ========================================================================
    # INJECTION
    # ========================================================================
    def inject_translated_content(
        self,
        slide_xml_path: str,
        translated_json: Dict[str, Any],
        output_path: str
    ) -> None:
        """
        Inject translated text back into the slide XML.

        IMPORTANT: This should be called AFTER the Visual Engine has processed
        the slide, so RTL flags are already set. This method preserves those flags.

        Args:
            slide_xml_path: Path to the (already RTL-transformed) slide XML
            translated_json: JSON with translated elements
            output_path: Path to save the final XML
        """
        tree = etree.parse(slide_xml_path, self.parser)
        root = tree.getroot()

        # Build translation lookup map
        translation_map = {
            elem["id"]: elem
            for elem in translated_json.get("elements", [])
        }

        # Find and update each shape
        shapes = root.findall(".//p:sp", NS)
        updated_count = 0

        for shape in shapes:
            c_nv_pr = shape.find(".//p:cNvPr", NS)
            if c_nv_pr is None:
                continue

            shape_id = c_nv_pr.get("id", "")

            if shape_id in translation_map:
                translated = translation_map[shape_id]
                self._update_shape_text(shape, translated)
                updated_count += 1

        # Save output
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        tree.write(output_path, encoding="UTF-8", xml_declaration=True)

        if self.verbose:
            print(f"[ContentProcessor] Updated {updated_count} shapes")
            print(f"[ContentProcessor] Saved to: {output_path}")

    def _update_shape_text(
        self,
        shape: etree._Element,
        translated: Dict[str, Any]
    ) -> None:
        """
        Update text content of a shape while preserving structure.

        Strategy:
        - If translated has 'paragraphs' array, map each translated paragraph
          to the corresponding original paragraph
        - If only 'text' is provided, split by newlines and create paragraphs
        - Preserve existing paragraph properties (alignment, RTL, etc.)
        """
        tx_body = shape.find("p:txBody", NS)
        if tx_body is None:
            return

        # Get translated paragraphs or split text by newlines
        if "paragraphs" in translated and translated["paragraphs"]:
            new_paragraphs = translated["paragraphs"]
        else:
            # Split text by newlines
            text = translated.get("text", "")
            new_paragraphs = [{"text": line} for line in text.split("\n")]

        # Get existing paragraphs
        existing_paras = tx_body.findall("a:p", NS)

        # Strategy: Update existing paragraphs where possible, add new ones if needed
        for i, new_para in enumerate(new_paragraphs):
            new_text = new_para.get("text", "")

            if i < len(existing_paras):
                # Update existing paragraph
                self._update_paragraph_text(existing_paras[i], new_text)
            else:
                # Create new paragraph
                self._create_paragraph(tx_body, new_text)

        # Remove extra paragraphs if we have fewer translations
        while len(existing_paras) > len(new_paragraphs):
            extra = existing_paras.pop()
            tx_body.remove(extra)

    def _update_paragraph_text(
        self,
        paragraph: etree._Element,
        new_text: str
    ) -> None:
        """
        Update text in a paragraph while preserving properties.

        Preserves:
        - Paragraph properties (pPr) - alignment, RTL, level
        - First run's formatting (rPr) - font, size, color
        """
        # Find existing runs
        runs = paragraph.findall("a:r", NS)

        if runs:
            # Update first run's text, remove others
            first_run = runs[0]
            t = first_run.find("a:t", NS)
            if t is not None:
                t.text = new_text
            else:
                # Create text element
                t = etree.SubElement(first_run, qn("a:t"))
                t.text = new_text

            # Ensure run has Arabic language set
            r_pr = first_run.find("a:rPr", NS)
            if r_pr is not None:
                r_pr.set("lang", "ar-SA")

            # Remove extra runs
            for run in runs[1:]:
                paragraph.remove(run)
        else:
            # No runs exist, create one
            self._create_run(paragraph, new_text)

    def _create_paragraph(
        self,
        tx_body: etree._Element,
        text: str
    ) -> etree._Element:
        """Create a new paragraph with RTL properties."""
        p = etree.SubElement(tx_body, qn("a:p"))

        # Add paragraph properties with RTL
        p_pr = etree.SubElement(p, qn("a:pPr"))
        p_pr.set("rtl", "1")
        p_pr.set("algn", "r")

        # Add run with text
        self._create_run(p, text)

        return p

    def _create_run(
        self,
        paragraph: etree._Element,
        text: str
    ) -> etree._Element:
        """Create a new text run with Arabic language."""
        r = etree.SubElement(paragraph, qn("a:r"))

        # Run properties
        r_pr = etree.SubElement(r, qn("a:rPr"))
        r_pr.set("lang", "ar-SA")
        r_pr.set("dirty", "0")

        # Text content
        t = etree.SubElement(r, qn("a:t"))
        t.text = text

        return r

    # ========================================================================
    # UTILITIES
    # ========================================================================
    def save_json(self, content: Dict[str, Any], output_path: str) -> None:
        """Save extracted content to a JSON file."""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(content, f, ensure_ascii=False, indent=2)
        if self.verbose:
            print(f"[ContentProcessor] JSON saved to: {output_path}")

    def load_json(self, input_path: str) -> Dict[str, Any]:
        """Load content from a JSON file."""
        with open(input_path, 'r', encoding='utf-8') as f:
            return json.load(f)


# ============================================================================
# MAIN EXECUTION
# ============================================================================
if __name__ == "__main__":
    import glob
    import datetime

    processor = ContentProcessor(verbose=True)

    # Timestamp for output files
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # Find the latest source XML file (pattern: slide*_structure*.xml or slide*.xml)
    source_patterns = [
        "output_xmls/slide*_structure*.xml",
        "output_xmls/slide*.xml",
        "slide*_structure*.xml",
        "slide*.xml",
    ]

    source_files = []
    for pattern in source_patterns:
        source_files.extend(glob.glob(pattern))

    if not source_files:
        print("ERROR: No slide XML files found. Run the extractor first.")
        print(f"Searched patterns: {source_patterns}")
        exit(1)

    # Pick the latest file by modification time
    SOURCE_XML = max(source_files, key=os.path.getmtime)
    OUTPUT_JSON = f"output_xmls/slide_content_{timestamp}.json"

    print(f"\n{'='*50}")
    print("CONTENT EXTRACTION TEST")
    print(f"{'='*50}")
    print(f"Using source: {SOURCE_XML}")
    print(f"Output JSON:  {OUTPUT_JSON}")

    # Ensure output directory exists
    os.makedirs("output_xmls", exist_ok=True)

    # 1. Extract content
    content = processor.extract_content_for_llm(SOURCE_XML)
    processor.save_json(content, OUTPUT_JSON)

    print("\nExtracted content structure:")
    print(json.dumps(content, ensure_ascii=False, indent=2)[:1000] + "...")

    # 2. Simulate translation (for testing)
    print(f"\n{'='*50}")
    print("SIMULATED TRANSLATION")
    print(f"{'='*50}")

    translated = content.copy()
    for elem in translated["elements"]:
        # Mock Arabic translation (in real use, this comes from LLM)
        elem["text"] = f"[AR] {elem['text']}"
        if "paragraphs" in elem:
            for para in elem["paragraphs"]:
                para["text"] = f"[AR] {para['text']}"

    # 3. Test injection (if RTL XML exists)
    # Find most recent RTL file
    rtl_patterns = [
        "output_xmls/slide*_RTL*.xml",
        "output_xmls/*_rtl*.xml",
    ]

    rtl_files = []
    for pattern in rtl_patterns:
        rtl_files.extend(glob.glob(pattern))

    if rtl_files:
        RTL_XML = max(rtl_files, key=os.path.getmtime)
        FINAL_XML = f"output_xmls/slide_Final_{timestamp}.xml"

        print(f"\nUsing RTL file: {RTL_XML}")
        print(f"Output Final:   {FINAL_XML}")

        processor.inject_translated_content(RTL_XML, translated, FINAL_XML)
        print(f"Final output: {FINAL_XML}")
    else:
        print(f"\nNo RTL XML found. Run the Visual Engine first.")
        print(f"Searched patterns: {rtl_patterns}")
        print("Skipping injection test.")
