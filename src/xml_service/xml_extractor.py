# src/xml_service/xml_extractor.py
"""
PPTX XML Extractor

Extracts XML files from PowerPoint (.pptx) archives for processing.

Key features:
- Extracts raw XML without prettifying (preserves PowerPoint compatibility)
- Supports extracting any internal file from the PPTX archive
- Option to prettify for human inspection (separate from processing)
"""

import zipfile
import os
from typing import Optional, List
from lxml import etree


class PPTXXMLExtractor:
    """
    Extracts internal XML files from a PowerPoint (.pptx) archive.

    PPTX files are ZIP archives containing XML files that define:
    - Presentation structure (presentation.xml)
    - Individual slides (ppt/slides/slide1.xml, etc.)
    - Themes, layouts, and masters
    - Relationships between files
    """

    def __init__(self, pptx_path: str):
        """
        Initialize the extractor with the path to a .pptx file.

        Args:
            pptx_path: Path to the source .pptx file.

        Raises:
            FileNotFoundError: If the file doesn't exist.
            zipfile.BadZipFile: If the file is not a valid ZIP/PPTX.
        """
        self.pptx_path = pptx_path

        if not os.path.exists(self.pptx_path):
            raise FileNotFoundError(f"File not found: '{self.pptx_path}'")

        # Validate it's a valid ZIP file
        if not zipfile.is_zipfile(self.pptx_path):
            raise zipfile.BadZipFile(f"Not a valid PPTX/ZIP file: '{self.pptx_path}'")

    def list_contents(self) -> List[str]:
        """List all files inside the PPTX archive."""
        with zipfile.ZipFile(self.pptx_path, 'r') as archive:
            return archive.namelist()

    def extract_raw(self, internal_path: str) -> bytes:
        """
        Extract raw bytes of a file from the PPTX archive.

        Args:
            internal_path: Path inside the archive (e.g., 'ppt/slides/slide1.xml')

        Returns:
            Raw bytes of the file content.

        Raises:
            KeyError: If the file doesn't exist in the archive.
        """
        with zipfile.ZipFile(self.pptx_path, 'r') as archive:
            if internal_path not in archive.namelist():
                raise KeyError(f"File not found in archive: '{internal_path}'")
            return archive.read(internal_path)

    def extract_slide_xml(
        self,
        slide_index: int,
        output_filename: str,
        prettify: bool = False
    ) -> None:
        """
        Extract a slide's XML to a file.

        IMPORTANT: For processing, use prettify=False to preserve exact XML structure.
        Use prettify=True only for human inspection.

        Args:
            slide_index: 1-based slide number (1 = first slide)
            output_filename: Path to save the extracted XML
            prettify: If True, format XML for readability (not for processing!)
        """
        internal_path = f"ppt/slides/slide{slide_index}.xml"

        try:
            xml_bytes = self.extract_raw(internal_path)

            if prettify:
                # Parse and re-serialize with indentation (for human reading only)
                tree = etree.fromstring(xml_bytes)
                xml_content = etree.tostring(
                    tree,
                    encoding='UTF-8',
                    pretty_print=True,
                    xml_declaration=True
                )
                with open(output_filename, 'wb') as f:
                    f.write(xml_content)
            else:
                # Write raw bytes exactly as they are in the PPTX
                # This preserves namespace prefixes and structure
                with open(output_filename, 'wb') as f:
                    f.write(xml_bytes)

            print(f"[Extractor] Slide {slide_index} -> {output_filename}")

        except KeyError as e:
            print(f"ERROR: {e}")
            raise

    def extract_presentation_xml(
        self,
        output_filename: str,
        prettify: bool = False
    ) -> None:
        """
        Extract the main presentation.xml file.

        This file contains:
        - Slide dimensions (sldSz)
        - Slide ordering
        - Default text styles
        """
        internal_path = "ppt/presentation.xml"

        try:
            xml_bytes = self.extract_raw(internal_path)

            if prettify:
                tree = etree.fromstring(xml_bytes)
                xml_content = etree.tostring(
                    tree,
                    encoding='UTF-8',
                    pretty_print=True,
                    xml_declaration=True
                )
                with open(output_filename, 'wb') as f:
                    f.write(xml_content)
            else:
                with open(output_filename, 'wb') as f:
                    f.write(xml_bytes)

            print(f"[Extractor] presentation.xml -> {output_filename}")

        except KeyError as e:
            print(f"ERROR: {e}")
            raise

    def extract_all_slides(
        self,
        output_dir: str,
        prettify: bool = False
    ) -> List[str]:
        """
        Extract all slides to a directory.

        Returns:
            List of output file paths.
        """
        os.makedirs(output_dir, exist_ok=True)

        # Find all slide files
        contents = self.list_contents()
        slide_files = sorted([
            f for f in contents
            if f.startswith('ppt/slides/slide') and f.endswith('.xml')
        ])

        output_paths = []
        for slide_file in slide_files:
            # Extract slide number from path
            filename = os.path.basename(slide_file)
            output_path = os.path.join(output_dir, filename)

            xml_bytes = self.extract_raw(slide_file)

            if prettify:
                tree = etree.fromstring(xml_bytes)
                xml_content = etree.tostring(
                    tree,
                    encoding='UTF-8',
                    pretty_print=True,
                    xml_declaration=True
                )
                with open(output_path, 'wb') as f:
                    f.write(xml_content)
            else:
                with open(output_path, 'wb') as f:
                    f.write(xml_bytes)

            output_paths.append(output_path)
            print(f"[Extractor] {slide_file} -> {output_path}")

        return output_paths

    def get_slide_count(self) -> int:
        """Count the number of slides in the presentation."""
        contents = self.list_contents()
        slide_files = [
            f for f in contents
            if f.startswith('ppt/slides/slide') and f.endswith('.xml')
        ]
        return len(slide_files)

    def get_slide_master_count(self) -> int:
        """Count the number of slide masters in the presentation."""
        contents = self.list_contents()
        master_files = [
            f for f in contents
            if f.startswith('ppt/slideMasters/slideMaster') and f.endswith('.xml')
        ]
        return len(master_files)

    def get_slide_layout_count(self) -> int:
        """Count the number of slide layouts in the presentation."""
        contents = self.list_contents()
        layout_files = [
            f for f in contents
            if f.startswith('ppt/slideLayouts/slideLayout') and f.endswith('.xml')
        ]
        return len(layout_files)

    def extract_slide_master_xml(
        self,
        master_index: int,
        output_filename: str,
        prettify: bool = False
    ) -> None:
        """
        Extract a slide master's XML to a file.

        Args:
            master_index: 1-based master number
            output_filename: Path to save the extracted XML
            prettify: If True, format XML for readability
        """
        internal_path = f"ppt/slideMasters/slideMaster{master_index}.xml"
        self._extract_xml_file(internal_path, output_filename, prettify, f"SlideMaster {master_index}")

    def extract_slide_layout_xml(
        self,
        layout_index: int,
        output_filename: str,
        prettify: bool = False
    ) -> None:
        """
        Extract a slide layout's XML to a file.

        Args:
            layout_index: 1-based layout number
            output_filename: Path to save the extracted XML
            prettify: If True, format XML for readability
        """
        internal_path = f"ppt/slideLayouts/slideLayout{layout_index}.xml"
        self._extract_xml_file(internal_path, output_filename, prettify, f"SlideLayout {layout_index}")

    def _extract_xml_file(
        self,
        internal_path: str,
        output_filename: str,
        prettify: bool,
        description: str
    ) -> None:
        """Generic XML file extraction helper."""
        try:
            xml_bytes = self.extract_raw(internal_path)

            if prettify:
                tree = etree.fromstring(xml_bytes)
                xml_content = etree.tostring(
                    tree,
                    encoding='UTF-8',
                    pretty_print=True,
                    xml_declaration=True
                )
                with open(output_filename, 'wb') as f:
                    f.write(xml_content)
            else:
                with open(output_filename, 'wb') as f:
                    f.write(xml_bytes)

            print(f"[Extractor] {description} -> {output_filename}")

        except KeyError as e:
            print(f"ERROR: {e}")
            raise

    def extract_all_masters(
        self,
        output_dir: str,
        prettify: bool = False
    ) -> List[str]:
        """
        Extract all slide masters to a directory.

        Returns:
            List of output file paths.
        """
        os.makedirs(output_dir, exist_ok=True)
        master_count = self.get_slide_master_count()
        output_paths = []

        for i in range(1, master_count + 1):
            output_path = os.path.join(output_dir, f"slideMaster{i}.xml")
            self.extract_slide_master_xml(i, output_path, prettify)
            output_paths.append(output_path)

        return output_paths

    def extract_all_layouts(
        self,
        output_dir: str,
        prettify: bool = False
    ) -> List[str]:
        """
        Extract all slide layouts to a directory.

        Returns:
            List of output file paths.
        """
        os.makedirs(output_dir, exist_ok=True)
        layout_count = self.get_slide_layout_count()
        output_paths = []

        for i in range(1, layout_count + 1):
            output_path = os.path.join(output_dir, f"slideLayout{i}.xml")
            self.extract_slide_layout_xml(i, output_path, prettify)
            output_paths.append(output_path)

        return output_paths

    def get_chart_count(self) -> int:
        """Count the number of charts in the presentation."""
        contents = self.list_contents()
        chart_files = [
            f for f in contents
            if f.startswith('ppt/charts/chart') and f.endswith('.xml') and '_rels' not in f
        ]
        return len(chart_files)

    def extract_chart_xml(
        self,
        chart_index: int,
        output_filename: str,
        prettify: bool = False
    ) -> None:
        """
        Extract a chart's XML to a file.

        Args:
            chart_index: 1-based chart number
            output_filename: Path to save the extracted XML
            prettify: If True, format XML for readability
        """
        internal_path = f"ppt/charts/chart{chart_index}.xml"
        self._extract_xml_file(internal_path, output_filename, prettify, f"Chart {chart_index}")

    def extract_all_charts(
        self,
        output_dir: str,
        prettify: bool = False
    ) -> List[str]:
        """
        Extract all charts to a directory.

        Returns:
            List of output file paths.
        """
        os.makedirs(output_dir, exist_ok=True)
        chart_count = self.get_chart_count()
        output_paths = []

        for i in range(1, chart_count + 1):
            output_path = os.path.join(output_dir, f"chart{i}.xml")
            self.extract_chart_xml(i, output_path, prettify)
            output_paths.append(output_path)

        return output_paths


# ============================================================================
# MAIN EXECUTION
# ============================================================================
if __name__ == "__main__":
    import glob
    import datetime

    OUTPUT_DIR = "./output_xmls"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # Find the latest PPTX file in common locations
    pptx_patterns = [
        "*.pptx",
        os.path.join(os.path.dirname(__file__), "..", "..", "*.pptx"),
        "C:\\Users\\user\\Downloads\\*.pptx",
    ]

    pptx_files = []
    for pattern in pptx_patterns:
        for f in glob.glob(pattern):
            # Exclude output files (those with timestamps or known outputs)
            basename = os.path.basename(f).lower()
            if not any(exclude in basename for exclude in ["output", "translated", "flipped"]):
                pptx_files.append(f)

    if not pptx_files:
        print("ERROR: No PPTX files found.")
        print("Searched patterns:")
        for pattern in pptx_patterns:
            print(f"  - {pattern}")
        exit(1)

    # Pick the latest PPTX file by modification time
    pptx_path = max(pptx_files, key=os.path.getmtime)

    # Timestamped output filenames
    OUTPUT_SLIDE_XML = os.path.join(OUTPUT_DIR, f"slide1_{timestamp}.xml")
    OUTPUT_PRES_XML = os.path.join(OUTPUT_DIR, f"presentation_{timestamp}.xml")
    OUTPUT_READABLE_XML = os.path.join(OUTPUT_DIR, f"slide1_readable_{timestamp}.xml")

    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"\n{'='*50}")
    print("PPTX XML EXTRACTION")
    print(f"{'='*50}")
    print(f"Input PPTX: {pptx_path}")
    print(f"Timestamp:  {timestamp}")
    print(f"Output dir: {OUTPUT_DIR}")

    try:
        extractor = PPTXXMLExtractor(pptx_path)

        print(f"\n[Extractor] Source: {pptx_path}")
        print(f"[Extractor] Slide count: {extractor.get_slide_count()}")

        # Extract slide 1 (raw, for processing)
        extractor.extract_slide_xml(
            slide_index=1,
            output_filename=OUTPUT_SLIDE_XML,
            prettify=False  # IMPORTANT: Keep raw for processing
        )

        # Extract presentation.xml (raw, for processing)
        extractor.extract_presentation_xml(
            output_filename=OUTPUT_PRES_XML,
            prettify=False  # IMPORTANT: Keep raw for processing
        )

        # Also create prettified versions for human inspection
        extractor.extract_slide_xml(
            slide_index=1,
            output_filename=OUTPUT_READABLE_XML,
            prettify=True
        )

        print("\n[Extractor] Complete!")
        print(f"  For processing: {OUTPUT_SLIDE_XML}")
        print(f"                  {OUTPUT_PRES_XML}")
        print(f"  For reading:    {OUTPUT_READABLE_XML}")

    except FileNotFoundError as e:
        print(f"ERROR: {e}")
    except zipfile.BadZipFile as e:
        print(f"ERROR: {e}")
