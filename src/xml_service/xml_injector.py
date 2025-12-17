import zipfile
import os
import shutil

#from translator.visual_engine import OUTPUT_FILENAME

class PPTXRebuilder:
    """
    A tool to inject modified XML files back into a PowerPoint (.pptx) archive.
    This allows you to verify that your XML edits result in a working presentation.
    """

    def __init__(self, original_pptx_path: str):
        """
        Args:
            original_pptx_path (str): The path to the valid source PPTX file.
        """
        self.original_pptx_path = original_pptx_path
        if not os.path.exists(self.original_pptx_path):
            raise FileNotFoundError(f"Original file '{self.original_pptx_path}' not found.")

    def inject_slide_xml(self, modified_xml_path: str, slide_index: int, output_pptx_path: str) -> None:
        """
        Creates a new PPTX file by copying the original and replacing a specific slide's XML.

        Args:
            modified_xml_path (str): Path to your edited XML file.
            slide_index (int): The 1-based index of the slide to replace (e.g., 1).
            output_pptx_path (str): The name of the new PPTX file to generate.
        """
        target_internal_file = f"ppt/slides/slide{slide_index}.xml"
        
        print(f"--- Starting Injection ---")
        print(f"Source: {self.original_pptx_path}")
        print(f"Injecting: {modified_xml_path} -> {target_internal_file}")
        
        try:
            # Open the original PPTX (Read Mode) and the New PPTX (Write Mode)
            with zipfile.ZipFile(self.original_pptx_path, 'r') as zin:
                with zipfile.ZipFile(output_pptx_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                    
                    # Iterate through every file in the original PPTX
                    for item in zin.infolist():
                        buffer = zin.read(item.filename)
                        
                        if item.filename == target_internal_file:
                            # FOUND IT: Don't write the original. Write the modified XML instead.
                            print(f"  > Replacing {item.filename}...")
                            
                            with open(modified_xml_path, 'rb') as f_xml:
                                modified_content = f_xml.read()
                                # We write the modified content to the new zip using the original filename
                                zout.writestr(item, modified_content)
                        else:
                            # Write the original file unchanged
                            zout.writestr(item, buffer)
                            
            print(f"Success! Created '{output_pptx_path}'")
            print("You can now open this file in PowerPoint to verify your XML edits.")

        except Exception as e:
            print(f"Error during rebuilding: {e}")
            # Clean up partial file if failed
            if os.path.exists(output_pptx_path):
                os.remove(output_pptx_path)

    def inject_presentation_xml(self, modified_xml_path: str, output_pptx_path: str) -> None:
        """
        Replaces the main presentation.xml (useful if you changed slide order/size).
        """
        target_internal = "ppt/presentation.xml"
        self._generic_inject(target_internal, modified_xml_path, output_pptx_path)

    def inject_slide_master_xml(self, modified_xml_path: str, master_index: int, output_pptx_path: str) -> None:
        """
        Replaces a slide master's XML.

        Args:
            modified_xml_path: Path to the modified master XML file.
            master_index: The 1-based index of the master to replace.
            output_pptx_path: The name of the new PPTX file to generate.
        """
        target_internal = f"ppt/slideMasters/slideMaster{master_index}.xml"
        self._generic_inject(target_internal, modified_xml_path, output_pptx_path)

    def inject_slide_layout_xml(self, modified_xml_path: str, layout_index: int, output_pptx_path: str) -> None:
        """
        Replaces a slide layout's XML.

        Args:
            modified_xml_path: Path to the modified layout XML file.
            layout_index: The 1-based index of the layout to replace.
            output_pptx_path: The name of the new PPTX file to generate.
        """
        target_internal = f"ppt/slideLayouts/slideLayout{layout_index}.xml"
        self._generic_inject(target_internal, modified_xml_path, output_pptx_path)

    def inject_multiple_files(self, replacements: dict, output_pptx_path: str) -> None:
        """
        Replace multiple files in one pass.

        Args:
            replacements: Dict mapping internal paths to local XML file paths
                         e.g., {"ppt/slides/slide1.xml": "/path/to/modified.xml"}
            output_pptx_path: The output PPTX file path
        """
        print(f"--- Starting Multi-File Injection ---")
        print(f"Source: {self.original_pptx_path}")
        print(f"Replacements: {len(replacements)} files")

        try:
            with zipfile.ZipFile(self.original_pptx_path, 'r') as zin:
                with zipfile.ZipFile(output_pptx_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:

                    for item in zin.infolist():
                        if item.filename in replacements:
                            # Replace with modified content
                            local_path = replacements[item.filename]
                            print(f"  > Replacing {item.filename}...")
                            with open(local_path, 'rb') as f:
                                zout.writestr(item, f.read())
                        else:
                            # Keep original
                            zout.writestr(item, zin.read(item.filename))

            print(f"Success! Created '{output_pptx_path}'")

        except Exception as e:
            print(f"Error during multi-file injection: {e}")
            if os.path.exists(output_pptx_path):
                os.remove(output_pptx_path)

    def _generic_inject(self, target_internal_filename: str, local_xml_path: str, output_path: str):
        """Helper method to replace any file inside the archive."""
        try:
            with zipfile.ZipFile(self.original_pptx_path, 'r') as zin:
                with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == target_internal_filename:
                            with open(local_xml_path, 'rb') as f:
                                zout.writestr(item, f.read())
                        else:
                            zout.writestr(item, zin.read(item.filename))
            print(f"Replaced '{target_internal_filename}' and saved to '{output_path}'")
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    import glob
    import datetime

    OUTPUT_DIR = "./output_pptx"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # Find the latest PPTX source file (not output files)
    pptx_patterns = [
        "*.pptx",
        os.path.join(os.path.dirname(__file__), "..", "..", "*.pptx"),
    ]

    pptx_files = []
    for pattern in pptx_patterns:
        for f in glob.glob(pattern):
            basename = os.path.basename(f).lower()
            if not any(exclude in basename for exclude in ["output", "translated", "flipped"]):
                pptx_files.append(f)

    if not pptx_files:
        print("ERROR: No source PPTX files found.")
        print("Searched patterns:")
        for pattern in pptx_patterns:
            print(f"  - {pattern}")
        exit(1)

    ORIGINAL_PPTX = max(pptx_files, key=os.path.getmtime)

    # Find the latest modified XML (Final first, then RTL)
    xml_patterns = [
        "output_xmls/slide*_Final*.xml",
        "output_xmls/slide*_RTL*.xml",
        "output_xmls/*_rtl*.xml",
    ]

    xml_files = []
    for pattern in xml_patterns:
        xml_files.extend(glob.glob(pattern))

    if not xml_files:
        print("ERROR: No modified XML files found.")
        print("Searched patterns:")
        for pattern in xml_patterns:
            print(f"  - {pattern}")
        print("\nRun the Visual Engine or Content Processor first.")
        exit(1)

    MODIFIED_XML = max(xml_files, key=os.path.getmtime)

    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    OUTPUT_PPTX = os.path.join(OUTPUT_DIR, f"Flipped_Output_{timestamp}.pptx")

    print(f"\n{'='*50}")
    print("PPTX REBUILDER")
    print(f"{'='*50}")
    print(f"Original PPTX: {ORIGINAL_PPTX}")
    print(f"Modified XML:  {MODIFIED_XML}")
    print(f"Output PPTX:   {OUTPUT_PPTX}")

    # Run the Rebuilder
    rebuilder = PPTXRebuilder(ORIGINAL_PPTX)
    rebuilder.inject_slide_xml(
        modified_xml_path=MODIFIED_XML,
        slide_index=1,
        output_pptx_path=OUTPUT_PPTX
    )

    print(f"\n[Rebuilder] Complete!")
    print(f"  Output: {OUTPUT_PPTX}")