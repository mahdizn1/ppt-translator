# main.py
"""
Slide Translator - Main Orchestrator

End-to-end pipeline for translating PowerPoint slides from English (LTR) to Arabic (RTL).

Pipeline:
1. Extract XML from PPTX
2. Extract text content for LLM
3. Translate via LLM (OpenAI/Anthropic)
4. Apply visual RTL transformation
5. Inject translated text
6. Rebuild PPTX

Usage:
    python main.py input.pptx output.pptx
    python main.py input.pptx output.pptx --slide 1
    python main.py input.pptx output.pptx --mock  # Skip LLM, use mock translation
"""

import os
import sys
import json
import argparse
import datetime
from pathlib import Path
from typing import Optional, Dict, Any

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from xml_service.xml_extractor import PPTXXMLExtractor
from xml_service.xml_injector import PPTXRebuilder
from translator.visual_engine import RTLVisualEngine
from translator.content_processor import ContentProcessor
from translator.chart_processor import ChartProcessor
from translator.text_translator import TextTranslator
from translator.llm_prompts import get_anthropic_prompt


# ============================================================================
# LLM TRANSLATION
# ============================================================================
def translate_with_openai(content_json: Dict[str, Any], api_key: Optional[str] = None) -> Dict[str, Any]:
    """
    Translate content using OpenAI API with structured output.

    Uses the TextTranslator service with Pydantic validation.
    """
    translator = TextTranslator(api_key=api_key, model="gpt-5-mini")
    return translator.translate_to_dict(content_json)


def translate_with_anthropic(content_json: Dict[str, Any], api_key: Optional[str] = None) -> Dict[str, Any]:
    """
    Translate content using Anthropic Claude API.

    Requires: pip install anthropic
    Set ANTHROPIC_API_KEY environment variable or pass api_key.
    """
    try:
        import anthropic
    except ImportError:
        raise ImportError("Please install anthropic: pip install anthropic")

    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY not set")

    client = anthropic.Anthropic(api_key=api_key)

    system_prompt, user_message = get_anthropic_prompt(
        json.dumps(content_json, ensure_ascii=False, indent=2)
    )

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}]
    )

    translated_text = response.content[0].text
    return json.loads(translated_text)


def translate_mock(content_json: Dict[str, Any]) -> Dict[str, Any]:
    """
    Mock translation for testing without LLM API.

    Uses sample Arabic consulting text to test proper rendering.
    The ID field is preserved to ensure correct text-to-shape mapping.
    """
    # Sample Arabic consulting translations for testing
    ARABIC_SAMPLES = {
        "title": "إطار التحول الاستراتيجي",
        "subtitle": "نظرة عامة على الخدمات الاستشارية",
        "header": "النتائج الرئيسية والتوصيات",
        "content": "تحسين الكفاءة التشغيلية بنسبة 25%",
        "body": "تقليص وقت الوصول إلى السوق بنسبة 40%",
    }

    translated = json.loads(json.dumps(content_json))  # Deep copy
    translated["slide_context"] = "شريحة استشارية - عرض تقديمي احترافي"

    for elem in translated.get("elements", []):
        # CRITICAL: Preserve the ID for correct mapping back to shapes
        element_id = elem.get("id")
        role = elem.get("role", "content")

        # Get appropriate sample translation based on role
        if role in ARABIC_SAMPLES:
            base_text = ARABIC_SAMPLES[role]
        else:
            base_text = ARABIC_SAMPLES["content"]

        # Create Arabic text with element ID for verification
        elem["text"] = f"{base_text} (#{element_id})"

        # Also update paragraphs if present - preserve structure
        if "paragraphs" in elem:
            for i, para in enumerate(elem["paragraphs"]):
                para_text = para.get("text", "")
                # Keep paragraph structure, add Arabic sample
                para["text"] = f"{base_text} - فقرة {i+1}"

    return translated


def translate_mock_chart(chart_json: Dict[str, Any]) -> Dict[str, Any]:
    """Mock translation for chart text."""
    translated = json.loads(json.dumps(chart_json))  # Deep copy

    if translated.get("chart_title"):
        translated["chart_title"] = "عنوان الرسم البياني"

    for i, series in enumerate(translated.get("series", [])):
        series["name"] = f"السلسلة {i+1}"

    for i, category in enumerate(translated.get("categories", [])):
        translated["categories"][i] = f"الفئة {i+1}"

    return translated


def translate_with_openai_chart(chart_json: Dict[str, Any], api_key: Optional[str] = None) -> Dict[str, Any]:
    """Translate chart text using OpenAI API."""
    # Create a simple prompt for chart translation
    translator = TextTranslator(api_key=api_key, model="gpt-5-mini")

    # Build a simple content structure for translation
    content = {
        "slide_context": "Chart - Professional business presentation",
        "elements": []
    }

    # Track element types for reconstruction
    element_map = []  # List of (type, original_id) tuples in order

    elem_id = 0
    if chart_json.get("chart_title"):
        content["elements"].append({
            "id": str(elem_id),
            "role": "title",
            "text": chart_json["chart_title"],
            "paragraphs": [{"text": chart_json["chart_title"], "level": 0, "is_bold": True}]
        })
        element_map.append(("title", None))
        elem_id += 1

    for series in chart_json.get("series", []):
        content["elements"].append({
            "id": str(elem_id),
            "role": "content",
            "text": series["name"],
            "paragraphs": [{"text": series["name"], "level": 0, "is_bold": False}]
        })
        element_map.append(("series", series["id"]))
        elem_id += 1

    for i, cat in enumerate(chart_json.get("categories", [])):
        content["elements"].append({
            "id": str(elem_id),
            "role": "content",
            "text": cat,
            "paragraphs": [{"text": cat, "level": 0, "is_bold": False}]
        })
        element_map.append(("category", i))
        elem_id += 1

    # Translate
    translated_content = translator.translate_to_dict(content)

    # Extract back to chart format using element_map for ordering
    result = {
        "chart_title": None,
        "series": [],
        "categories": []
    }

    translated_elements = translated_content.get("elements", [])

    for i, (elem_type, orig_id) in enumerate(element_map):
        if i >= len(translated_elements):
            break

        elem = translated_elements[i]
        translated_text = elem.get("text", "")

        # Also check paragraphs if text is empty
        if not translated_text and elem.get("paragraphs"):
            translated_text = elem["paragraphs"][0].get("text", "")

        if elem_type == "title":
            result["chart_title"] = translated_text
        elif elem_type == "series":
            result["series"].append({
                "id": orig_id,
                "name": translated_text
            })
        elif elem_type == "category":
            result["categories"].append(translated_text)

    return result


def translate_with_anthropic_chart(chart_json: Dict[str, Any], api_key: Optional[str] = None) -> Dict[str, Any]:
    """Translate chart text using Anthropic Claude API."""
    # For simplicity, use the same approach as OpenAI
    # In practice, you could create a specialized Anthropic prompt for charts
    return translate_with_openai_chart(chart_json, api_key)


# ============================================================================
# MAIN PIPELINE
# ============================================================================
class SlideTranslator:
    """
    Main orchestrator for the slide translation pipeline.
    """

    def __init__(
        self,
        input_pptx: str,
        output_pptx: str,
        work_dir: Optional[str] = None,
        verbose: bool = True
    ):
        self.input_pptx = input_pptx
        self.output_pptx = output_pptx
        self.verbose = verbose

        # Create working directory
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.work_dir = work_dir or f"./work_{timestamp}"
        os.makedirs(self.work_dir, exist_ok=True)

        # Initialize components
        self.extractor = PPTXXMLExtractor(input_pptx)
        self.content_processor = ContentProcessor(verbose=verbose)
        self.chart_processor = ChartProcessor(verbose=verbose)
        self.slide_count = self.extractor.get_slide_count()
        self.master_count = self.extractor.get_slide_master_count()
        self.layout_count = self.extractor.get_slide_layout_count()
        self.chart_count = self.extractor.get_chart_count()

        # Store paths for multi-slide processing
        self.pres_xml = os.path.join(self.work_dir, "presentation.xml")

        # Track transformed masters/layouts (transform once, reuse)
        self._masters_transformed = False
        self._transformed_masters = {}  # master_index -> transformed_xml_path
        self._transformed_layouts = {}  # layout_index -> transformed_xml_path
        self._transformed_charts = {}  # chart_index -> transformed_xml_path

        if verbose:
            print(f"\n{'='*60}")
            print("SLIDE TRANSLATOR INITIALIZED")
            print(f"{'='*60}")
            print(f"Input:  {input_pptx}")
            print(f"Output: {output_pptx}")
            print(f"Work:   {self.work_dir}")
            print(f"Slides: {self.slide_count}")
            print(f"Masters: {self.master_count}")
            print(f"Layouts: {self.layout_count}")
            print(f"Charts: {self.chart_count}")

    def _process_single_slide(
        self,
        slide_index: int,
        translator: str,
        api_key: Optional[str] = None
    ) -> str:
        """
        Process a single slide through the translation pipeline.

        Returns:
            Path to the final transformed XML for this slide.
        """
        if self.verbose:
            print(f"\n{'-'*50}")
            print(f"PROCESSING SLIDE {slide_index}")
            print(f"{'-'*50}")

        # File paths for this slide
        slide_xml = os.path.join(self.work_dir, f"slide{slide_index}.xml")
        content_json_path = os.path.join(self.work_dir, f"slide{slide_index}_content.json")
        translated_json_path = os.path.join(self.work_dir, f"slide{slide_index}_translated.json")
        rtl_xml = os.path.join(self.work_dir, f"slide{slide_index}_rtl.xml")
        final_xml = os.path.join(self.work_dir, f"slide{slide_index}_final.xml")

        # Step 1: Extract slide XML
        if self.verbose:
            print(f"  [1/5] Extracting slide XML...")
        self.extractor.extract_slide_xml(slide_index, slide_xml, prettify=False)

        # Step 2: Extract content for LLM
        if self.verbose:
            print(f"  [2/5] Extracting text content...")
        content_json = self.content_processor.extract_content_for_llm(slide_xml)
        self.content_processor.save_json(content_json, content_json_path)

        # Step 3: Translate
        if self.verbose:
            print(f"  [3/5] Translating ({translator})...")

        if translator == "openai":
            translated_json = translate_with_openai(content_json, api_key)
        elif translator == "anthropic":
            translated_json = translate_with_anthropic(content_json, api_key)
        elif translator == "mock":
            translated_json = translate_mock(content_json)
        else:
            raise ValueError(f"Unknown translator: {translator}")

        self.content_processor.save_json(translated_json, translated_json_path)

        # Step 4: Visual RTL transformation
        if self.verbose:
            print(f"  [4/5] Applying RTL visual transformation...")

        engine = RTLVisualEngine(
            presentation_xml_path=self.pres_xml,
            slide_xml_path=slide_xml,
            layout_flip_ratio=0.4,
            flip_connectors=True,
            verbose=False  # Reduce noise for multi-slide
        )
        engine.transform()
        engine.save(rtl_xml)

        # Step 5: Inject translated text
        if self.verbose:
            print(f"  [5/5] Injecting translated text...")

        self.content_processor.inject_translated_content(
            rtl_xml,
            translated_json,
            final_xml
        )

        return final_xml

    def _translate_all_charts(
        self,
        translator: str,
        api_key: Optional[str] = None
    ) -> None:
        """
        Extract and translate all charts in the presentation.

        This should be called once before processing slides.
        Charts are stored as separate XML files and referenced by slides.
        """
        if self.chart_count == 0:
            return

        if self.verbose:
            print(f"\n{'-'*50}")
            print(f"TRANSLATING {self.chart_count} CHARTS")
            print(f"{'-'*50}")

        for i in range(1, self.chart_count + 1):
            chart_xml = os.path.join(self.work_dir, f"chart{i}.xml")
            chart_content_json = os.path.join(self.work_dir, f"chart{i}_content.json")
            chart_translated_json = os.path.join(self.work_dir, f"chart{i}_translated.json")
            chart_final = os.path.join(self.work_dir, f"chart{i}_final.xml")

            if self.verbose:
                print(f"  Translating chart{i}...")

            # Extract chart XML
            self.extractor.extract_chart_xml(i, chart_xml, prettify=False)

            # Extract chart text
            chart_content = self.chart_processor.extract_chart_text(chart_xml)
            self.chart_processor.save_json(chart_content, chart_content_json)

            # Translate chart text
            if translator == "openai":
                translated_chart = translate_with_openai_chart(chart_content, api_key)
            elif translator == "anthropic":
                translated_chart = translate_with_anthropic_chart(chart_content, api_key)
            elif translator == "mock":
                translated_chart = translate_mock_chart(chart_content)
            else:
                raise ValueError(f"Unknown translator: {translator}")

            self.chart_processor.save_json(translated_chart, chart_translated_json)

            # Inject translated text into chart XML
            self.chart_processor.inject_chart_text(
                chart_xml,
                translated_chart,
                chart_final
            )

            self._transformed_charts[i] = chart_final

        if self.verbose:
            print(f"  Translated {self.chart_count} charts")

    def _transform_masters_and_layouts(
        self,
        translator: str = "mock",
        api_key: Optional[str] = None
    ) -> None:
        """
        Transform all slide masters and layouts for RTL and translate their text.

        This should be called once before processing slides.
        Masters and layouts contain elements like logos, navigation bars,
        and footers that appear on all slides.
        """
        if self._masters_transformed:
            return

        if self.verbose:
            print(f"\n{'-'*50}")
            print("TRANSFORMING SLIDE MASTERS AND LAYOUTS")
            print(f"{'-'*50}")

        # Transform all slide masters
        for i in range(1, self.master_count + 1):
            master_xml = os.path.join(self.work_dir, f"slideMaster{i}.xml")
            master_content_json = os.path.join(self.work_dir, f"slideMaster{i}_content.json")
            master_translated_json = os.path.join(self.work_dir, f"slideMaster{i}_translated.json")
            master_rtl = os.path.join(self.work_dir, f"slideMaster{i}_rtl.xml")
            master_final = os.path.join(self.work_dir, f"slideMaster{i}_final.xml")

            if self.verbose:
                print(f"  Transforming slideMaster{i}...")

            # Extract master XML
            self.extractor.extract_slide_master_xml(i, master_xml, prettify=False)

            # Extract text content
            master_content = self.content_processor.extract_content_for_llm(master_xml)

            # Only translate if there's text to translate
            if master_content.get("elements"):
                self.content_processor.save_json(master_content, master_content_json)

                # Translate text
                if translator == "openai":
                    translated_master = translate_with_openai(master_content, api_key)
                elif translator == "anthropic":
                    translated_master = translate_with_anthropic(master_content, api_key)
                elif translator == "mock":
                    translated_master = translate_mock(master_content)
                else:
                    raise ValueError(f"Unknown translator: {translator}")

                self.content_processor.save_json(translated_master, master_translated_json)

            # Apply RTL visual transformation
            engine = RTLVisualEngine(
                presentation_xml_path=self.pres_xml,
                slide_xml_path=master_xml,
                layout_flip_ratio=0.4,
                flip_connectors=True,
                verbose=False
            )
            engine.transform()
            engine.save(master_rtl)

            # Inject translated text into RTL-transformed XML
            if master_content.get("elements"):
                self.content_processor.inject_translated_content(
                    master_rtl,
                    translated_master,
                    master_final
                )
                self._transformed_masters[i] = master_final
            else:
                # No text to translate, use RTL XML as-is
                self._transformed_masters[i] = master_rtl

        # Transform all slide layouts
        for i in range(1, self.layout_count + 1):
            layout_xml = os.path.join(self.work_dir, f"slideLayout{i}.xml")
            layout_content_json = os.path.join(self.work_dir, f"slideLayout{i}_content.json")
            layout_translated_json = os.path.join(self.work_dir, f"slideLayout{i}_translated.json")
            layout_rtl = os.path.join(self.work_dir, f"slideLayout{i}_rtl.xml")
            layout_final = os.path.join(self.work_dir, f"slideLayout{i}_final.xml")

            if self.verbose:
                print(f"  Transforming slideLayout{i}...")

            # Extract layout XML
            self.extractor.extract_slide_layout_xml(i, layout_xml, prettify=False)

            # Extract text content
            layout_content = self.content_processor.extract_content_for_llm(layout_xml)

            # Only translate if there's text to translate
            if layout_content.get("elements"):
                self.content_processor.save_json(layout_content, layout_content_json)

                # Translate text
                if translator == "openai":
                    translated_layout = translate_with_openai(layout_content, api_key)
                elif translator == "anthropic":
                    translated_layout = translate_with_anthropic(layout_content, api_key)
                elif translator == "mock":
                    translated_layout = translate_mock(layout_content)
                else:
                    raise ValueError(f"Unknown translator: {translator}")

                self.content_processor.save_json(translated_layout, layout_translated_json)

            # Apply RTL visual transformation
            engine = RTLVisualEngine(
                presentation_xml_path=self.pres_xml,
                slide_xml_path=layout_xml,
                layout_flip_ratio=0.4,
                flip_connectors=True,
                verbose=False
            )
            engine.transform()
            engine.save(layout_rtl)

            # Inject translated text into RTL-transformed XML
            if layout_content.get("elements"):
                self.content_processor.inject_translated_content(
                    layout_rtl,
                    translated_layout,
                    layout_final
                )
                self._transformed_layouts[i] = layout_final
            else:
                # No text to translate, use RTL XML as-is
                self._transformed_layouts[i] = layout_rtl

        self._masters_transformed = True

        if self.verbose:
            print(f"  Transformed {self.master_count} masters and {self.layout_count} layouts")

    def translate_slides(
        self,
        slide_indices: list,
        translator: str = "mock",
        api_key: Optional[str] = None
    ) -> str:
        """
        Translate multiple slides.

        Args:
            slide_indices: List of 1-based slide numbers (e.g., [1, 2])
            translator: "openai", "anthropic", or "mock"
            api_key: Optional API key

        Returns:
            Path to the output PPTX file
        """
        if self.verbose:
            print(f"\n{'='*60}")
            print(f"TRANSLATING {len(slide_indices)} SLIDES: {slide_indices}")
            print(f"{'='*60}")

        # Extract presentation.xml once (shared across slides)
        self.extractor.extract_presentation_xml(self.pres_xml, prettify=False)

        # Transform all masters and layouts for RTL and translate their text (once)
        self._transform_masters_and_layouts(translator, api_key)

        # Translate all charts (once)
        self._translate_all_charts(translator, api_key)

        # Process each slide and collect the final XMLs
        final_xmls = {}
        for slide_index in slide_indices:
            if slide_index < 1 or slide_index > self.slide_count:
                print(f"WARNING: Slide {slide_index} out of range (1-{self.slide_count}), skipping")
                continue
            final_xml = self._process_single_slide(slide_index, translator, api_key)
            final_xmls[slide_index] = final_xml

        # Build replacements dict for multi-file injection
        replacements = {}

        # Add transformed slides
        for slide_index, xml_path in final_xmls.items():
            internal_path = f"ppt/slides/slide{slide_index}.xml"
            replacements[internal_path] = xml_path

        # Add transformed masters
        for master_index, xml_path in self._transformed_masters.items():
            internal_path = f"ppt/slideMasters/slideMaster{master_index}.xml"
            replacements[internal_path] = xml_path

        # Add transformed layouts
        for layout_index, xml_path in self._transformed_layouts.items():
            internal_path = f"ppt/slideLayouts/slideLayout{layout_index}.xml"
            replacements[internal_path] = xml_path

        # Add translated charts
        for chart_index, xml_path in self._transformed_charts.items():
            internal_path = f"ppt/charts/chart{chart_index}.xml"
            replacements[internal_path] = xml_path

        # Rebuild PPTX with all modified files in one pass
        if self.verbose:
            print(f"\n{'='*60}")
            print("REBUILDING PPTX")
            print(f"{'='*60}")
            print(f"  Slides: {len(final_xmls)}")
            print(f"  Masters: {len(self._transformed_masters)}")
            print(f"  Layouts: {len(self._transformed_layouts)}")
            print(f"  Charts: {len(self._transformed_charts)}")
            print(f"  Total replacements: {len(replacements)}")

        rebuilder = PPTXRebuilder(self.input_pptx)
        rebuilder.inject_multiple_files(replacements, self.output_pptx)

        if self.verbose:
            print(f"\n{'='*60}")
            print("TRANSLATION COMPLETE")
            print(f"{'='*60}")
            print(f"Output: {self.output_pptx}")
            print(f"Slides translated: {list(final_xmls.keys())}")

        return self.output_pptx

    def translate_slide(
        self,
        slide_index: int = 1,
        translator: str = "mock",
        api_key: Optional[str] = None
    ) -> str:
        """
        Translate a single slide (convenience wrapper).

        Args:
            slide_index: 1-based slide number
            translator: "openai", "anthropic", or "mock"
            api_key: Optional API key (or use environment variable)

        Returns:
            Path to the output PPTX file
        """
        return self.translate_slides([slide_index], translator, api_key)

    def translate_all(
        self,
        translator: str = "mock",
        api_key: Optional[str] = None
    ) -> str:
        """
        Translate all slides in the presentation.

        Returns:
            Path to the output PPTX file
        """
        all_indices = list(range(1, self.slide_count + 1))
        return self.translate_slides(all_indices, translator, api_key)


# ============================================================================
# CLI ENTRY POINT
# ============================================================================
def parse_slides_arg(slides_str: str, max_slides: int) -> list:
    """
    Parse slides argument into a list of slide indices.

    Supports:
    - "all" -> all slides
    - "1" -> [1]
    - "1,2" -> [1, 2]
    - "1-3" -> [1, 2, 3]
    - "1,3-5" -> [1, 3, 4, 5]
    """
    if slides_str.lower() == "all":
        return list(range(1, max_slides + 1))

    result = []
    parts = slides_str.split(",")

    for part in parts:
        part = part.strip()
        if "-" in part:
            # Range: "1-3"
            start, end = part.split("-", 1)
            start = int(start.strip())
            end = int(end.strip())
            result.extend(range(start, end + 1))
        else:
            # Single number
            result.append(int(part))

    return sorted(set(result))  # Remove duplicates and sort


def main():
    parser = argparse.ArgumentParser(
        description="Translate PowerPoint slides from English to Arabic with RTL layout"
    )
    parser.add_argument("input", help="Input PPTX file path")
    parser.add_argument("output", help="Output PPTX file path")
    parser.add_argument(
        "--slides", "-s",
        type=str,
        default="1",
        help="Slides to translate: '1', '1,2', '1-3', or 'all' (default: 1)"
    )
    parser.add_argument(
        "--translator", "-t",
        choices=["openai", "anthropic", "mock"],
        default="mock",
        help="Translation engine (default: mock)"
    )
    parser.add_argument(
        "--api-key", "-k",
        help="API key (or use OPENAI_API_KEY/ANTHROPIC_API_KEY env var)"
    )
    parser.add_argument(
        "--work-dir", "-w",
        help="Working directory for intermediate files"
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress verbose output"
    )

    args = parser.parse_args()

    # Validate input file
    if not os.path.exists(args.input):
        print(f"ERROR: Input file not found: {args.input}")
        sys.exit(1)

    try:
        translator = SlideTranslator(
            input_pptx=args.input,
            output_pptx=args.output,
            work_dir=args.work_dir,
            verbose=not args.quiet
        )

        # Parse slides argument
        slide_indices = parse_slides_arg(args.slides, translator.slide_count)

        if not slide_indices:
            print(f"ERROR: No valid slides specified")
            sys.exit(1)

        translator.translate_slides(
            slide_indices=slide_indices,
            translator=args.translator,
            api_key=args.api_key
        )

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
