# src/translator/text_translator.py
"""
LLM-powered Translation Service for Consulting Slides

Uses OpenAI GPT-4o with structured outputs to translate slide content
from English to Arabic while preserving:
- JSON structure with element IDs for position mapping
- Consulting-style professional language
- Hierarchical context (title vs bullets vs sub-points)
- Formatting attributes (bold, alignment, bullet markers)

Requires: pip install openai pydantic python-dotenv
"""
from __future__ import annotations

import json
import os
from typing import Any, Optional
from enum import Enum

from pydantic import BaseModel, Field, field_validator
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


# ============================================================================
# PYDANTIC MODELS FOR STRUCTURED OUTPUT
# ============================================================================
class TextRole(str, Enum):
    """Semantic role of text element on slide."""
    TITLE = "title"
    SUBTITLE = "subtitle"
    HEADER = "header"
    BODY = "body"
    CONTENT = "content"
    FOOTER = "footer"


class ParagraphAlignment(str, Enum):
    """Text alignment options."""
    LEFT = "l"
    CENTER = "ctr"
    RIGHT = "r"
    JUSTIFY = "just"


class TranslatedParagraph(BaseModel):
    """A single paragraph within a text element."""
    text: str = Field(description="The translated Arabic text for this paragraph")
    lvl: int = Field(default=0, ge=0, le=5, description="Bullet indentation level (0=main, 1+=nested)")
    bold: bool = Field(default=False, description="Whether the text is bold")
    alignment: str = Field(default="r", description="Text alignment (r=right for RTL)")
    bullet: bool = Field(default=False, description="Whether this paragraph has a bullet marker")


class BoundingBox(BaseModel):
    """Position and dimensions of an element (in EMUs)."""
    x: int = Field(description="X coordinate (left edge)")
    y: int = Field(description="Y coordinate (top edge)")
    width: int = Field(description="Width of element")
    height: int = Field(description="Height of element")


class TranslatedElement(BaseModel):
    """A single text element on the slide."""
    id: str = Field(description="Unique shape ID - MUST be preserved exactly for position mapping")
    role: str = Field(description="Semantic role: title, subtitle, header, body, content")
    name: str = Field(description="Shape name from PowerPoint")
    bbox: BoundingBox = Field(description="Bounding box coordinates")
    paragraphs: list[TranslatedParagraph] = Field(description="List of translated paragraphs")

    @field_validator('id')
    @classmethod
    def id_must_not_be_empty(cls, v: str) -> str:
        if not v.strip():
            raise ValueError('Element ID cannot be empty')
        return v


class TranslatedSlide(BaseModel):
    """Complete translated slide content."""
    slide_context: str = Field(description="Brief Arabic description of the slide context")
    elements: list[TranslatedElement] = Field(description="List of translated text elements")


# ============================================================================
# CONSULTING-STYLE TRANSLATION PROMPT
# ============================================================================
SYSTEM_PROMPT = """You are an expert translator specializing in professional consulting presentations for McKinsey, BCG, and Bain-style strategy decks.

## YOUR TASK
Translate slide content from English to Modern Standard Arabic (فصحى) while maintaining the authoritative, strategic tone of top-tier consulting firms.

## CONTEXT-AWARE TRANSLATION (CRITICAL)

You are NOT translating isolated text fragments. Each element exists within a slide context:

1. **UNDERSTAND THE SLIDE FIRST**: Read the "slide_context" field to understand what this slide is about before translating any element.

2. **ELEMENT RELATIONSHIPS**: Elements on the same slide are related. A title introduces the topic, bullets expand on it. Translate with this relationship in mind.

3. **WORD PLACEMENT MATTERS**: The same English word should be translated differently based on where it appears:
   - "Overview" as a title → "نظرة شاملة" (comprehensive view)
   - "Overview" as a nav item → "نظرة عامة" (general overview)
   - "Impact" as a title → "الأثر الاستراتيجي" (strategic impact)
   - "Impact" in a bullet → "التأثير" (the effect)

4. **VISUAL HIERARCHY = TRANSLATION TONE**:
   - Elements at the TOP of slide (low Y coordinate) are usually more prominent
   - LARGER elements (bigger bbox) deserve more impactful translation
   - Title-positioned text should sound authoritative
   - Footer/nav text should be concise

## CRITICAL RULES FOR JSON STRUCTURE

1. **PRESERVE ALL IDs EXACTLY**: Each element has an "id" field that maps translated text back to the correct shape position. NEVER change, omit, or reorder IDs.

2. **MAINTAIN ELEMENT ORDER**: The elements array order reflects the visual hierarchy on the slide. Do not reorder elements.

3. **KEEP ALL METADATA**: Preserve bbox, role, name, lvl, bold, alignment, bullet values exactly as received.

## TRANSLATION HIERARCHY BY ROLE

The "role" field tells you the semantic importance of each element:

| Role | Translation Style | Example |
|------|-------------------|---------|
| title | Bold, commanding headlines. Use impactful verbs. Short and punchy. | "التحول الاستراتيجي نحو التميز" |
| subtitle | Supporting context, explanatory. Provides additional detail. | "نظرة شاملة على محفظة الأعمال" |
| header | Section headers, clear and direct. Labels for content below. | "النتائج الرئيسية" |
| body/content | Bullet points, parallel structure, action-oriented. The meat of the slide. | "تحسين الكفاءة التشغيلية بنسبة 25%" |
| footer | Navigation, page numbers, disclaimers. Keep very concise. | "التحليل" |

## BULLET LEVEL AWARENESS

The "lvl" field (0-5) indicates bullet nesting - translate accordingly:
- lvl=0: **Main points** - Most important, decisive language, can stand alone
- lvl=1: **Supporting details** - Explanatory, elaborates on lvl=0
- lvl=2+: **Sub-bullets** - Specific data points, examples, evidence

Example: If lvl=0 says "Improve efficiency" and lvl=1 says "Reduce costs by 25%", the lvl=1 translation should feel like it supports/explains the lvl=0 point.

## ARABIC CONSULTING LANGUAGE GUIDELINES

1. **Tone**: Authoritative, strategic, forward-looking
   - Use active voice: "نُعزّز" (we enhance) not "يتم تعزيز" (is enhanced)
   - Be decisive: "يجب" (must) over "من الممكن" (it's possible)

2. **Vocabulary**: Use established Arabic business terminology
   - Strategy = الاستراتيجية (not الخطة)
   - ROI = العائد على الاستثمار (keep "ROI" in parentheses first time)
   - KPIs = مؤشرات الأداء الرئيسية

3. **Formatting Preservation**:
   - Keep English acronyms: ROI, KPI, EBITDA, M&A
   - Keep brand names in English: McKinsey, Statkraft, BLOOM
   - Preserve numbers and percentages: 25%, $1.5M
   - Keep currency symbols and units

4. **Parallel Structure**: Bullets at the same level should have consistent grammatical structure
   - "زيادة الإيرادات", "تقليص التكاليف", "تحسين الكفاءة" (all verbal nouns)

## OUTPUT FORMAT
Return ONLY valid JSON matching the TranslatedSlide schema. No markdown, no explanations, no code blocks.

## EXAMPLE

Input:
{
  "slide_context": "Strategy overview",
  "elements": [
    {"id": "5", "role": "title", "name": "Title 1", "bbox": {"x": 0, "y": 0, "width": 9144000, "height": 1000000}, "paragraphs": [{"text": "Digital Transformation Roadmap", "lvl": 0, "bold": true, "alignment": "ctr", "bullet": false}]},
    {"id": "8", "role": "content", "name": "Content", "bbox": {"x": 500000, "y": 1500000, "width": 8000000, "height": 3000000}, "paragraphs": [{"text": "Increase operational efficiency by 25%", "lvl": 0, "bold": false, "alignment": "l", "bullet": true}]}
  ]
}

Output:
{
  "slide_context": "نظرة عامة على الاستراتيجية",
  "elements": [
    {"id": "5", "role": "title", "name": "Title 1", "bbox": {"x": 0, "y": 0, "width": 9144000, "height": 1000000}, "paragraphs": [{"text": "خارطة طريق التحول الرقمي", "lvl": 0, "bold": true, "alignment": "ctr", "bullet": false}]},
    {"id": "8", "role": "content", "name": "Content", "bbox": {"x": 500000, "y": 1500000, "width": 8000000, "height": 3000000}, "paragraphs": [{"text": "رفع الكفاءة التشغيلية بنسبة 25%", "lvl": 0, "bold": false, "alignment": "r", "bullet": true}]}
  ]
}

NOTICE:
- ID "5" and "8" are preserved exactly
- bbox values unchanged
- alignment changed from "l" to "r" for RTL
- Bold and bullet preserved
- Numbers preserved (25%)
- Consulting vocabulary used (الكفاءة التشغيلية)"""


USER_PROMPT_TEMPLATE = """Translate the following consulting slide content from English to Arabic.

IMPORTANT REMINDERS:
1. Preserve ALL "id" values exactly - they are critical for mapping text to shapes
2. Use professional Arabic consulting terminology
3. Keep brand names, acronyms, numbers, and percentages in their original form
4. Return ONLY valid JSON

Slide Content:
{json_content}

Translate now (JSON only):"""


# ============================================================================
# TRANSLATION SERVICE
# ============================================================================
class TranslationError(Exception):
    """Raised when translation fails."""
    pass


class TextTranslator:
    """
    OpenAI-powered translation service for consulting slides.

    Uses GPT-4o with structured outputs for reliable JSON generation.
    """

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: str = "gpt-5-mini",
        temperature: float = 0.2,
    ):
        """
        Initialize the translator.

        Args:
            api_key: OpenAI API key. If not provided, reads from OPENAI_API_KEY env var.
            model: Model to use (default: gpt-4o)
            temperature: Sampling temperature (lower = more consistent)
        """
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY")
        if not self.api_key:
            raise ValueError(
                "OPENAI_API_KEY is required. Set it in .env file or pass api_key parameter."
            )

        self.model = model
        self.temperature = temperature

        # Lazy import to avoid dependency issues if not using OpenAI
        try:
            from openai import OpenAI
            self.client = OpenAI(api_key=self.api_key)
        except ImportError:
            raise ImportError(
                "OpenAI library required. Install with: pip install openai"
            )

    def translate(self, content: dict[str, Any]) -> TranslatedSlide:
        """
        Translate slide content using OpenAI.

        Args:
            content: Dictionary with slide_context and elements

        Returns:
            TranslatedSlide with Arabic content

        Raises:
            TranslationError: If translation fails
        """
        # Prepare the user message with content
        user_message = USER_PROMPT_TEMPLATE.format(
            json_content=json.dumps(content, ensure_ascii=False, indent=2)
        )

        try:
            # Note: gpt-5-mini only supports default temperature (1.0)
            # Don't pass temperature parameter for this model
            api_params = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_message}
                ],
                "response_format": {"type": "json_object"}
            }

            # Only add temperature for models that support it (not gpt-5-mini)
            if self.model != "gpt-5-mini":
                api_params["temperature"] = self.temperature

            response = self.client.chat.completions.create(**api_params)

            # Extract response content
            response_text = response.choices[0].message.content
            if not response_text:
                raise TranslationError("Empty response from OpenAI")

            # Parse LLM response
            response_data = json.loads(response_text)

            # Merge LLM translations with original metadata
            # LLM only provides: id, role, paragraphs (with text)
            # We preserve from original: name, bbox, and any missing paragraph metadata
            merged_data = self._merge_with_original(content, response_data)

            # Validate merged data with Pydantic
            translated = TranslatedSlide.model_validate(merged_data)

            # Verify IDs were preserved
            self._verify_ids(content, translated)

            return translated

        except json.JSONDecodeError as e:
            raise TranslationError(f"Invalid JSON in response: {e}")
        except Exception as e:
            raise TranslationError(f"Translation failed: {e}")

    def translate_to_dict(self, content: dict[str, Any]) -> dict[str, Any]:
        """
        Translate and return as dictionary (for compatibility with existing code).

        Args:
            content: Dictionary with slide_context and elements

        Returns:
            Dictionary with translated content
        """
        translated = self.translate(content)
        return translated.model_dump()

    def _merge_with_original(self, original: dict, llm_response: dict) -> dict:
        """
        Merge LLM translations with original metadata.

        The LLM should only translate text - we preserve all structural metadata
        (name, bbox, lvl, bold, alignment, bullet) from the original content.
        """
        # Build lookup from original elements by ID
        original_by_id = {
            elem.get("id"): elem
            for elem in original.get("elements", [])
        }

        merged = {
            "slide_context": llm_response.get("slide_context", original.get("slide_context", "")),
            "elements": []
        }

        for llm_elem in llm_response.get("elements", []):
            elem_id = llm_elem.get("id")
            orig_elem = original_by_id.get(elem_id, {})

            # Start with original element and update with translations
            merged_elem = {
                "id": elem_id,
                "role": llm_elem.get("role", orig_elem.get("role", "content")),
                "name": orig_elem.get("name", ""),
                "bbox": orig_elem.get("bbox", {"x": 0, "y": 0, "width": 0, "height": 0}),
                "paragraphs": []
            }

            # Merge paragraphs - preserve metadata, update text
            orig_paragraphs = orig_elem.get("paragraphs", [])
            llm_paragraphs = llm_elem.get("paragraphs", [])

            for i, llm_para in enumerate(llm_paragraphs):
                orig_para = orig_paragraphs[i] if i < len(orig_paragraphs) else {}

                merged_para = {
                    "text": llm_para.get("text", ""),
                    "lvl": llm_para.get("lvl", orig_para.get("lvl", 0)),
                    "bold": llm_para.get("bold", orig_para.get("bold", False)),
                    "alignment": llm_para.get("alignment", orig_para.get("alignment", "r")),
                    "bullet": llm_para.get("bullet", orig_para.get("bullet", False))
                }
                merged_elem["paragraphs"].append(merged_para)

            merged["elements"].append(merged_elem)

        return merged

    def _verify_ids(self, original: dict, translated: TranslatedSlide) -> None:
        """Verify that all original IDs are preserved in translation."""
        original_ids = {elem.get("id") for elem in original.get("elements", [])}
        translated_ids = {elem.id for elem in translated.elements}

        missing_ids = original_ids - translated_ids
        if missing_ids:
            raise TranslationError(
                f"Translation lost element IDs: {missing_ids}. "
                "IDs must be preserved for correct text positioning."
            )


# ============================================================================
# CONVENIENCE FUNCTIONS
# ============================================================================
def load_content_json(path: str) -> dict[str, Any]:
    """Load slide content from a JSON file."""
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_translated_json(translated: TranslatedSlide | dict, path: str) -> None:
    """Save translated content to a JSON file."""
    if isinstance(translated, TranslatedSlide):
        data = translated.model_dump()
    else:
        data = translated

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ============================================================================
# CLI FOR TESTING
# ============================================================================
if __name__ == "__main__":
    import argparse
    import glob

    parser = argparse.ArgumentParser(description="Translate slide content JSON")
    parser.add_argument(
        "--input", "-i",
        help="Input JSON file (or auto-find latest in output_xmls/)"
    )
    parser.add_argument(
        "--output", "-o",
        help="Output JSON file (default: adds _translated suffix)"
    )
    parser.add_argument(
        "--model", "-m",
        default="gpt-4o",
        help="OpenAI model to use (default: gpt-4o)"
    )

    args = parser.parse_args()

    # Auto-find input if not specified
    if args.input:
        input_path = args.input
    else:
        patterns = [
            "output_xmls/slide*_content*.json",
            "work_*/slide*_content*.json",
        ]
        json_files = []
        for pattern in patterns:
            json_files.extend(glob.glob(pattern))

        if not json_files:
            print("ERROR: No content JSON files found.")
            print("Run the content processor first to extract text.")
            exit(1)

        input_path = max(json_files, key=os.path.getmtime)

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_translated{ext}"

    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")
    print(f"Model:  {args.model}")
    print()

    # Load content
    content = load_content_json(input_path)
    print(f"Loaded {len(content.get('elements', []))} elements")

    # Translate
    print("Translating...")
    translator = TextTranslator(model=args.model)
    translated = translator.translate(content)

    # Save
    save_translated_json(translated, output_path)
    print(f"\nTranslation complete!")
    print(f"Saved to: {output_path}")

    # Preview
    print("\n" + "="*50)
    print("PREVIEW (first 2 elements)")
    print("="*50)
    for elem in translated.elements[:2]:
        print(f"\n[{elem.role}] ID={elem.id}")
        for para in elem.paragraphs:
            print(f"  {para.text[:60]}...")
