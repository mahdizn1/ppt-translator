# src/translator/chart_processor.py
"""
Chart Processor for Slide Translator

Handles extraction and injection of text content from PowerPoint charts.
Charts are stored in separate XML files (ppt/charts/chartN.xml) and contain:
- Chart titles
- Series names (legend labels)
- Category names (axis labels)
- Data labels (if shown)

Key features:
- Uses lxml for proper namespace handling
- Extracts all text elements from chart XML
- Preserves formatting and structure during injection
"""

import json
import os
from typing import Dict, List, Optional, Any
from lxml import etree


# ============================================================================
# NAMESPACE CONFIGURATION
# ============================================================================
NAMESPACES: Dict[str, str] = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
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
# CHART PROCESSOR
# ============================================================================
class ChartProcessor:
    """
    Extracts and injects text content from chart XML files.

    Workflow:
    1. extract_chart_text() -> JSON structure with all chart text
    2. LLM translates the JSON
    3. inject_chart_text() -> Updates chart XML with translated text
    """

    def __init__(self, verbose: bool = True):
        self.verbose = verbose
        self.parser = etree.XMLParser(remove_blank_text=False)

    # ========================================================================
    # EXTRACTION
    # ========================================================================
    def extract_chart_text(self, chart_xml_path: str) -> Dict[str, Any]:
        """
        Extract all text content from a chart XML file.

        Returns a JSON-serializable structure:
        {
            "chart_title": "Sales Performance",
            "series": [
                {"id": "0", "name": "Q1 2024"},
                {"id": "1", "name": "Q2 2024"}
            ],
            "categories": ["Product A", "Product B", "Product C"]
        }
        """
        tree = etree.parse(chart_xml_path, self.parser)
        root = tree.getroot()

        result = {
            "chart_title": None,
            "series": [],
            "categories": []
        }

        # Extract chart title
        title_elem = root.find(".//c:chart/c:title/c:tx/c:rich/a:p/a:r/a:t", NS)
        if title_elem is not None and title_elem.text:
            result["chart_title"] = title_elem.text

        # Extract series names (legend labels)
        series_elements = root.findall(".//c:ser", NS)
        for idx, ser in enumerate(series_elements):
            ser_idx = ser.find("c:idx", NS)
            ser_order = ser.find("c:order", NS)

            # Get series ID
            series_id = ser_idx.get("val") if ser_idx is not None else str(idx)

            # Get series name
            ser_name_elem = ser.find(".//c:tx/c:strRef/c:strCache/c:pt/c:v", NS)
            if ser_name_elem is not None and ser_name_elem.text:
                result["series"].append({
                    "id": series_id,
                    "name": ser_name_elem.text
                })

        # Extract categories (X-axis labels)
        # Categories are usually in the first series' <c:cat> element
        first_series = root.find(".//c:ser", NS)
        if first_series is not None:
            cat_elements = first_series.findall(".//c:cat/c:strRef/c:strCache/c:pt", NS)
            for cat_pt in cat_elements:
                cat_v = cat_pt.find("c:v", NS)
                if cat_v is not None and cat_v.text:
                    result["categories"].append(cat_v.text)

        if self.verbose:
            print(f"[ChartProcessor] Extracted chart text:")
            print(f"  - Title: {result['chart_title']}")
            print(f"  - Series: {len(result['series'])}")
            print(f"  - Categories: {len(result['categories'])}")

        return result

    # ========================================================================
    # INJECTION
    # ========================================================================
    def inject_chart_text(
        self,
        chart_xml_path: str,
        translated_json: Dict[str, Any],
        output_path: str
    ) -> None:
        """
        Inject translated text back into the chart XML.

        Args:
            chart_xml_path: Path to the original chart XML
            translated_json: JSON with translated chart text
            output_path: Path to save the modified chart XML
        """
        tree = etree.parse(chart_xml_path, self.parser)
        root = tree.getroot()

        updated_count = 0

        # Update chart title
        if "chart_title" in translated_json and translated_json["chart_title"]:
            title_elem = root.find(".//c:chart/c:title/c:tx/c:rich/a:p/a:r/a:t", NS)
            if title_elem is not None:
                title_elem.text = translated_json["chart_title"]
                updated_count += 1

                # Update language attribute for title
                r_pr = root.find(".//c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr", NS)
                if r_pr is not None:
                    r_pr.set("lang", "ar-SA")

        # Update series names
        if "series" in translated_json:
            series_map = {s["id"]: s["name"] for s in translated_json["series"]}

            series_elements = root.findall(".//c:ser", NS)
            for ser in series_elements:
                ser_idx = ser.find("c:idx", NS)
                if ser_idx is not None:
                    series_id = ser_idx.get("val")
                    if series_id in series_map:
                        # Update series name
                        ser_name_elem = ser.find(".//c:tx/c:strRef/c:strCache/c:pt/c:v", NS)
                        if ser_name_elem is not None:
                            ser_name_elem.text = series_map[series_id]
                            updated_count += 1

        # Update categories
        if "categories" in translated_json and translated_json["categories"]:
            # Update categories in all series (they should all have the same categories)
            series_elements = root.findall(".//c:ser", NS)
            for ser in series_elements:
                cat_elements = ser.findall(".//c:cat/c:strRef/c:strCache/c:pt", NS)
                for i, cat_pt in enumerate(cat_elements):
                    if i < len(translated_json["categories"]):
                        cat_v = cat_pt.find("c:v", NS)
                        if cat_v is not None:
                            cat_v.text = translated_json["categories"][i]
                            # Only count once (not for each series)
                            if ser == series_elements[0]:
                                updated_count += 1

        # RTL Bar Chart Adjustment: Flip horizontal bar charts
        # For horizontal bar charts (barDir="bar"), reverse the value axis orientation
        # so bars grow from right to left instead of left to right
        bar_chart = root.find(".//c:barChart", NS)
        if bar_chart is not None:
            bar_dir = bar_chart.find("c:barDir", NS)
            # barDir="bar" means horizontal bars (as opposed to "col" for vertical columns)
            if bar_dir is not None and bar_dir.get("val") == "bar":
                # Find the value axis (the horizontal axis for bar charts)
                val_ax = root.find(".//c:valAx", NS)
                if val_ax is not None:
                    # Find or create the scaling element
                    scaling = val_ax.find("c:scaling", NS)
                    if scaling is None:
                        # Insert scaling before axId
                        ax_id = val_ax.find("c:axId", NS)
                        if ax_id is not None:
                            idx = list(val_ax).index(ax_id)
                            scaling = etree.Element(f"{{{NS['c']}}}scaling")
                            val_ax.insert(idx, scaling)
                        else:
                            scaling = etree.SubElement(val_ax, f"{{{NS['c']}}}scaling")

                    # Find or create orientation element
                    orientation = scaling.find("c:orientation", NS)
                    if orientation is None:
                        orientation = etree.SubElement(scaling, f"{{{NS['c']}}}orientation")

                    # Set orientation to maxMin to reverse bar direction (RTL)
                    orientation.set("val", "maxMin")
                    updated_count += 1

                    if self.verbose:
                        print(f"[ChartProcessor] Flipped horizontal bar chart to RTL")

        # Save output
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        tree.write(output_path, encoding="UTF-8", xml_declaration=True)

        if self.verbose:
            print(f"[ChartProcessor] Updated {updated_count} text elements")
            print(f"[ChartProcessor] Saved to: {output_path}")

    # ========================================================================
    # UTILITIES
    # ========================================================================
    def save_json(self, content: Dict[str, Any], output_path: str) -> None:
        """Save extracted chart content to a JSON file."""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(content, f, ensure_ascii=False, indent=2)
        if self.verbose:
            print(f"[ChartProcessor] JSON saved to: {output_path}")

    def load_json(self, input_path: str) -> Dict[str, Any]:
        """Load chart content from a JSON file."""
        with open(input_path, 'r', encoding='utf-8') as f:
            return json.load(f)


# ============================================================================
# MAIN EXECUTION
# ============================================================================
if __name__ == "__main__":
    import datetime

    processor = ChartProcessor(verbose=True)

    # Test with the extracted chart
    CHART_XML = "chart1_sample.xml"

    if not os.path.exists(CHART_XML):
        print(f"ERROR: {CHART_XML} not found. Extract a chart first.")
        exit(1)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    OUTPUT_JSON = f"chart1_content_{timestamp}.json"
    OUTPUT_TRANSLATED = f"chart1_translated_{timestamp}.xml"

    print(f"\n{'='*50}")
    print("CHART TEXT EXTRACTION TEST")
    print(f"{'='*50}")
    print(f"Using source: {CHART_XML}")

    # 1. Extract chart text
    chart_content = processor.extract_chart_text(CHART_XML)
    processor.save_json(chart_content, OUTPUT_JSON)

    print("\nExtracted chart content:")
    print(json.dumps(chart_content, ensure_ascii=False, indent=2))

    # 2. Simulate translation
    print(f"\n{'='*50}")
    print("SIMULATED TRANSLATION")
    print(f"{'='*50}")

    translated = {
        "chart_title": "عنوان الرسم البياني",
        "series": [
            {"id": "0", "name": "السلسلة 1"},
            {"id": "1", "name": "السلسلة 2"},
            {"id": "2", "name": "السلسلة 3"}
        ],
        "categories": ["الفئة 1", "الفئة 2", "الفئة 3", "الفئة 4"]
    }

    # 3. Test injection
    processor.inject_chart_text(CHART_XML, translated, OUTPUT_TRANSLATED)
    print(f"\nTranslated chart saved to: {OUTPUT_TRANSLATED}")
