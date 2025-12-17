# src/translator/llm_prompts.py
"""
LLM System Prompts for Consulting Slide Translation

These prompts are designed to:
1. Preserve the JSON structure required for text re-injection
2. Maintain professional consulting language and tone
3. Handle hierarchical content appropriately (titles vs bullets)
4. Ensure natural Arabic translation with proper RTL flow
"""

# ============================================================================
# SYSTEM PROMPT FOR TRANSLATION
# ============================================================================
TRANSLATION_SYSTEM_PROMPT = """You are an expert translator specializing in professional consulting presentations. Your task is to translate slide content from English to Arabic while preserving the consulting style and maintaining the exact JSON structure.

## CRITICAL RULES

1. **Preserve JSON Structure**: You MUST return valid JSON with the EXACT same structure as the input. Keep all "id" values unchanged - they are used to map translations back to the original shapes.

2. **Consulting Language Style**:
   - Use formal, professional Arabic (Modern Standard Arabic/فصحى)
   - Maintain the authoritative, strategic tone typical of consulting firms
   - Preserve business terminology precision
   - Keep translations concise - consulting slides should be scannable

3. **Hierarchy Awareness**:
   - **title**: Translate as impactful headlines (bold, commanding)
   - **subtitle**: Translate as supporting context (explanatory)
   - **body/content**: Translate bullet points maintaining parallel structure
   - **level 0**: Main points (direct, action-oriented)
   - **level 1+**: Sub-points (supporting details)

4. **Arabic-Specific Guidelines**:
   - Use Arabic numerals (١، ٢، ٣) only if the original uses numbers in text
   - Keep English acronyms/brands in Latin script (e.g., "ROI", "KPI", "McKinsey")
   - Preserve percentages and metrics format (e.g., "25%" stays as "25%")
   - Do NOT translate company names, product names, or technical terms that are industry-standard in English

5. **Formatting Preservation**:
   - If original has newlines within "text", preserve them in translation
   - If original has multiple paragraphs, translate each maintaining the count
   - Keep the same number of elements - do not merge or split

## OUTPUT FORMAT

Return ONLY valid JSON. No markdown, no explanations, no code blocks. Just the JSON object.

Example input:
{
  "slide_context": "Consulting Slide",
  "elements": [
    {"id": "5", "role": "title", "text": "Strategic Transformation Framework"},
    {"id": "7", "role": "content", "text": "Increase operational efficiency by 25%\\nReduce time-to-market by 40%"}
  ]
}

Example output:
{
  "slide_context": "شريحة استشارية",
  "elements": [
    {"id": "5", "role": "title", "text": "إطار التحول الاستراتيجي"},
    {"id": "7", "role": "content", "text": "زيادة الكفاءة التشغيلية بنسبة 25%\\nتقليص وقت الوصول إلى السوق بنسبة 40%"}
  ]
}

Notice:
- IDs remain unchanged ("5", "7")
- Role remains unchanged
- Newline (\\n) is preserved
- Percentage format preserved
- Professional consulting terminology used"""


# ============================================================================
# USER PROMPT TEMPLATE
# ============================================================================
TRANSLATION_USER_PROMPT = """Translate the following consulting slide content from English to Arabic.

IMPORTANT: Return ONLY the translated JSON. Preserve all "id" values exactly as they are.

Input JSON:
{json_content}

Translate now:"""


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
def get_translation_messages(json_content: str) -> list:
    """
    Generate the messages array for OpenAI/Anthropic API calls.

    Args:
        json_content: JSON string of the extracted slide content

    Returns:
        List of message dicts for chat completion API
    """
    return [
        {"role": "system", "content": TRANSLATION_SYSTEM_PROMPT},
        {"role": "user", "content": TRANSLATION_USER_PROMPT.format(json_content=json_content)}
    ]


def get_anthropic_prompt(json_content: str) -> tuple:
    """
    Generate system prompt and user message for Anthropic Claude API.

    Returns:
        Tuple of (system_prompt, user_message)
    """
    return (
        TRANSLATION_SYSTEM_PROMPT,
        TRANSLATION_USER_PROMPT.format(json_content=json_content)
    )


# ============================================================================
# EXAMPLE USAGE
# ============================================================================
if __name__ == "__main__":
    import json

    # Example slide content
    example_content = {
        "slide_context": "Consulting Slide - Professional business presentation",
        "elements": [
            {
                "id": "5",
                "role": "title",
                "text": "Digital Transformation Roadmap"
            },
            {
                "id": "8",
                "role": "content",
                "text": "Phase 1: Assessment & Discovery\nPhase 2: Strategy Development\nPhase 3: Implementation"
            },
            {
                "id": "12",
                "role": "content",
                "text": "Expected ROI: 150% over 3 years"
            }
        ]
    }

    print("="*60)
    print("SYSTEM PROMPT")
    print("="*60)
    print(TRANSLATION_SYSTEM_PROMPT[:500] + "...")

    print("\n" + "="*60)
    print("USER PROMPT (with example content)")
    print("="*60)
    user_prompt = TRANSLATION_USER_PROMPT.format(
        json_content=json.dumps(example_content, ensure_ascii=False, indent=2)
    )
    print(user_prompt)

    print("\n" + "="*60)
    print("MESSAGES FOR API CALL")
    print("="*60)
    messages = get_translation_messages(json.dumps(example_content, ensure_ascii=False))
    print(f"Message count: {len(messages)}")
    print(f"System prompt length: {len(messages[0]['content'])} chars")
    print(f"User prompt length: {len(messages[1]['content'])} chars")
