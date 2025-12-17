"""
Streamlit Web Application for Slide Translator
Project C - English to Arabic Presentation Translation

This app provides a user-friendly interface for translating PowerPoint presentations
from English to Arabic with RTL layout transformation.
"""

import streamlit as st
import os
import sys
import datetime
import tempfile
from pathlib import Path
from io import StringIO

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from xml_service.xml_extractor import PPTXXMLExtractor

# Import the main translation orchestrator
sys.path.insert(0, os.path.dirname(__file__))
from main import SlideTranslator, parse_slides_arg

# Page configuration
st.set_page_config(
    page_title="Slide Translator - Project C",
    page_icon="🌐",
    layout="wide"
)

# Custom CSS for better UI
st.markdown("""
    <style>
    .project-watermark {
        position: absolute;
        top: 80px;
        right: 20px;
        font-size: 4rem;
        font-weight: 900;
        color: rgba(30, 136, 229, 0.08);
        letter-spacing: 2px;
        z-index: 0;
        pointer-events: none;
    }
    .main-header {
        font-size: 2.5rem;
        color: #1e88e5;
        font-weight: bold;
        margin-bottom: 1rem;
        position: relative;
        z-index: 1;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #424242;
        margin-bottom: 2rem;
        position: relative;
        z-index: 1;
    }
    .info-box {
        background-color: #e3f2fd;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #c8e6c9;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff9c4;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Project watermark background
st.markdown('<div class="project-watermark">PROJECT C</div>', unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">🌐 Slide Translator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Professional English → Arabic Presentation Translation with RTL Layout</div>', unsafe_allow_html=True)

# Sidebar - User Guide
with st.sidebar:
    st.header("📖 User Guide")

    st.markdown("""
    ### How to Use This Tool

    **Step 1: Upload Your Presentation**
    - Click "Browse files" and select your PowerPoint file
    - Supported format: .pptx
    - Max file size: 200MB

    **Step 2: Configure Translation**
    - **Slides to Process**: Choose which slides to translate
      - *All*: Translate entire presentation
      - *Custom Range*: Specify slides (e.g., "1,3,5" or "1-10")

    **Step 3: Select AI Engine**
    - **OpenAI (GPT-5-mini)**:
      - Best for technical content
      - Fast and cost-efficient
      - Cost: ~$0.001-0.003 per slide
    - **Anthropic (Claude)**:
      - Excellent for consulting terminology
      - Context-aware translations
      - Cost: ~$0.02-0.04 per slide

    **Step 4: Provide API Key**
    - Get your API key from:
      - OpenAI: https://platform.openai.com/api-keys
      - Anthropic: https://console.anthropic.com/
    - Your key is used only for this session
    - Not stored anywhere

    **Step 5: Process**
    - Click "Start Translation"
    - Wait for processing (2-3 seconds per slide)
    - Download your translated presentation
    """)

    st.markdown("---")

    st.markdown("""
    ### ✨ What This Tool Does

    1. **Translates All Text**
       - Titles, body content, annotations
       - Chart labels and table content
       - Master templates and footers

    2. **Mirrors Visual Layout**
       - Flips all elements for Right-to-Left
       - Reverses table columns
       - Flips horizontal bar charts

    3. **Preserves Formatting**
       - Colors, fonts, animations
       - Corporate branding intact
       - Professional quality maintained

    ### ⚡ Advanced Features
    - Chart text translation
    - Bar chart orientation reversal
    - Arabic font injection
    - Multi-slide batch processing

    ### 🔒 Privacy & Security
    - Files processed locally
    - API keys not stored
    - Temporary files auto-deleted
    - No data sent to third parties
    """)

# Main content area
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📁 Upload Presentation")

    uploaded_file = st.file_uploader(
        "Choose a PowerPoint file (.pptx)",
        type=['pptx'],
        help="Upload your English presentation to translate to Arabic"
    )

    if uploaded_file is not None:
        # Show file info
        file_size = uploaded_file.size / (1024 * 1024)  # Convert to MB
        st.markdown(f"""
        <div class="info-box">
        <strong>📄 File Uploaded:</strong> {uploaded_file.name}<br>
        <strong>📊 Size:</strong> {file_size:.2f} MB
        </div>
        """, unsafe_allow_html=True)

        # Save uploaded file temporarily
        temp_dir = tempfile.mkdtemp()
        temp_input_path = os.path.join(temp_dir, uploaded_file.name)

        with open(temp_input_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())

        # Analyze presentation
        try:
            with st.spinner("Analyzing presentation..."):
                extractor = PPTXXMLExtractor(temp_input_path)
                slide_count = extractor.get_slide_count()
                chart_count = extractor.get_chart_count()
                master_count = extractor.get_slide_master_count()
                layout_count = extractor.get_slide_layout_count()

            st.markdown(f"""
            <div class="info-box">
            <strong>📊 Presentation Analysis:</strong><br>
            • Total Slides: {slide_count}<br>
            • Charts: {chart_count}<br>
            • Master Templates: {master_count}<br>
            • Layout Templates: {layout_count}
            </div>
            """, unsafe_allow_html=True)

            # Store in session state
            st.session_state.temp_input_path = temp_input_path
            st.session_state.slide_count = slide_count
            st.session_state.chart_count = chart_count

        except Exception as e:
            st.error(f"❌ Error analyzing presentation: {str(e)}")
            st.stop()

with col2:
    st.header("⚙️ Configuration")

    if uploaded_file is not None:
        # Slide selection
        st.subheader("1. Select Slides")
        slide_mode = st.radio(
            "Which slides to translate?",
            ["All slides", "Custom range"],
            help="Choose whether to translate the entire presentation or specific slides"
        )

        if slide_mode == "Custom range":
            slide_range = st.text_input(
                "Slide numbers",
                placeholder="e.g., 1,3,5 or 1-10",
                help="Enter slide numbers separated by commas, or use ranges (e.g., 1-5)"
            )
        else:
            slide_range = "all"

        # AI Engine selection
        st.subheader("2. Choose AI Engine")
        ai_engine = st.selectbox(
            "Translation Engine",
            ["OpenAI (GPT-5-mini)", "Anthropic (Claude)"],
            help="Select which AI to use for translation"
        )

        translator_name = "openai" if "OpenAI" in ai_engine else "anthropic"

        # API Key input
        st.subheader("3. Enter API Key")
        api_key = st.text_input(
            f"{ai_engine.split()[0]} API Key",
            type="password",
            help="Your API key is used only for this session and not stored"
        )

        if not api_key:
            st.markdown("""
            <div class="warning-box">
            ⚠️ API key required to use AI translation
            </div>
            """, unsafe_allow_html=True)

        # Estimated cost
        if api_key:
            estimated_slides = st.session_state.get('slide_count', 0)
            cost_per_slide = 0.002 if "OpenAI" in ai_engine else 0.03
            estimated_cost = estimated_slides * cost_per_slide

            st.markdown(f"""
            <div class="info-box">
            <strong>💰 Estimated Cost:</strong><br>
            {estimated_slides} slides × ${cost_per_slide:.3f} = ${estimated_cost:.3f}
            </div>
            """, unsafe_allow_html=True)

# Processing section
if uploaded_file is not None:
    st.markdown("---")
    st.header("🚀 Start Translation")

    col_left, col_right = st.columns([1, 1])

    with col_left:
        if api_key:
            process_button = st.button(
                "▶️ Start Translation",
                type="primary",
                use_container_width=True
            )
        else:
            st.button(
                "▶️ Start Translation",
                disabled=True,
                use_container_width=True
            )
            st.caption("⚠️ Please enter API key first")

    with col_right:
        if st.session_state.get('output_path'):
            with open(st.session_state.output_path, 'rb') as f:
                st.download_button(
                    "📥 Download Translated File",
                    data=f.read(),
                    file_name=f"Translated_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

    if 'process_button' in locals() and process_button:
        # Prepare output path
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"Translated_{os.path.splitext(uploaded_file.name)[0]}_{timestamp}.pptx"
        output_dir = os.path.join(os.path.dirname(__file__), "output_pptx")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_filename)

        # Get absolute paths
        input_path = os.path.abspath(st.session_state.temp_input_path)
        output_path_abs = os.path.abspath(output_path)

        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_container = st.expander("📋 Processing Log", expanded=True)

        try:
            # Capture stdout for logging
            import io
            import contextlib

            log_stream = StringIO()

            with log_container:
                st.text("🔍 Starting Translation:")
                st.text(f"Input: {input_path}")
                st.text(f"Output: {output_path_abs}")
                st.text(f"Slides: {slide_range}")
                st.text(f"Translator: {translator_name}")
                st.text("-" * 80)

            status_text.text("🔄 Initializing translation pipeline...")
            progress_bar.progress(10)

            # Create translator instance
            translator = SlideTranslator(
                input_pptx=input_path,
                output_pptx=output_path_abs,
                verbose=True
            )

            # Parse slide range
            if slide_range == "all":
                slide_indices = list(range(1, translator.slide_count + 1))
            else:
                slide_indices = parse_slides_arg(slide_range, translator.slide_count)

            progress_bar.progress(20)
            status_text.text("🔄 Processing slides...")

            # Run translation with captured output
            with log_container:
                log_placeholder = st.empty()

                # Redirect stdout to capture verbose output
                with contextlib.redirect_stdout(log_stream):
                    translator.translate_slides(
                        slide_indices=slide_indices,
                        translator=translator_name,
                        api_key=api_key
                    )

                # Display the log
                log_text = log_stream.getvalue()
                log_placeholder.text_area(
                    "Processing Output",
                    value=log_text,
                    height=300
                )

            progress_bar.progress(100)
            status_text.text("✅ Translation completed successfully!")

            # Show success message
            st.markdown("""
            <div class="success-box">
            <strong>✅ Translation Complete!</strong><br><br>
            Your presentation has been successfully translated to Arabic with RTL layout.<br><br>
            <strong>What was done:</strong><br>
            ✓ All text translated to Arabic<br>
            ✓ Visual layout mirrored for Right-to-Left reading<br>
            ✓ Charts translated and orientation reversed<br>
            ✓ Tables restructured for RTL<br>
            ✓ Arabic fonts applied throughout<br>
            ✓ All formatting and branding preserved
            </div>
            """, unsafe_allow_html=True)

            # Store output path for download button
            st.session_state.output_path = output_path_abs

            # Offer download
            with open(output_path_abs, 'rb') as f:
                st.download_button(
                    "📥 Download Translated Presentation",
                    data=f.read(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary"
                )

            st.balloons()

        except Exception as e:
            st.error(f"❌ Error during translation: {str(e)}")
            import traceback
            st.error(f"Traceback: {traceback.format_exc()}")
            progress_bar.progress(0)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #757575; padding: 2rem;">
<strong>Slide Translator</strong> - Project C<br>
Professional English → Arabic Presentation Translation System<br>
Built with Pure Python for Enterprise Scalability
</div>
""", unsafe_allow_html=True)
