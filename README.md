# Slide Translator - Project C

Professional English → Arabic Presentation Translation System with RTL Layout

## 🚀 Quick Start

### Web Interface (Recommended)
```bash
uv run streamlit run streamlit_app.py
```
Open your browser to http://localhost:8501 and use the visual interface!

See [STREAMLIT_GUIDE.md](STREAMLIT_GUIDE.md) for detailed instructions.

### Command Line Interface
```bash
# Translate all slides
uv run python main.py input.pptx output.pptx --slides all --translator openai

# Translate specific slides
uv run python main.py input.pptx output.pptx --slides 1,3,5 --translator anthropic

# Set API key
export OPENAI_API_KEY="your-key-here"
export ANTHROPIC_API_KEY="your-key-here"
```

## 📖 Documentation

- **[EXECUTIVE_SUMMARY.md](EXECUTIVE_SUMMARY.md)** - System overview and key features (5-minute read)
- **[STREAMLIT_GUIDE.md](STREAMLIT_GUIDE.md)** - Web app user guide
- **[DATAFLOW_DIAGRAM.md](DATAFLOW_DIAGRAM.md)** - Visual data flow diagrams
- **[SYSTEM_ARCHITECTURE_DOCUMENT.md](SYSTEM_ARCHITECTURE_DOCUMENT.md)** - Comprehensive technical documentation

## ✨ What It Does

1. **Translates all text** - Titles, body content, charts, tables, annotations
2. **Mirrors visual layout** - Perfect Right-to-Left geometric transformation
3. **Preserves formatting** - Colors, fonts, animations, corporate branding
4. **Advanced chart support** - Translates AND reverses bar chart orientation

## 🎯 Features

- ✅ Context-aware AI translation (OpenAI GPT-4 / Anthropic Claude)
- ✅ Complete visual mirroring for RTL layout
- ✅ Chart text translation + orientation reversal
- ✅ Table column reversal
- ✅ Arabic font injection
- ✅ Master template & layout translation
- ✅ Batch processing support
- ✅ Error-resilient processing
- ✅ Web interface for easy use

## 📦 Installation

```bash
# Install UV (if not already installed)
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone/download the project
cd ppt-translator

# Dependencies are managed by UV automatically
```

## 🔧 Requirements

- Python 3.10+
- UV package manager
- OpenAI or Anthropic API key
- Internet connection (for AI translation)

## 💰 Cost

- OpenAI: ~$0.01-0.03 per slide
- Anthropic: ~$0.02-0.04 per slide
- Processing: Free (runs locally)

## 🛠️ Project Structure

```
ppt-translator/
├── main.py                    # CLI entry point
├── streamlit_app.py           # Web interface
├── src/
│   ├── xml_service/          # PowerPoint file handling
│   ├── translator/           # Translation & transformation
│   └── ...
├── EXECUTIVE_SUMMARY.md      # Overview for stakeholders
├── STREAMLIT_GUIDE.md        # Web app user guide
└── DATAFLOW_DIAGRAM.md       # Visual architecture
```

## 📊 Performance

- Processing: 2-3 seconds per slide
- 100-slide presentation: ~5 minutes
- Parallel processing supported
- Scales to 500+ slides

## 🔒 Privacy & Security

- All processing happens locally
- API keys never stored
- Only text sent to AI (no images)
- Temporary files auto-deleted
- No tracking or data collection

## 🎓 Built With

- Pure Python (lxml, etree)
- OpenAI / Anthropic APIs
- Streamlit (web interface)
- UV package manager

## 📝 License

Project C - Academic/Research Use

## 🤝 Support

For issues or questions:
1. Check [STREAMLIT_GUIDE.md](STREAMLIT_GUIDE.md) troubleshooting section
2. Review [EXECUTIVE_SUMMARY.md](EXECUTIVE_SUMMARY.md) for system overview
3. Consult processing logs for detailed errors

---

**Built for Enterprise Scalability** | **Pure Python Implementation** | **Professional Quality Translation**
