# Streamlit Web App - Quick Start Guide

## Launch the Web Application

### Option 1: Using UV (Recommended)
```bash
uv run streamlit run streamlit_app.py
```

### Option 2: Using Python directly
```bash
streamlit run streamlit_app.py
```

The app will automatically open in your default browser at `http://localhost:8501`

---

## Using the Web Interface

### Step 1: Upload Your Presentation
1. Click **"Browse files"** button
2. Select your PowerPoint file (.pptx)
3. Wait for automatic analysis (shows slide count, charts, etc.)

### Step 2: Configure Translation Settings

#### Slide Selection
- **All slides**: Translates the entire presentation
- **Custom range**: Specify which slides to process
  - Examples:
    - `1,3,5` - Translates slides 1, 3, and 5
    - `1-10` - Translates slides 1 through 10
    - `1-5,8,10-15` - Mix of ranges and individual slides

#### Choose AI Engine
- **OpenAI (GPT-4)**:
  - Best for: Technical content, fast processing
  - Cost: ~$0.01-0.03 per slide
  - Get API key: https://platform.openai.com/api-keys

- **Anthropic (Claude)**:
  - Best for: Consulting terminology, context-aware
  - Cost: ~$0.02-0.04 per slide
  - Get API key: https://console.anthropic.com/

### Step 3: Enter API Key
1. Select your chosen AI engine
2. Paste your API key in the password field
3. The key is only used for this session and not stored

**Security Note**: Your API key is used directly to call the AI service. It's never stored on disk or sent to any third-party servers.

### Step 4: Start Translation
1. Click **"Start Translation"** button
2. Monitor progress in the processing log
3. Wait for completion (typical: 2-3 seconds per slide)

### Step 5: Download Result
1. Click **"Download Translated Presentation"** button
2. Open the file in PowerPoint to verify
3. All text will be in Arabic with RTL layout

---

## What Gets Translated

### ✅ Text Content
- Slide titles
- Body text and bullet points
- Text boxes and annotations
- Table cell content
- Chart titles, series names, axis labels
- Master template text
- Footer and header text

### ✅ Visual Transformations
- All elements mirrored (Right-to-Left layout)
- Text direction set to RTL
- Text alignment reversed (Left↔Right)
- Table columns reversed
- Horizontal bar charts flipped
- Arabic fonts automatically applied

### ✅ Preserved Elements
- Colors and formatting
- Fonts styles (bold, italic, size)
- Animations and transitions
- Images and logos (repositioned, not flipped)
- Corporate branding
- Slide layouts and masters

---

## Troubleshooting

### "API key required" message
- Make sure you've entered your API key in the password field
- Verify the key is valid (test it on OpenAI/Anthropic website)

### "Error analyzing presentation"
- Ensure the file is a valid .pptx format
- Try opening the file in PowerPoint first to verify it's not corrupted
- Check if the file is password-protected (not supported)

### Translation fails
- Check the processing log for specific error messages
- Verify your API key has sufficient credits
- Ensure you have internet connection for AI API calls

### Slow processing
- Large presentations (100+ slides) take longer
- Processing time: ~2-3 seconds per slide
- Charts add ~5-10 seconds per chart
- Consider translating in batches (use custom range)

### Output file doesn't open
- Check if the original file had any corruption
- Try translating fewer slides first (e.g., slides 1-5)
- Review the processing log for errors

---

## Best Practices

### For Best Results:
1. **Test with a single slide first**: Use custom range "1" to test quickly
2. **Process in batches**: For large presentations, do 20-30 slides at a time
3. **Check the output**: Open and verify before proceeding with full presentation
4. **Keep originals**: Always maintain a backup of your original file

### Performance Tips:
- **Small presentations (1-20 slides)**: Translate all at once (~30-60 seconds)
- **Medium presentations (21-50 slides)**: Translate all or in 2-3 batches
- **Large presentations (51-100+ slides)**: Process in batches of 25-30 slides

### Cost Optimization:
- **Test with mock translator first**: Add `--translator mock` for testing (no API cost)
- **Translate only necessary slides**: Use custom range to skip title/thank you slides
- **Choose the right engine**: OpenAI is slightly cheaper for straightforward content

---

## Features Explained

### Chart Processing
- Automatically detects all charts in the presentation
- Translates chart titles, series names, and axis labels
- **Advanced**: Reverses horizontal bar chart orientation (bars grow right-to-left)
- Preserves data accuracy and visual quality

### Table Handling
- Reverses column order for RTL reading flow
- Translates all cell content
- Maintains table formatting and colors

### Arabic Font Support
- Automatically injects "Simplified Arabic" font
- Ensures proper text rendering across all PowerPoint versions
- Maintains corporate font styles where possible

### Error Recovery
- If one element fails, processing continues
- Detailed error logging for troubleshooting
- Partial results delivered even if some elements fail

---

## Privacy & Security

### What Happens to Your Files:
1. **Upload**: File saved temporarily on your local machine
2. **Processing**: Processed locally using the Python system
3. **AI Translation**: Only text content sent to AI API (OpenAI/Anthropic)
4. **Output**: Generated on your local machine
5. **Cleanup**: Temporary files automatically deleted after session

### What's NOT Stored:
- ❌ Your API keys (not saved anywhere)
- ❌ Your presentations (temporary files deleted)
- ❌ Processing history (no database)
- ❌ User information (no tracking)

### API Key Security:
- Entered as password (hidden from view)
- Used only for AI translation API calls
- Exists only in current session memory
- Never written to disk or logs

---

## Technical Requirements

### Minimum Requirements:
- Python 3.10 or higher
- 4GB RAM (8GB recommended)
- Internet connection (for AI API calls)
- Modern web browser (Chrome, Firefox, Safari, Edge)

### Supported File Formats:
- ✅ PowerPoint 2007+ (.pptx)
- ❌ PowerPoint 97-2003 (.ppt) - not supported
- ❌ Password-protected files - not supported

### Maximum File Size:
- Recommended: Up to 50MB
- Maximum: 200MB
- Large files (100MB+) may take longer to upload and process

---

## Support & Feedback

### Need Help?
- Check the sidebar **"User Guide"** for detailed instructions
- Review the **"Processing Log"** for specific error messages
- Consult the main documentation: `EXECUTIVE_SUMMARY.md`

### Report Issues:
- Document the error message from processing log
- Note which slide failed (if applicable)
- Provide presentation characteristics (slide count, chart count)

---

## Next Steps After Translation

### Quality Review Checklist:
1. ✅ Open translated file in PowerPoint
2. ✅ Check first and last slides
3. ✅ Verify charts are displaying correctly
4. ✅ Confirm text is right-aligned and RTL
5. ✅ Review any slides with complex layouts
6. ✅ Test animations and transitions

### If Issues Found:
1. Note the specific slide number
2. Check the original slide for complexity
3. Consider re-translating that specific slide
4. Review processing log for warnings

### Ready for Delivery:
- Save the translated file with appropriate name
- Share with stakeholders for review
- Make any manual adjustments if needed
- Archive both original and translated versions
