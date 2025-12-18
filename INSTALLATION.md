# Installation Guide - Slide Translator (Project C)

## Quick Installation (Just Copy & Paste!)

### Step 1: Open PowerShell
- Press `Windows Key` on your keyboard
- Type `PowerShell`
- Click on "Windows PowerShell" (the blue icon)

### Step 2: Install UV Package Manager

Copy and paste this command, then press Enter:
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Wait for it to finish, then **close PowerShell and open it again**.

### Step 3: Download and Install the Tool

Copy and paste ALL of these commands together, then press Enter:
```powershell
cd $HOME\Desktop
git clone https://github.com/mahdizn1/ppt-translator.git
cd ppt-translator
uv sync
```

Wait for installation to complete (30-60 seconds).

### Step 4: Run the Translation Tool

Copy and paste this command:
```powershell
uv run streamlit run streamlit_app.py
```

**Done!** Your browser will open automatically with the translation tool ready to use.

---

## Running the Tool Again Later

When you want to use the tool again:

1. Open PowerShell
2. Copy and paste these commands:
   ```powershell
   cd $HOME\Desktop\ppt-translator
   uv run streamlit run streamlit_app.py
   ```

---

## Getting Your API Key

You'll need an API key from OpenAI or Anthropic:

- **OpenAI**: Visit https://platform.openai.com/api-keys
- **Anthropic**: Visit https://console.anthropic.com/

The tool will ask for your API key when you use it. The key is only used for translation and is never stored.

---

## Need Help?

If you encounter any issues:
1. Make sure you have an internet connection
2. Try closing and reopening PowerShell
3. Check that you completed Step 2 (installing UV) and reopened PowerShell before continuing
