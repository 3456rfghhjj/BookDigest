# BookDigest 📚

A modern desktop application built with Python and CustomTkinter to summarize large documents and whole books (PDF, EPUB, DOCX, TXT) using Google Gemini AI, featuring an eye-friendly dark green HTML output.

## Features

- **Multi-Format Support:** Reads `.pdf`, `.epub`, `.docx`, and `.txt` files.
- **Whole-Book Processing:** Splits large texts into 70,000-character chunks to process entire books without hitting API context or rate limits.
- **Concise Summaries:** Extracts 3–5 key takeaways per section directly, without conversational filler.
- **Eye-Friendly Reading View:** Automatically generates and opens a clean, dark-green themed HTML summary in your default browser.
- **Model Fallback:** Automatically switches between lightweight Gemini models if the server experiences high traffic.

## Requirements

- Python 3.10+
- Google Gemini API key

## Installation

1. Clone this repository:
```bash
git clone https://github.com/YOUR_USERNAME/BookDigest.git
cd BookDigest
```
2. Install dependencies:
```bash
pip install customtkinter google-genai python-docx EbookLib beautifulsoup4 pypdf
```

## Configuration

Open the script and add your Gemini API key:
```python
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

## Usage

Run the script:

```bash
python BookDigest.py
```

1. Click Select File / Book and choose your file.
2. Click Generate Summary.
3. Once finished, the formatted summary opens automatically in your browser.

## License

MIT License
