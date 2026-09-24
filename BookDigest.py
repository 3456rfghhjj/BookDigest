import os
import sys
import re
import time
import threading
import tempfile
import webbrowser
import customtkinter as ctk
from tkinter import filedialog, messagebox
from google import genai
import docx
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup
from pypdf import PdfReader

# --- API CONFIGURATION ---
# Insert your Gemini API key inside the quotes:
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"

client = genai.Client(api_key=GEMINI_API_KEY)

# --- TEXT EXTRACTION UTILITIES ---
def extract_text_from_txt(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as file:
        return file.read()

def extract_text_from_docx(path):
    document = docx.Document(path)
    return "\n".join([paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()])

def extract_text_from_epub(path):
    book = epub.read_epub(path)
    text_segments = []
    for item in book.get_items():
        if item.get_type() == ebooklib.ITEM_DOCUMENT:
            soup = BeautifulSoup(item.get_content(), "html.parser")
            text = soup.get_text()
            if text.strip():
                text_segments.append(text.strip())
    return "\n\n".join(text_segments)

def extract_text_from_pdf(path):
    reader = PdfReader(path)
    text_segments = []
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text and page_text.strip():
            text_segments.append(page_text.strip())
    return "\n\n".join(text_segments)

def load_file_content(path):
    extension = os.path.splitext(path)[1].lower()
    if extension == ".txt":
        return extract_text_from_txt(path)
    elif extension == ".docx":
        return extract_text_from_docx(path)
    elif extension == ".epub":
        return extract_text_from_epub(path)
    elif extension == ".pdf":
        return extract_text_from_pdf(path)
    else:
        raise ValueError(f"Unsupported file format: {extension}")

# Split large text into smaller chunks for processing
def split_into_chunks(text, chunk_size=70000):
    chunks = []
    start_index = 0
    total_length = len(text)
    while start_index < total_length:
        end_index = start_index + chunk_size
        if end_index >= total_length:
            chunks.append(text[start_index:])
            break
        split_at = text.rfind("\n\n", start_index, end_index)
        if split_at == -1 or split_at <= start_index:
            split_at = text.rfind(". ", start_index, end_index)
        if split_at == -1 or split_at <= start_index:
            split_at = end_index
        else:
            split_at += 1
        chunks.append(text[start_index:split_at].strip())
        start_index = split_at
    return chunks

# --- MARKDOWN TO HTML CONVERTER ---
def parse_markdown_to_html(text):
    lines = text.split("\n")
    html_lines = []
    is_inside_list = False

    for line in lines:
        stripped_line = line.strip()

        if not stripped_line:
            if is_inside_list:
                html_lines.append("</ul>")
                is_inside_list = False
            continue

        if stripped_line.startswith("### "):
            if is_inside_list:
                html_lines.append("</ul>")
                is_inside_list = False
            html_lines.append(f"<h3>{stripped_line[4:]}</h3>")
            continue
        elif stripped_line.startswith("## "):
            if is_inside_list:
                html_lines.append("</ul>")
                is_inside_list = False
            html_lines.append(f"<h2>{stripped_line[3:]}</h2>")
            continue
        elif stripped_line.startswith("# "):
            if is_inside_list:
                html_lines.append("</ul>")
                is_inside_list = False
            html_lines.append(f"<h1>{stripped_line[2:]}</h1>")
            continue

        if stripped_line.startswith("* ") or stripped_line.startswith("- "):
            if not is_inside_list:
                html_lines.append("<ul>")
                is_inside_list = True
            content = stripped_line[2:]
            content = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", content)
            content = re.sub(r"\*(.+?)\*", r"<em>\1</em>", content)
            html_lines.append(f"<li>{content}</li>")
            continue

        if is_inside_list:
            html_lines.append("</ul>")
            is_inside_list = False

        content = stripped_line
        content = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", content)
        content = re.sub(r"\*(.+?)\*", r"<em>\1</em>", content)
        html_lines.append(f"<p>{content}</p>")

    if is_inside_list:
        html_lines.append("</ul>")

    return "\n".join(html_lines)

def open_summary_in_browser(markdown_text, source_title):
    body_content = parse_markdown_to_html(markdown_text)

    # Warm, eye-friendly dark green theme
    html_page = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Summary: {source_title}</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            line-height: 1.8;
            color: #d6e2d9;
            background-color: #121915;
            margin: 0;
            padding: 40px 20px;
        }}
        .container {{
            max-width: 840px;
            margin: 0 auto;
            background: #1c2720;
            padding: 45px 55px;
            border-radius: 16px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.45);
            border: 1px solid #29382f;
        }}
        h1 {{
            font-size: 26px;
            border-bottom: 2px solid #52b788;
            padding-bottom: 12px;
            margin-top: 0;
            color: #b7e4c7;
            font-weight: 600;
        }}
        h2 {{
            color: #95d5b2;
            margin-top: 32px;
            font-size: 20px;
            border-bottom: 1px solid #2c3e34;
            padding-bottom: 6px;
        }}
        h3 {{
            color: #74c69d;
            margin-top: 20px;
            font-size: 17px;
        }}
        p {{
            font-size: 16px;
            margin: 12px 0;
            color: #d6e2d9;
        }}
        ul {{
            font-size: 16px;
            padding-left: 24px;
            margin: 12px 0;
        }}
        li {{
            margin-bottom: 8px;
            color: #cad8ce;
        }}
        li::marker {{
            color: #52b788;
        }}
        strong {{
            color: #eaf4ee;
            font-weight: 600;
        }}
        em {{
            color: #a3c9b1;
        }}
        .header-meta {{
            font-size: 12px;
            color: #749681;
            margin-bottom: 18px;
            text-transform: uppercase;
            letter-spacing: 1px;
            font-weight: 600;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header-meta">AI Book & Document Summary</div>
        <h1>{source_title}</h1>
        {body_content}
    </div>
</body>
</html>"""

    temp_dir = tempfile.gettempdir()
    html_file_path = os.path.join(temp_dir, "summary_result.html")
    with open(html_file_path, "w", encoding="utf-8") as file:
        file.write(html_page)

    webbrowser.open(f"file://{html_file_path}")

# --- GRAPHICAL USER INTERFACE ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class SummarizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("BookDigest - AI Document Summarizer")
        self.after(0, lambda: self.wm_state("zoomed"))

        self.selected_file_path = None
        self.last_result = ""

        # Top control bar
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.pack(fill="x", padx=15, pady=10)

        self.btn_select = ctk.CTkButton(self.top_frame, text="Select File / Book", command=self.select_file)
        self.btn_select.pack(side="left", padx=10, pady=10)

        self.lbl_file = ctk.CTkLabel(self.top_frame, text="No file selected", anchor="w")
        self.lbl_file.pack(side="left", fill="x", expand=True, padx=10)

        self.btn_open_html = ctk.CTkButton(
            self.top_frame, 
            text="Open in Browser", 
            command=self.manual_open_html, 
            state="disabled", 
            fg_color="#2d6a4f", 
            hover_color="#1b4332"
        )
        self.btn_open_html.pack(side="right", padx=10, pady=10)

        self.btn_run = ctk.CTkButton(
            self.top_frame, 
            text="Generate Summary", 
            command=self.start_summary_thread, 
            fg_color="green", 
            hover_color="darkgreen"
        )
        self.btn_run.pack(side="right", padx=10, pady=10)

        # Status label
        self.lbl_status = ctk.CTkLabel(self, text="Ready", font=ctk.CTkFont(size=12, slant="italic"))
        self.lbl_status.pack(anchor="w", padx=20, pady=(0, 5))

        # Text output box
        self.txt_output = ctk.CTkTextbox(self, font=ctk.CTkFont(size=14), wrap="word")
        self.txt_output.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select a book or document",
            filetypes=[
                ("Supported files", "*.pdf *.epub *.docx *.txt"),
                ("PDF documents", "*.pdf"),
                ("EPUB e-books", "*.epub"),
                ("Word documents", "*.docx"),
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        if path:
            self.selected_file_path = path
            self.lbl_file.configure(text=os.path.basename(path))
            self.lbl_status.configure(text="File loaded. Click 'Generate Summary' to start.")

    def manual_open_html(self):
        if self.last_result:
            title = os.path.basename(self.selected_file_path) if self.selected_file_path else "Document"
            open_summary_in_browser(self.last_result, title)

    def start_summary_thread(self):
        if not self.selected_file_path:
            messagebox.showwarning("Warning", "Please select a file first!")
            return

        self.btn_run.configure(state="disabled")
        self.btn_select.configure(state="disabled")
        self.btn_open_html.configure(state="disabled")
        self.lbl_status.configure(text="Preparing document...")
        self.txt_output.delete("1.0", "end")

        threading.Thread(target=self.process_summary, daemon=True).start()

    def process_summary(self):
        try:
            raw_text = load_file_content(self.selected_file_path)
            
            if not raw_text.strip():
                raise ValueError("File is empty or contains no extractable text (e.g. scanned PDF without OCR).")

            chunks = split_into_chunks(raw_text, chunk_size=70000)
            total_chunks = len(chunks)

            models = [
                "gemini-flash-lite-latest",
                "gemini-3.1-flash-lite",
                "gemini-flash-latest"
            ]

            all_summaries = []

            for index, chunk in enumerate(chunks, 1):
                self.lbl_status.configure(text=f"Processing part {index} of {total_chunks}...")

                # Prompt in English, strictly concise
                prompt = (
                    f"Provide a concise and key-point summary of part {index}/{total_chunks} of this book/document. "
                    "Extract only 3 to 5 key points and crucial facts. "
                    "Be direct, without filler words or conversational intro. Use bullet points:\n\n" + chunk
                )

                chunk_response = None
                for model_name in models:
                    try:
                        response = client.models.generate_content(
                            model=model_name,
                            contents=prompt,
                        )
                        if response and response.text:
                            chunk_response = response.text
                            break
                    except Exception:
                        time.sleep(2)
                        continue

                if not chunk_response:
                    chunk_response = f"*(Part {index} could not be processed due to high server load)*"

                section_header = f"## Part {index} of {total_chunks}\n\n"
                full_section = section_header + chunk_response + "\n\n"
                all_summaries.append(full_section)
                
                self.txt_output.insert("end", full_section)
                self.txt_output.see("end")

                if index < total_chunks:
                    time.sleep(3)

            final_text = "".join(all_summaries)
            self.last_result = final_text
            self.update_gui_success(final_text)

        except Exception as e:
            self.update_gui_error(str(e))

    def update_gui_success(self, text):
        self.lbl_status.configure(text="Summary completed. Opening browser view.")
        self.btn_run.configure(state="normal")
        self.btn_select.configure(state="normal")
        self.btn_open_html.configure(state="normal")

        title = os.path.basename(self.selected_file_path) if self.selected_file_path else "Document"
        open_summary_in_browser(text, title)

    def update_gui_error(self, error_message):
        messagebox.showerror("Error", f"An error occurred: {error_message}")
        self.lbl_status.configure(text="Processing error.")
        self.btn_run.configure(state="normal")
        self.btn_select.configure(state="normal")

if __name__ == "__main__":
    app = SummarizerApp()
    app.mainloop()