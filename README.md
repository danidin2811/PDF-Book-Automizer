# PDF Book Architect 📚

A professional desktop automation tool built to streamline the digital publication workflow at the **Yad Yitzhak Ben-Zvi Institute**. 

Previously, the process required extensive manual labor using Adobe Acrobat to prepare books for digital sale.
I leveraged my programming skills to first automate these tasks via Command Line scripts, and have since evolved the project into a modern, high-performance desktop application.

## 🚀 The Problem
To sell digital copies of publications alongside physical ones, the institute must prepare several specific PDF assets for every book.
This manual workflow was time-consuming and prone to human error, including:
- **Segmenting Files:** Manually splitting the main PDF into specific files for the full book, example chapters, preface, and table of contents.
- **Bi-Directional Content:** Handling books with English sections, which require a separate file with the page order completely reversed to match digital reading standards.
- **File Management:** Manually normalizing book titles for database consistency, moving files between server directories, and updating cover image paths.

## ✨ Features
- **Modern UI/UX:** Built with `CustomTkinter` for a clean, dark-mode desktop experience.
- **One-Click Automation:** Replaces multiple manual Acrobat steps with a single automated process.
- **Smart Formatting:** Real-time normalization of book titles into system-safe folder names.
- **English Section Handling:** Automatic detection and reversal of English page sequences.
- **Error Reduction:** Automated file moving and naming ensures consistency across the institute's digital library.

## 🛠️ Tech Stack
- **Language:** Python 3.10+
- **GUI Framework:** CustomTkinter
- **PDF Processing:** PyPDF2
- **Data Management:** Openpyxl (Excel Integration)
- **Version Control:** Git

## 📂 Project Structure
```text
PDF-Book-Automizer/
├── src/                # Core application logic and UI
│   ├── main_app.py     # Entry point & GUI Controller
│   ├── toc.py          # Bookmark and TOC processing
│   └── processBook.py  # PDF & File system automation
├── assets/             # UI Icons and brand assets
├── .gitignore          # Keeps work-specific PDFs/Data private
└── README.md           # Documentation
```

## ⚙️ Installation & Usage

### 1. Prerequisites
Ensure you have **Python 3.10** or higher installed on your system. You can check your version by running `python --version` in your terminal.

### 2. Clone the Repository
Open your terminal or PowerShell and run:
```bash
git clone https://github.com/danidin2811/PDF-Book-Automizer.git
cd PDF-Book-Automizer
```

### 3. Install Dependencies
This project uses several specialized libraries for PDF manipulation and GUI rendering. Install them all at once using the provided requirements file:
```bash
pip install -r requirements.txt
```

### 4. Run the Application
To launch the graphical interface, execute the main entry point located in the src directory:
```bash
python src/main_app.py
```
