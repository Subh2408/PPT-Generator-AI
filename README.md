# PPT-Generator-AI

This project automates the creation of PowerPoint presentations from Excel files, enriched with AI-generated insights using a locally installed Ollama model.

---

## 🛠 Features

- **Excel Input**: Select Excel (.xlsx) files to use as input.
- **Custom Slide Generation**: Choose the number of slides to generate.
- **AI-Powered Insights**: Connects to your locally installed Ollama model to create recommendations and insights for each slide.
- **Interactive AI Queries**: Ask additional questions to the local Ollama model and get responses to insert into your slides.
- **Packaged as .exe**: Easy-to-use executable available — no need to run Python manually.
- **Guided Ollama Setup**: Comes with a helper guide to install and initialize Ollama.

---

## 🚀 How It Works

1. Launch the application.
2. Select an Excel file(s) as the input source.
3. Choose the number of slides you want to generate.
4. The app reads your data and creates PowerPoint slides, adding AI-generated content.
5. Use the interactive query tool to refine or add more content to your slides.
6. Export the final PowerPoint file.

---

## 📦 Installation

### Prerequisites

- Python 3.8+
- Ollama installed locally (Ollama **must** be downloaded and installed separately. See setup guide.)
- Windows system (for the `.exe` packaged version)

### Ollama Setup

> **Important:**  
> This application requires Ollama to be installed and running on your system.

1. Download and install Ollama from [https://ollama.com/download](https://ollama.com/download).
2. Install and run the model you want (e.g., `ollama run llama3`).
3. Keep Ollama running in the background while using this app.
4. See the `setup_ollama_readme.txt` file in the repository for step-by-step instructions.

---

## ⚙️ Usage

- **Run the Application:**
  - If using the executable: simply double-click to launch.
  - If using Python: `python app.py`
- **Choose Excel File(s):** Select the input file using the GUI.
- **Set Number of Slides:** Enter how many slides you want generated.
- **Click Generate:** The application will create a PowerPoint file.
- **Optional:** Use the query input section to ask more questions and receive answers to paste into your slides.

---

## 🛡️ License

This project is covered by a custom “No Modifications Allowed” license.  
You may view, use, and and share the code **as-is** for personal and non-commercial purposes, but you **may not modify** or redistribute it without explicit permission.  
See the [LICENSE](LICENSE.txt) file for full terms and contact information.  

---

## 🙌 Acknowledgments

- [Ollama](https://ollama.com/) — for providing easy-to-use local models.
- Python community packages — `tkinter`, `pandas`, `python-pptx`, `requests`.

---
