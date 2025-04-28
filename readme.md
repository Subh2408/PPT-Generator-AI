-------------------------------------
Excel to PowerPoint AI Generator Setup
-------------------------------------
Version: 1.0 (Date: YYYY-MM-DD) - Adjust as needed

**Purpose:**
This tool automatically generates PowerPoint presentations from data found in multiple Excel files,
using Artificial Intelligence (AI) analysis provided by a local Ollama language model. It can
also perform custom AI queries on the processed data.

**-------------------------------**
**!! IMPORTANT PREREQUISITES !!**
**-------------------------------**

Before you can use this Generator tool, you MUST install and set up the Ollama service on your computer.
This is a **one-time setup**.

**1. Install Ollama Application:**
   - Go to the official Ollama website: https://ollama.com/
   - Download and run the installer for your operating system (Windows / macOS / Linux).

**2. Download the Required AI Model:**
   - Once Ollama is installed, open your system's **Terminal** (Mac/Linux) or **Command Prompt / PowerShell** (Windows).
     (You might need to run Command Prompt/PowerShell "As Administrator" on Windows).
   - Type the following command EXACTLY and press Enter:
     ```
     ollama pull llama3:8b
     ```
   - This will download the specific AI model needed (Llama 3, 8 billion parameters). It's a large file (several GB), so be patient.
   - Wait until the download and verification are fully complete.
   - *Note: If you know this tool was configured for a different model, use that specific `ollama pull model:tag` command instead.*

**-------------------------------**
**HOW TO RUN THE GENERATOR**
**-------------------------------**

You need to follow these steps **each time** you want to use the tool:

**1. Start Ollama Server:**
   - Open a **NEW** Terminal or Command Prompt/PowerShell window.
   - Type the command below and press Enter:
     ```
     ollama serve
     ```
   - You should see messages indicating the server is "Listening".
   - **CRITICAL: You MUST leave this terminal window open and running the entire time you are using the PowerPoint Generator.** Closing this window will stop the AI service.

**2. Run the PowerPoint Generator Application:**
   - Navigate to the folder named `PPT_Generator_AI`.
   - Inside this folder, find and double-click the `PPT_Generator_AI.exe` file.
   - The application window should appear.

**3. Use the Application:**
   - Click "Add Excel..." to select one or more Excel files.
   - Click "Browse..." to select your PowerPoint template (`.pptx`) file.
   - Set the desired "Max Slides" count.
   - Click "Generate Presentation".
   - Choose where to save the output file.
   - Wait for the process to complete (watch the status bar).
   - After successful generation, the "Custom LLM Query..." button will become active if you wish to ask follow-up questions about the data.

**-------------------------------**
**TROUBLESHOOTING**
**-------------------------------**

*   **Error mentioning "Ollama connection error", "Failed during chat", or similar:**
    - Make sure the terminal window running `ollama serve` is still open and running.
    - Double-check that you successfully ran `ollama pull llama3:8b` previously. You can verify by opening another terminal and running `ollama list`. `llama3:8b` should appear in the list.
    - Restart your computer and try running `ollama serve` then the application again.

*   **Application doesn't start / Closes immediately:** There might be an issue with the packaged files or a missing dependency on the target computer. Contact the provider of this tool.

*   **Excel Reading Errors:** Ensure the Excel files contain sheets with "main data" or "results" in their names (case-insensitive) and that the tables are reasonably standard.

-------------------------------------