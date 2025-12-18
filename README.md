# Voice Excel - Voice-Controlled Spreadsheet Assistant

## Overview
Voice Excel is a Python application that allows you to control Microsoft Excel using voice commands. Hold the SPACE key to speak, and the application will convert your speech into Excel actions automatically using AI interpretation.

## Features
- **Voice-to-Speech Recognition**: Real-time speech capture using Google's speech recognition API
- **AI Command Interpretation**: Google Gemini AI converts natural language commands into structured Excel operations
- **Excel Automation**: Direct COM interface to Microsoft Excel for seamless control
- **Interactive Commands**:
  - Write data to cells
  - Delete cell contents
  - Insert rows and columns
  - Calculate column sums
  - Format text (bold)
  - Create charts
  - Run regression analysis
  - Sort columns
  - Filter values

## Project Structure
- `main.py` - Main application loop that handles voice commands and Excel operations
- `speech_to_text.py` - Speech recognition module using keyboard trigger (SPACE key)
- `gemini_ai.py` - AI command interpreter powered by Google Gemini API
- `excel_actions.py` - Excel automation functions using Windows COM interface
- `.env` - Environment variables (API keys, Excel file path)
- `requirements.txt` - Python package dependencies

## Requirements
- Python 3.7+
- Microsoft Excel installed on Windows
- Internet connection (for speech recognition and Gemini AI)
- Google Gemini API key

## Installation
1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Create a `.env` file with:
   ```
   GEMINI_API_KEY=your_api_key_here
   EXCEL_FILE=path_to_your_excel_file.xlsx
   MODEL_NAME=models/gemini-pro-latest
   ```

## Usage
1. Open your Excel file
2. Run the application:
   ```
   python main.py
   ```
3. Hold SPACE key to record your voice command (max 15 seconds)
4. Release SPACE to process the command
5. The AI interprets your command and executes it in Excel

## Example Commands
- "Write hello in cell A1"
- "Delete B5"
- "Insert row 3"
- "Sum column A"
- "Make column B bold"
- "Create a chart with A and B columns"

## Notes
- This is the Windows version with Python virtual environment support
- Requires active Excel application running
- Speech recognition works best in quiet environments

