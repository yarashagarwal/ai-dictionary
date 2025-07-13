# Norwegian Dictionary CLI

This project is a command-line tool for building and managing a Norwegian-English dictionary using AI and Excel. It allows you to add new words with AI-generated translations and examples, search for words, and view or export your dictionary in both Excel and Markdown formats.

## Features
- **Add new words**: Enter Norwegian words, and the tool uses Azure OpenAI to generate English meanings and example sentences. New entries are saved to both `dict.xlsx` (Excel) and `dict.md` (Markdown).
- **Avoid duplicates**: The tool checks for existing words before making API calls or adding to the dictionary.
- **Read all entries**: Display the entire dictionary in a readable table format.
- **Search**: Find a specific word and display its details.
- **Alphabetical order**: The Excel file is always kept alphabetically sorted by word.

## File Structure
- `main.py` — The main CLI entry point. Handles command-line arguments for writing, reading, and searching the dictionary.
- `getaidata.py` — Contains the function to query Azure OpenAI for word meanings and examples. Loads API credentials from `.env`.
- `excel.py` — Provides the `excel_action` function for reading, writing, and appending to the Excel file.
- `dict.xlsx` — The main Excel dictionary file.
- `dict.md` — Markdown export of all new entries added.
- `.env` — Stores your Azure OpenAI credentials (not tracked by git).
- `.gitignore` — Ensures `.env` and other sensitive files are not committed.

## Usage

### 1. Setup
- Install dependencies:
  ```zsh
  pip install -r requirements.txt
  ```
- Create a `.env` file with your Azure OpenAI credentials:
  ```env
  AZURE_OPENAI_KEY=your-key-here
  AZURE_OPENAI_ENDPOINT=https://your-endpoint.openai.azure.com/
  AZURE_OPENAI_DEPLOYMENT=your-deployment-name
  ```

### 2. Commands
- **Add new words:**
  ```zsh
  python main.py -w
  ```
  Enter words one per line. The tool will generate and add new entries.

- **Read all words:**
  ```zsh
  python main.py -r
  ```
  Shows the entire dictionary in a table.

- **Search for a word:**
  ```zsh
  python main.py -s <word>
  ```
  Finds and displays the entry for `<word>`.

## Notes
- The tool uses the `tabulate` library for pretty table output and `openpyxl` for Excel file handling.
- All new entries are appended to both `dict.xlsx` and `dict.md`.
- The Excel file is always kept alphabetically sorted after each write.

## License
This project is for educational and personal use.
