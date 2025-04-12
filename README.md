# Email_gemini_analyzer_project

An intelligent email analysis tool that reads emails from a PST file and leverages Google's Gemini AI to answer natural language questions about them.

## Features
- Read and parse emails from Outlook PST files
- Summarize and extract useful email content
- Ask natural language questions about emails
- Powered by Gemini 1.5 Flash

## Setup

### Prerequisites
- Python 3.8+
- Microsoft Outlook (for reading PST files via COM)
- A valid Gemini API Key from Google Generative AI

### Installation

```bash
pip install -r requirements.txt
```

### Environment Setup
Create a `.env` file or set the API key manually:
```bash
# .env file or environment variable
GEMINI_API_KEY=your_api_key_here
```

## Usage

```bash
python main.py
```
Follow the prompt to input your PST file path and ask questions about its contents.

## Example Questions
- "How many emails were sent by John?"
- "Summarize the topics discussed last week."
- "Are there any urgent emails I missed?"

## License
This project is licensed under the MIT License. See `LICENSE` for details.

## Disclaimer
Ensure compliance with data privacy and legal guidelines when processing personal emails.
