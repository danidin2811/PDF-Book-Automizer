from utils.input_output_tools import print_red


def load_gemini_prompt() -> str:
    """Reads the transcription prompt from the resources folder."""

    prompt_path = "gemini_con_prompt.txt"

    print(prompt_path)
    try:
        with open(prompt_path, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        print_red(f"Warning: Prompt file not found at {prompt_path}")
        return "Please transcribe the attached Table of Contents to CSV."

def handle_gemini_toc_transcription(source_folder, con_file_path):
    """
        Copies the transcription prompt to clipboard and opens the Gemini URL.
    """

    import webbrowser
    import pyperclip
    import os

    gemini_prompt = load_gemini_prompt()
    pyperclip.copy(gemini_prompt) # Copy prompt to clipboard for easy pasting

    print(f"Prompt copied to clipboard. Opening Gemini...")
    print(f"File to upload: {con_file_path}")

    # Open Gemini in the default browser
    webbrowser.open("https://gemini.google.com/app")

    # Open the folder so you can drag the file easily
    os.startfile(source_folder)