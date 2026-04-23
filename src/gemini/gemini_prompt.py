from utils.input_output_tools import print_red


def load_gemini_prompt() -> str:
    """Reads the transcription prompt from the resources folder."""

    from src.constants import PROMPT_PATH

    prompt_path = PROMPT_PATH

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

    from utils.input_output_tools import wait_for_ready_signal, ask_offset

    offset = ask_offset()
    raw_prompt = load_gemini_prompt()
    formatted_prompt = raw_prompt.format(offset=offset)
    pyperclip.copy(formatted_prompt) # Copy prompt to clipboard for easy pasting

    print(f"Prompt copied to clipboard. Opening Gemini...")
    print(f"File to upload: {con_file_path}")

    # Open Gemini in the default browser
    webbrowser.open("https://gemini.google.com/app")

    # Open the folder so you can drag the file easily
    os.startfile(source_folder)

    instructions = (
        f"\nACTION REQUIRED: Gemini Transcription\n"
        f"--------------------------------------\n"
        f"1. A new Gemini chat has been opened in your browser\n"
        f"2. Drag the file to the chat: {os.path.basename(con_file_path)}\n"
        f"3. Paste the prompt (already copied to your clipboard)\n"
        f"4. Save the AI-generated CSV as 'toc.csv' in the book folder\n"
        f"--------------------------------------\n"
        f"Press Enter once 'toc.csv' is saved and you are ready to proceed: "
    )

    wait_for_ready_signal(instructions)

if __name__ == "__main__":
    source_folder = r"C:\Users\system1\Desktop\קבצי ספרים בעבודה\studies_in_the_history_of_eretz_israel"
    con_file_path = r"C:\Users\system1\Desktop\קבצי ספרים בעבודה\studies_in_the_history_of_eretz_israel\studies_in_the_history_of_eretz_israel_con.pdf"
    handle_gemini_toc_transcription(source_folder,con_file_path)