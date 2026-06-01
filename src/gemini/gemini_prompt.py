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

def handle_gemini_toc_transcription(source_folder, con_file_path, interface):
    """
    Copies the transcription prompt to clipboard and opens the Gemini URL.
    Supports skipping terminal prompts if GUI context handles the offset value.
    """
    import webbrowser
    import pyperclip
    import os
    from utils.input_output_tools import wait_for_ready_signal

    # --- CHOOSE BETWEEN GUI PARAMETERS AND CLI PROMPTS ---
    offset = ask_offset(interface)

    raw_prompt = load_gemini_prompt()
    formatted_prompt = raw_prompt.format(offset=offset)
    pyperclip.copy(formatted_prompt)  # Copy prompt to clipboard for easy pasting

    interface.print_info(f"Prompt copied to clipboard. Opening Gemini...")
    interface.print_info(f"File to upload: {con_file_path}")

    # Open Gemini in the default browser
    webbrowser.open("https://gemini.google.com/app")

    interface.print_info(f"# Open the folder so you can drag the file easily {source_folder}")
    os.startfile(source_folder)

    # Use the interface layer to determine blocking rules
    if interface.is_gui:
        # If running inside a GUI background thread, leverage a checklist step or signal event
        if hasattr(interface, 'request_step_completion'):
            interface.request_step_completion(
                step_name="Gemini Transcription",
                message=f"Please drop {os.path.basename(con_file_path)} into Gemini, run prompt, and save 'toc.csv' to the book folder."
            )
    else:
        # Standard legacy terminal blocking signal for CLI
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