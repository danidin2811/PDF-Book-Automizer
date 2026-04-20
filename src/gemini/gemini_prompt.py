from pathlib import Path
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
