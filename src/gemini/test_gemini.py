def handle_gemini_toc_transcription(source_folder):
    """
        Copies the transcription prompt to clipboard and opens the Gemini URL.
    """

    import webbrowser
    import pyperclip
    import os
    from src.gemini.gemini_prompt import load_gemini_prompt

    con_file_path = r"C:\Users\system1\Desktop\קבצי ספרים בעבודה\poke\poke_con.pdf"

    gemini_prompt = load_gemini_prompt()
    pyperclip.copy(gemini_prompt) # Copy prompt to clipboard for easy pasting

    print(f"Prompt copied to clipboard. Opening Gemini...")
    print(f"File to upload: {con_file_path}")

    # Open Gemini in the default browser
    webbrowser.open("https://gemini.google.com/app")

    # Open the folder so you can drag the file easily
    os.startfile(source_folder)

if __name__ == "__main__":
    # Define the folder path you want to open
    test_folder = r"C:\Users\system1\Desktop\קבצי ספרים בעבודה\poke"

    # Call the function
    handle_gemini_toc_transcription(test_folder)