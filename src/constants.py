from pathlib import Path

# Title Normalization
SMALL_WORDS = {'and', 'or', 'the', 'of', 'in', 'on', 'a', 'an', 'to', 'at'}
VALID_TITLE_REGEX = r"^[a-zA-Z0-9\s\-\'\,\"\.\?\!]+$"

COVERS_FOLDER = Path(r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\קבצי ספרים מוכנים להעלאה לאמזון\00 תמונות של כריכות ספרים לאמזון")
BOOK_TRACKER_EXCEL_FILE_PATH = Path(r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\טבלה מרכזת ספרים דיגיטליים.xlsx")
PROMPT_PATH = Path(r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\python\PDF-Book-Automizer\src\gemini\gemini_con_prompt.txt")