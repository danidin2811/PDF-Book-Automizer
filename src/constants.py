from pathlib import Path

# Title Normalization
SMALL_WORDS = {'and', 'or', 'the', 'of', 'in', 'on', 'a', 'an', 'to', 'at'}
VALID_TITLE_REGEX = r"^[a-zA-Z0-9\s\-\'\,\"\.\?\!]+$"

COVERS_FOLDER = Path(r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\קבצי ספרים מוכנים להעלאה לאמזון\00 תמונות של כריכות ספרים לאמזון")
BOOK_TRACKER_EXCEL_FILE_PATH = Path(r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\טבלה מרכזת ספרים דיגיטליים.xlsx")
PROMPT_PATH = Path(r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\python\PDF-Book-Automizer\src\gemini\gemini_con_prompt.txt")
READY_TO_UPLOAD_TO_AMAZON_FOLDER = Path(r"R:\Documents\001אתר האינטרנט ופרויקטים דיגיטליים\הכנת כתבי עת לאתר\הכנת ספרים לאתר\קבצי ספרים מוכנים להעלאה לאמזון")
FOLDER_NAME_COL: str = 'L'

BASE_DESIGN_TEMPLATE = {
    "phoneFlipShortcutButton": True,
    "updateURLForPage": False,
    "appLogoLinkURL": "https://ybz.org.il/",
    "FlipStyle": "Switch",
    "RightToLeft": True,
    "loadingCaption": "טוען את הספר",
    "loadingCaptionColor": "0xffffff",
    "restorePageVisible": True,
    "isAccessibilityButtonVisible": True,
    "ZoomMapVisible": True,
    "searchKeywordFontColor": "0xffb000",
    "searchHightlightColor": "0xfdc606",
    "ShareButtonVisible": False,
    "BookMarkButtonVisible": True,
    "HomeButtonVisible": True,
    "AutoPlayButtonVisible": True,
    "SelectTextButtonVisible": True,
    "MagnifierButtonVisible": True,
    "InstructionsButtonVisible": True,
    "showInstructionOnStart": True,
    "LeftShadowWidth": 50,
    "RightShadowWidth": 20,
    "ShowTopLeftShadow": False,
    "restorePageDuration": "1",
}

ENGLISH_DESIGN_TEMPLATE = {
    **BASE_DESIGN_TEMPLATE,
    "RightToLeft": False,
}

FREE_HEBREW_DESIGN_TEMPLATE = {
    **BASE_DESIGN_TEMPLATE,
    "ShareButtonVisible": True
}