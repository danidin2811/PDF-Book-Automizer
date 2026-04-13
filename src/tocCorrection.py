import re

# Updated Regex patterns
pattern_type_1 = r"^\s*–\s*([\u0590-\u05FF\s:;,.!?\-]+)\s+(\d+)\s*פרק\s*$"  
# Type 1: Dash + Hebrew + Number + "פרק"

pattern_type_2 = r"^\s*–\s*([\u0590-\u05FF\s:;,.!?\-]+)\s+(\d+\.\d+)\s*$"  
# Type 2: Dash + Hebrew + Decimal Number

pattern_type_3 = r"^\s*–\s*([\u0590-\u05FF\s:;,.!?\-]+)(\d+(?:\.\d+)+)$"  
# Type 3: Dash + Hebrew + Section Number (any number of parts)

pattern_type_4 = r"^\s*\)\s*–\s*([\u0590-\u05FF\s\(\)\-:;,.!?\"]+)(\d+(?:\.\d+)+)$"  
# Type 4: Starts with a ')' then a dash, then Hebrew text (which may include parentheses), then the section number

pattern_type_5 = r"^\s*'\s*–\s*([\u0590-\u05FF\s:;,.!?\-\(\)\"']+)(\d+(?:\.\d+)+)$"

pattern_type_6 = r"^\s*\?\s*–\s*([\u0590-\u05FF\s:;,.!?\-\(\)\"']+)(\d+(?:\.\d+)+)$"

pattern_type_7 = r"^\s*([\u0590-\u05FF\s\"׳:]+?)\s*פרק\s*([\u0590-\u05FF\s\"׳:]+)$"

def correct_line(line):
    match_type_1 = re.match(pattern_type_1, line)
    if match_type_1:
        hebrew_text, number = match_type_1.groups()
        corrected_line = f"פרק {number} – {hebrew_text.strip()}"
        return corrected_line

    match_type_2 = re.match(pattern_type_2, line)
    if match_type_2:
        hebrew_text, number = match_type_2.groups()
        corrected_line = f"{number} – {hebrew_text.strip()}"
        return corrected_line

    match_type_3 = re.match(pattern_type_3, line)
    if match_type_3:
        hebrew_text, section_number = match_type_3.groups()
        corrected_line = f"{section_number} – {hebrew_text.strip()}"
        return corrected_line

    match_type_4 = re.match(pattern_type_4, line)
    if match_type_4:
        hebrew_text, section_number = match_type_4.groups()
        corrected_line = f"{section_number} – {hebrew_text.strip()})"
        return corrected_line
        
    match_type_5 = re.match(pattern_type_5, line)
    if match_type_5:
        hebrew_text, section_number = match_type_5.groups()
        corrected_line = f"{section_number} – {hebrew_text.strip()}"
        return corrected_line
        
    match_type_6 = re.match(pattern_type_6, line)
    if match_type_6:
        hebrew_text, section_number = match_type_6.groups()
        # Append a question mark at the end.
        corrected_line = f"{section_number} – {hebrew_text.strip()}?"
        return corrected_line
        
    match_type_7 = re.match(pattern_type_7, line)
    if match_type_7:
        title, chapter = match_type_7.groups()
        corrected_line = f"פרק {chapter.strip()}: {title.strip()}"
        return corrected_line

    # If the line doesn't match any type, return it as is
    return line