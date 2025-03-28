import os
import re
import pandas as pd
import fitz
import ast

def safe_literal_eval(val):
    if isinstance(val, str):
        try:
            return ast.literal_eval(val)
        except:
            return val
    return val

# Paths
folder_path = "Test_Assets"
para_output_file = "PARA_LEVEL.xlsx"
word_output_file = "WORD_LEVEL.xlsx"
annotated_pdf_folder = "Errors_Highlighted"
os.makedirs(annotated_pdf_folder, exist_ok=True)

# Load data
df_para = pd.read_excel(para_output_file)
df_para["Bounding Box"] = df_para["Bounding Box"].apply(safe_literal_eval)
df_para["Bounding Box"] = df_para["Bounding Box"].apply(lambda x: [x[0], x[3], x[2], x[1]])  # Convert to correct format

df_words = pd.read_excel(word_output_file)
df_words["Bounding Box"] = df_words["Bounding Box"].apply(safe_literal_eval)

# Config
BUFFER = 5
LINE_TOLERANCE = 2

for filename in df_para["File Name"].unique():
    file_path = os.path.join(folder_path, filename)
    
    try:
        pdf_document = fitz.open(file_path)
    except Exception as e:
        print(f"Error opening file {filename}: {e}")
        continue
    
    df_para_file = df_para[df_para["File Name"] == filename]

    for _, para_row in df_para_file.iterrows():
        page_index = para_row["Page Number"] - 1
        para_bbox = para_row["Bounding Box"]
        error_phrases = para_row["error_phrase"]

        if not para_bbox:
            continue

        pdf_page = pdf_document[page_index]
        page_height = pdf_page.rect.height
        # Convert adobe to fitz coordinate system
        para_bbox = [
            para_bbox[0],  
            page_height - para_bbox[1],  
            para_bbox[2],  
            page_height - para_bbox[3]   
        ]

        para_bbox = [
            para_bbox[0] - BUFFER,
            para_bbox[1] - BUFFER,
            para_bbox[2] + BUFFER,
            para_bbox[3] + BUFFER
        ]
        
        rect = fitz.Rect(*para_bbox)
        pdf_page = pdf_document[page_index]

        # Parse error phrases
        error_list = []
        if isinstance(error_phrases, str):
            try:
                error_list = ast.literal_eval(error_phrases)
            except:
                error_list = [error_phrases]
        elif isinstance(error_phrases, list):
            error_list = error_phrases

        matching_rows = df_words[
            (df_words["File Name"] == filename) &
            (df_words["Page Number"] == page_index + 1)
        ]

        def is_within(bbox, container):
            return (
                bbox[0] >= container[0] and bbox[2] <= container[2] and
                bbox[1] >= container[1] and bbox[3] <= container[3]
            )

        matching_rows = matching_rows[matching_rows["Bounding Box"].apply(lambda b: is_within(b, para_bbox))].reset_index(drop=True)

        if matching_rows.empty:
            continue

        matching_rows["Spans"] = matching_rows["Spans"].apply(safe_literal_eval)
        matching_rows["Next Word Span"] = matching_rows["Next Word Span"].apply(safe_literal_eval)

        # New method: Create a continuous string of all words
        all_words_string = " ".join(matching_rows["Content"].astype(str))
        
        # Track used indices to prevent duplicate highlights
        used_indices = set()
        
        for error in error_list:
            if not error:
                continue

            # First, try direct string matching
            direct_matches = [m.start() for m in re.finditer(re.escape(error), all_words_string)]
            
            if direct_matches:
                # For the first direct match, find the corresponding words and highlight
                match_start = direct_matches[0]
                # Find the word indices for this match
                current_position = 0
                match_word_indices = []
                
                for idx, word in enumerate(matching_rows["Content"]):
                    word_len = len(str(word))
                    # Modified to capture word even if it's partially matching within the exact phrase
                    if current_position >= match_start and current_position < match_start + len(error):
                        if idx not in used_indices:
                            match_word_indices.append(idx)
                    elif (current_position < match_start + len(error)) and (current_position + word_len >= match_start):
                        # Capture words that overlap with the match
                        if idx not in used_indices:
                            match_word_indices.append(idx)
                    current_position += word_len + 1  # +1 for space
                    
                    if current_position > match_start + len(error):
                        break
                
                # Highlight the matched words
                if match_word_indices:
                    phrase_bboxes = matching_rows.loc[match_word_indices, "Bounding Box"].tolist()
                    phrase_bboxes.sort(key=lambda b: b[1])
                    
                    line_groups = [[phrase_bboxes[0]]]
                    for k in range(1, len(phrase_bboxes)):
                        if abs(phrase_bboxes[k][1] - phrase_bboxes[k - 1][1]) <= LINE_TOLERANCE:
                            line_groups[-1].append(phrase_bboxes[k])
                        else:
                            line_groups.append([phrase_bboxes[k]])

                    for line in line_groups:
                        x0, y0 = min(b[0] for b in line), min(b[1] for b in line)
                        x1, y1 = max(b[2] for b in line), max(b[3] for b in line)
                        phrase_rect = fitz.Rect(x0, y0, x1, y1)

                        highlight = pdf_page.add_highlight_annot(phrase_rect)
                        highlight.set_info({"content": f"Error Phrase: {error}"})
                    
                    # Mark these indices as used
                    used_indices.update(match_word_indices)
                    continue

            # Partial matching logic
            all_matches = []
            for start in range(len(error.split())):
                sub_error = " ".join(error.split()[start:])
                clean_error = sub_error.replace(" ", "")
                i = 0

                while i < len(matching_rows):
                    if i in used_indices:
                        i += 1
                        continue

                    built = str(matching_rows.loc[i, "Content"]).replace(" ", "")
                    temp_sequence = [i]
                    j = i + 1

                    while j < len(matching_rows) and clean_error.startswith(built):
                        prev_span = matching_rows.loc[j - 1, "Next Word Span"]
                        curr_span = matching_rows.loc[j, "Spans"]

                        if prev_span != curr_span:
                            break

                        next_content = str(matching_rows.loc[j, "Content"]).replace(" ", "")
                        built += next_content

                        if clean_error.startswith(built):
                            temp_sequence.append(j)
                            j += 1
                        else:
                            break

                    match_text = "".join([str(matching_rows.loc[x, "Content"]) for x in temp_sequence])
                    match_clean = match_text.replace(" ", "")

                    if clean_error.startswith(match_clean) and len(match_clean) > 0:
                        all_matches.append({
                            "sequence": temp_sequence.copy(),
                            "text": match_text,
                            "is_full_match": match_clean == clean_error,
                            "length": len(match_clean),
                            "sub_error": sub_error,
                            "original_error": error
                        })
                        i = j
                    else:
                        i += 1

            # Prefer full matches first
            full_matches = [m for m in all_matches if m["is_full_match"]]
            if full_matches:
                # If full matches exist, highlight only the first full match
                match_to_highlight = full_matches[0]
            else:
                # If no full matches, find the longest partial match
                if not all_matches:
                    continue
                
                # Find the maximum length of partial matches
                max_len = max(m["length"] for m in all_matches)
                
                # Select only the first match with the maximum length
                match_to_highlight = next(m for m in all_matches if m["length"] == max_len)

            # Check if the match uses any already used indices
            if not any(idx in used_indices for idx in match_to_highlight["sequence"]):
                phrase_bboxes = matching_rows.loc[match_to_highlight["sequence"], "Bounding Box"].tolist()
                phrase_bboxes.sort(key=lambda b: b[1])
                line_groups = [[phrase_bboxes[0]]]

                for k in range(1, len(phrase_bboxes)):
                    if abs(phrase_bboxes[k][1] - phrase_bboxes[k - 1][1]) <= LINE_TOLERANCE:
                        line_groups[-1].append(phrase_bboxes[k])
                    else:
                        line_groups.append([phrase_bboxes[k]])

                for line in line_groups:
                    x0, y0 = min(b[0] for b in line), min(b[1] for b in line)
                    x1, y1 = max(b[2] for b in line), max(b[3] for b in line)
                    phrase_rect = fitz.Rect(x0, y0, x1, y1)

                    highlight = pdf_page.add_highlight_annot(phrase_rect)
                    highlight.set_info({"content": f"Error Phrase: {match_to_highlight['original_error']}"})
                
                # Mark these indices as used
                used_indices.update(match_to_highlight["sequence"])

    try:
        annotated_pdf_path = os.path.join(annotated_pdf_folder, filename)
        pdf_document.save(annotated_pdf_path)
        print(f"Annotated PDF saved: {annotated_pdf_path}")
    except Exception as e:
        print(f"Error saving PDF {filename}: {e}")
    
    pdf_document.close()
