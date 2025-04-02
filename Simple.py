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
annotation_results_file = "Annotation_Results.xlsx"
os.makedirs(annotated_pdf_folder, exist_ok=True)

# Load data
df_para = pd.read_excel(para_output_file)
df_para["Clipbounds"] = df_para["Clipbounds"].apply(safe_literal_eval)
df_para["Clipbounds"] = df_para["Clipbounds"].apply(lambda x: [x[0], x[3], x[2], x[1]])  # Convert to correct format

df_words = pd.read_excel(word_output_file)
df_words["Clipbounds"] = df_words["Clipbounds"].apply(safe_literal_eval)

# Create a new dataframe for results
df_results = df_para.copy()
# Initialize the Annotation_bbox column with empty lists
df_results["Annotation_bbox"] = [[] for _ in range(len(df_results))]

# Config
BUFFER = 4
LINE_TOLERANCE = 2  # Tolerance for considering words on the same line

# Colors
EXACT_MATCH_COLOR = (1, 1, 0)       # Yellow for exact matches
POTENTIAL_ERROR_COLOR = (1, 0.7, 0.7)  # Light red for potential errors

for filename in df_para["Asset Name"].unique():
    file_path = os.path.join(folder_path, filename)
    
    try:
        pdf_document = fitz.open(file_path)
    except Exception as e:
        print(f"Error opening file {filename}: {e}")
        continue
    
    df_para_file = df_para[df_para["Asset Name"] == filename]

    for idx, para_row in df_para_file.iterrows():
        page_index = para_row["Page Number"] - 1
        para_bbox = para_row["Clipbounds"]
        error_phrases = para_row["error_phrase"]
        annotation_bboxes = []  # Store all bounding boxes for this paragraph

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

        # Add buffer to paragraph Clipbounds
        para_bbox = [
            para_bbox[0] - BUFFER,
            para_bbox[1] - BUFFER,
            para_bbox[2] + BUFFER,
            para_bbox[3] + BUFFER
        ]
        
        # Parse error phrases
        error_list = []
        if isinstance(error_phrases, str):
            try:
                error_list = ast.literal_eval(error_phrases)
            except:
                error_list = [error_phrases]
        elif isinstance(error_phrases, list):
            error_list = error_phrases
        
        if not error_list:
            continue
            
        # Get all words in this paragraph
        matching_rows = df_words[
            (df_words["File Name"] == filename) &
            (df_words["Page Number"] == page_index + 1)
        ]

        def is_within(bbox, container):
            return (
                bbox[0] >= container[0] and bbox[2] <= container[2] and
                bbox[1] >= container[1] and bbox[3] <= container[3]
            )

        matching_rows = matching_rows[matching_rows["Clipbounds"].apply(lambda b: is_within(b, para_bbox))].reset_index(drop=True)

        if matching_rows.empty:
            continue

        # Create a continuous string of all words for exact matching
        all_words_string = " ".join(matching_rows["Content"].astype(str))
        
        # Track if any exact matches were found
        exact_match_found = False
        
        for error in error_list:
            if not error:
                continue

            # Try direct string matching
            if error in all_words_string:
                exact_match_found = True
                
                # Find word indices that correspond to the exact match
                match_start = all_words_string.find(error)
                match_end = match_start + len(error)
                
                # Find word indices that make up this exact match
                current_position = 0
                match_word_indices = []
                
                for word_idx, word in enumerate(matching_rows["Content"]):
                    word_str = str(word)
                    word_start = current_position
                    word_end = word_start + len(word_str)
                    
                    # Check if this word overlaps with the match
                    if (word_start <= match_end) and (word_end >= match_start):
                        match_word_indices.append(word_idx)
                    
                    current_position = word_end + 1  # +1 for space
                
                # Get bounding boxes for the matched words
                if match_word_indices:
                    phrase_bboxes = matching_rows.loc[match_word_indices, "Clipbounds"].tolist()
                    
                    # Group bounding boxes by line
                    phrase_bboxes.sort(key=lambda b: b[1])  # Sort by y-coordinate
                    
                    # Group words by line based on vertical position
                    line_groups = [[phrase_bboxes[0]]]
                    for k in range(1, len(phrase_bboxes)):
                        if abs(phrase_bboxes[k][1] - phrase_bboxes[k-1][1]) <= LINE_TOLERANCE:
                            # Same line
                            line_groups[-1].append(phrase_bboxes[k])
                        else:
                            # New line
                            line_groups.append([phrase_bboxes[k]])
                    
                    # Highlight each line separately
                    for line in line_groups:
                        x0 = min(b[0] for b in line)
                        y0 = min(b[1] for b in line)
                        x1 = max(b[2] for b in line)
                        y1 = max(b[3] for b in line)
                        
                        line_bbox = [x0, y0, x1, y1]
                        annotation_bboxes.append(line_bbox)
                        
                        phrase_rect = fitz.Rect(x0, y0, x1, y1)
                        highlight = pdf_page.add_highlight_annot(phrase_rect)
                        highlight.set_colors(stroke=EXACT_MATCH_COLOR)  # Yellow for exact matches
                        highlight.update()
                        highlight.set_info({"content": f"Error Phrase: {error}"})
        
        # If no exact matches were found, highlight the entire paragraph
        if not exact_match_found:
            annotation_bboxes.append(para_bbox)
            para_rect = fitz.Rect(*para_bbox)
            highlight = pdf_page.add_highlight_annot(para_rect)
            highlight.set_colors(stroke=POTENTIAL_ERROR_COLOR)  # Light red for potential errors
            highlight.update()
            highlight.set_info({"content": f"Potential Errors: {', '.join(error_list)}"})
        
        # Update the results dataframe with annotation bounding boxes
        results_idx = df_results.index[(df_results["Asset Name"] == filename) & 
                                       (df_results["Page Number"] == para_row["Page Number"]) &
                                       (df_results["Atom ID"] == para_row["Atom ID"])].tolist()
        
        if results_idx:
            df_results.at[results_idx[0], "Annotation_bbox"] = annotation_bboxes

    try:
        annotated_pdf_path = os.path.join(annotated_pdf_folder, filename)
        pdf_document.save(annotated_pdf_path)
        print(f"Annotated PDF saved: {annotated_pdf_path}")
    except Exception as e:
        print(f"Error saving PDF {filename}: {e}")
    
    pdf_document.close()

# Save the results to Excel
try:
    df_results.to_excel(annotation_results_file, index=False)
    print(f"Annotation results saved to {annotation_results_file}")
except Exception as e:
    print(f"Error saving annotation results: {e}")
