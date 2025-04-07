import os
import re
import pandas as pd
import fitz
import ast


class PDFErrorAnnotator:
    """
    A class to annotate PDFs with error highlights based on paragraph and word-level data.
    """
    
    # Colors for annotations
    EXACT_MATCH_COLOR = (1, 1, 0)       # Yellow for exact matches
    POTENTIAL_ERROR_COLOR = (1, 0.7, 0.7)  # Light red for potential errors
    
    def __init__(self, buffer=5, line_tolerance=2):
        """
        Initialize the PDF Error Annotator.
        
        Args:
            buffer (int): Buffer to extend paragraph boundaries
            line_tolerance (int): Tolerance for considering words on the same line
        """
        self.BUFFER = buffer
        self.LINE_TOLERANCE = line_tolerance
        self.df_para = None
        self.df_words = None
        self.df_results = None
    
    def load_data(self, para_file, word_file):
        """
        Load paragraph and word level data from Excel files.
        
        Args:
            para_file (str): Path to paragraph level Excel file
            word_file (str): Path to word level Excel file
        """
        # Load paragraph data
        self.df_para = pd.read_excel(para_file)
        self.df_para["Clipbounds"] = self.df_para["Clipbounds"].apply(self.safe_literal_eval)
        self.df_para["Clipbounds"] = self.df_para["Clipbounds"].apply(
            lambda x: [x[0], x[3], x[2], x[1]] if x else None
        )  # Convert to correct format
        
        # Load word data
        self.df_words = pd.read_excel(word_file)
        self.df_words["Clipbounds"] = self.df_words["Clipbounds"].apply(self.safe_literal_eval)
        
        # Create results dataframe
        self.df_results = self.df_para.copy()
        # Initialize the Annotation_bbox column with empty lists
        self.df_results["Annotation_bbox"] = [[] for _ in range(len(self.df_results))]
    
    @staticmethod
    def safe_literal_eval(val):
        """
        Safely evaluate literal string to Python object.
        
        Args:
            val: Value to evaluate
            
        Returns:
            Evaluated value or original value if evaluation fails
        """
        if isinstance(val, str):
            try:
                return ast.literal_eval(val)
            except:
                return val
        return val
    
    @staticmethod
    def is_within(bbox, container):
        """
        Check if a bounding box is within a container box.
        
        Args:
            bbox: Bounding box to check
            container: Container bounding box
            
        Returns:
            bool: True if bbox is within container
        """
        if not bbox or not container:
            return False
        return (
            bbox[0] >= container[0] and bbox[2] <= container[2] and
            bbox[1] >= container[1] and bbox[3] <= container[3]
        )
    
    def parse_error_phrases(self, error_phrases):
        """
        Parse error phrases from string or list format.
        
        Args:
            error_phrases: Error phrases as string or list
            
        Returns:
            list: List of error phrases
        """
        error_list = []
        if isinstance(error_phrases, str):
            try:
                error_list = ast.literal_eval(error_phrases)
            except:
                # If evaluation fails, treat as a single string
                error_list = [error_phrases]
        elif isinstance(error_phrases, list):
            error_list = error_phrases
        
        return error_list
    
    def process_file(self, filename, folder_path, output_folder):
        """
        Process a single PDF file and create annotated version.
        
        Args:
            filename (str): Name of the PDF file
            folder_path (str): Path to the folder containing PDF files
            output_folder (str): Path to save annotated PDFs
            
        Returns:
            bool: True if processing was successful
        """
        file_path = os.path.join(folder_path, filename)
        
        try:
            pdf_document = fitz.open(file_path)
        except Exception as e:
            print(f"Error opening file {filename}: {e}")
            return False
        
        df_para_file = self.df_para[self.df_para["Asset Name"] == filename]
        
        # Process each paragraph in the file
        for idx, para_row in df_para_file.iterrows():
            self.process_paragraph(pdf_document, para_row, idx, filename)
        
        try:
            annotated_pdf_path = os.path.join(output_folder, filename)
            pdf_document.save(annotated_pdf_path)
            print(f"Annotated PDF saved: {annotated_pdf_path}")
            pdf_document.close()
            return True
        except Exception as e:
            print(f"Error saving PDF {filename}: {e}")
            pdf_document.close()
            return False
    
    def process_paragraph(self, pdf_document, para_row, idx, filename):
        """
        Process a single paragraph for error highlighting.
        
        Args:
            pdf_document: PyMuPDF document object
            para_row: DataFrame row containing paragraph data
            idx: Index of the paragraph in the results DataFrame
            filename: Name of the PDF file
        """
        page_index = para_row["Page Number"] - 1
        para_bbox = para_row["Clipbounds"]
        error_phrases = para_row["error_phrase"]
        
        # Skip if no clipbounds or error phrases
        if not para_bbox or pd.isna(error_phrases):
            return
            
        # Initialize empty annotation bboxes list for this paragraph
        annotation_bboxes = []
        
        try:
            pdf_page = pdf_document[page_index]
        except IndexError:
            print(f"Page index {page_index} out of range for document {filename}")
            return
            
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
            para_bbox[0] - self.BUFFER,
            para_bbox[1] - self.BUFFER,
            para_bbox[2] + self.BUFFER,
            para_bbox[3] + self.BUFFER
        ]
        
        # Parse error phrases
        error_list = self.parse_error_phrases(error_phrases)
        
        if not error_list:
            return
            
        # Get all words in this paragraph
        matching_rows = self.df_words[
            (self.df_words["File Name"] == filename) &
            (self.df_words["Page Number"] == page_index + 1)
        ]
        
        matching_rows = matching_rows[
            matching_rows["Clipbounds"].apply(lambda b: self.is_within(b, para_bbox))
        ].reset_index(drop=True)
        
        if matching_rows.empty:
            # No words found within paragraph bbox, use paragraph bbox as a fallback
            self.highlight_paragraph(pdf_page, para_bbox, error_list, annotation_bboxes)
        else:
            # Process word matches
            self.process_word_matches(pdf_page, matching_rows, error_list, para_bbox, annotation_bboxes)
        
        # Store annotation bboxes in results DataFrame
        self.df_results.at[idx, "Annotation_bbox"] = annotation_bboxes
    
    def highlight_paragraph(self, pdf_page, para_bbox, error_list, annotation_bboxes):
        """
        Highlight the entire paragraph as a potential error.
        
        Args:
            pdf_page: PyMuPDF page object
            para_bbox: Paragraph bounding box
            error_list: List of error phrases
            annotation_bboxes: List to store annotation bounding boxes
        """
        annotation_bboxes.append(para_bbox)
        para_rect = fitz.Rect(*para_bbox)
        highlight = pdf_page.add_highlight_annot(para_rect)
        highlight.set_colors(stroke=self.POTENTIAL_ERROR_COLOR)
        highlight.update()
        highlight.set_info({"content": f"Potential Errors: {', '.join(error_list)}"})
    
    def process_word_matches(self, pdf_page, matching_rows, error_list, para_bbox, annotation_bboxes):
        """
        Process word-level matches for error phrases.
        
        Args:
            pdf_page: PyMuPDF page object
            matching_rows: DataFrame rows with matching words
            error_list: List of error phrases
            para_bbox: Paragraph bounding box
            annotation_bboxes: List to store annotation bounding boxes
        """
        # Create a continuous string of all words for exact matching
        all_words_string = " ".join(matching_rows["Content"].astype(str))
        
        matches_found = False
        
        for error in error_list:
            if not error or not isinstance(error, str):
                continue
                
            # Try direct string matching
            if error in all_words_string:
                matches_found = True
                
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
                    self.highlight_matched_words(
                        pdf_page, matching_rows, match_word_indices, error, annotation_bboxes
                    )
        
        # If no matches were found for any error phrase, highlight the entire paragraph
        if not matches_found:
            self.highlight_paragraph(pdf_page, para_bbox, error_list, annotation_bboxes)
    
    def highlight_matched_words(self, pdf_page, matching_rows, match_word_indices, error, annotation_bboxes):
        """
        Highlight matched words on the PDF page.
        
        Args:
            pdf_page: PyMuPDF page object
            matching_rows: DataFrame rows with matching words
            match_word_indices: Indices of matching words
            error: Error phrase being highlighted
            annotation_bboxes: List to store annotation bounding boxes
        """
        phrase_bboxes = matching_rows.loc[match_word_indices, "Clipbounds"].tolist()
        
        # Group bounding boxes by line
        phrase_bboxes.sort(key=lambda b: b[1])  # Sort by y-coordinate
        
        # Group words by line based on vertical position
        line_groups = [[phrase_bboxes[0]]]
        for k in range(1, len(phrase_bboxes)):
            if abs(phrase_bboxes[k][1] - phrase_bboxes[k-1][1]) <= self.LINE_TOLERANCE:
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
            highlight.set_colors(stroke=self.EXACT_MATCH_COLOR)
            highlight.update()
            highlight.set_info({"content": f"Error Phrase: {error}"})
    
    def save_results(self, output_file):
        """
        Save annotation results to Excel file.
        
        Args:
            output_file (str): Path to save results Excel file
            
        Returns:
            bool: True if saving was successful
        """
        try:
            self.df_results.to_excel(output_file, index=False)
            print(f"Annotation results saved to {output_file}")
            return True
        except Exception as e:
            print(f"Error saving annotation results: {e}")
            return False
    
    def verify_annotations(self):
        """
        Verify all rows with error phrases have annotation bboxes.
        
        Returns:
            DataFrame: Rows with error phrases but no annotations
        """
        empty_annotations = self.df_results[
            (self.df_results["error_phrase"].notna()) & 
            (self.df_results["Annotation_bbox"].apply(lambda x: len(x) == 0))
        ]
        
        if not empty_annotations.empty:
            print(f"Warning: {len(empty_annotations)} rows still have empty annotation bounding boxes")
        
        return empty_annotations
    
    def process_all_files(self, folder_path, output_folder):
        """
        Process all PDF files in the dataset.
        
        Args:
            folder_path (str): Path to the folder containing PDF files
            output_folder (str): Path to save annotated PDFs
        """
        # Process each unique file
        for filename in self.df_para["Asset Name"].unique():
            self.process_file(filename, folder_path, output_folder)
