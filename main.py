import os
from pdf_error_annotator import PDFErrorAnnotator

def main():
    """
    Main function to run the PDF error annotation process.
    """
    # Paths
    folder_path = "Test_Assets"
    para_output_file = "PARA_LEVEL.xlsx"
    word_output_file = "WORD_LEVEL.xlsx"
    annotated_pdf_folder = "Errors_Highlighted"
    annotation_results_file = "Annotation_Results.xlsx"
    
    # Create output directory if it doesn't exist
    os.makedirs(annotated_pdf_folder, exist_ok=True)
    
    # Initialize annotator
    annotator = PDFErrorAnnotator(buffer=5, line_tolerance=2)
    
    # Load data
    annotator.load_data(para_output_file, word_output_file)
    
    # Process all files
    annotator.process_all_files(folder_path, annotated_pdf_folder)
    
    # Verify annotations
    empty_annotations = annotator.verify_annotations()
    
    # Save results
    annotator.save_results(annotation_results_file)
    
    print("PDF annotation process completed!")

if __name__ == "__main__":
    main()
