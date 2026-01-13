import argparse
from .core import SubtitlesProcessor 

def main():
    parser = argparse.ArgumentParser(description="CapSlide: Convert subtitles into PPT slides.")
    
    # Argument definitions with standardized English descriptions
    parser.add_argument("input", help="Path to the input subtitle file (.json/.txt).")
    parser.add_argument("-o", "--output", help="Path for the generated PPTX file (default: output.pptx).", default="output.pptx")
    parser.add_argument("-t", "--template", help="Path to the template PPTX file.")
    parser.add_argument("-p", "--placeholder", help="The text placeholder in the template to be replaced.", default="subtitle")
    parser.add_argument("-n", "--template_slide_page_number", 
                        help="The specific slide index in the template to use. Use 0 for the last slide (default: 0).", 
                        default=0, type=int)
    parser.add_argument("-i", "--ignore_marks", help="Exclude punctuation marks from the slides.", action="store_true")
    parser.add_argument("-v", "--verbose", help="Display detailed processing logs.", action="store_true")
    
    
    args = parser.parse_args()
    
    # Initializing the processor with corrected variable mappings
    processor = SubtitlesProcessor(
        output_path=args.output,
        template_path=args.template,
        template_slide_page_number=args.template_slide_page_number,
        placeholder=args.placeholder, 
        ignore_masks=args.ignore_marks,
        verbose=args.verbose
    )
    
    # Console feedback in English
    print(f"Status: Processing file '{args.input}'...")
    
    # Logic to trigger the conversion (Assuming a 'run' or 'process' method exists)
    # processor.process(args.input)
    # print("Success: Transformation complete.")

if __name__ == "__main__":
    main()