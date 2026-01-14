import argparse
import sys
import os
from .core import SubtitlesProcessor 

def main():
    parser = argparse.ArgumentParser(
        description="CapSlide: A tool to convert subtitles (.json/.txt) into professional PPT slides."
    )
    
    # 1. Positional Argument
    parser.add_argument("subtitles", help="Path to the input subtitle file (.json/.txt).")
    
    # 2. Required Arguments
    parser.add_argument("-t", "--template", help="Path to the template PPTX file.", required=True)
    
    # 3. Optional Arguments
    parser.add_argument("-o", "--output", help="Path for the generated PPTX file (default: output.pptx).", default="output.pptx")
    parser.add_argument("-p", "--placeholder", help="The text placeholder in the template to be replaced.", default="subtitle")
    parser.add_argument("-n", "--template_slide_page_number", 
                        help="Slide page number to use (1-based). Use 0 for the last slide (default: 0).", 
                        default=0, type=int)
    
    parser.add_argument("-i", "--ignore_marks", help="Exclude punctuation marks from the slides.", action="store_true")
    parser.add_argument("-v", "--verbose", help="Display detailed processing logs.", action="store_true")
    
    args = parser.parse_args()

    # --- Robustness Checks ---
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found.")
        sys.exit(1)
    
    if not os.path.exists(args.template):
        print(f"Error: Template file '{args.template}' not found.")
        sys.exit(1)

    try:
        # Initializing the processor
        processor = SubtitlesProcessor(
            output_path=args.output,
            template_path=args.template,
            template_slide_page_number=args.template_slide_page_number,
            placeholder=args.placeholder, 
            ignore_masks=args.ignore_marks,
            verbose=args.verbose
        )
        
        print(f"Status: Processing '{args.subtitles}'...")
        
        # 4. Trigger actual processing logic
        # Assuming the method is named 'process' or 'run'
        processor.append_slides_from_file(args.subtitles)
        
        print(f"Success: PPT generated at '{args.outsubtitlesput}'")

    except Exception as e:
        print(f"An error occurred during processing: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()