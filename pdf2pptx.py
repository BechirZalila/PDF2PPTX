import argparse
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches, Pt

# Function to calculate new slide dimensions while preserving the image's aspect ratio
def calculate_slide_size(img_width, img_height):
    # Standard slide size in inches (can be set as any initial value)
    base_width = 10.0  # for example, 10 inches as the base width
    img_ratio = img_width / img_height
    slide_height = base_width / img_ratio
    return Inches(base_width), Inches(slide_height)

# Main function
def pdf_to_pptx(pdf_file, pptx_file, skip_first, skip_pages):
    # Convert the PDF to images
    pages = convert_from_path(pdf_file)

    # Skip the first page if the option is specified
    if skip_first:
        pages = pages[1:]

    # Convert the page numbers to skip into indices. No risk of 
    # offset if --skip-first is given because the two options 
    # are incompatible and cannot be given simultaneously.
    skip_indices = set(page_num - 1 for page_num in skip_pages)

    # Create a PowerPoint presentation
    prs = Presentation()

    # Add each image as a slide, skipping the specified pages
    for i, page in enumerate(pages):
        if i in skip_indices:
            continue

        # Save each page as an image (to access its size)
        img_path = 'temp_image.png'
        page.save(img_path, 'PNG')

        # Get image size (in pixels)
        img_width, img_height = page.size

        
        # Calculate new slide size based on the image's aspect ratio
        slide_width, slide_height = calculate_slide_size(img_width, img_height)

        # Adjust the slide size to match the image's aspect ratio
        prs.slide_width = slide_width
        prs.slide_height = slide_height

        # Add slide and insert image
        slide_layout = prs.slide_layouts[6]  # Blank slide
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.add_picture(img_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

    # Save the PowerPoint presentation
    prs.save(pptx_file)
    print(f"PPTX presentation created: {pptx_file}")

# Argument parser configuration
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert a PDF to a PowerPoint presentation")
    
    # Add a mutually exclusive group for --skip-first and --skip
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--skip-first", action="store_true", help="Skip the first page of the PDF")
    group.add_argument("--skip", type=str, help="List of pages to skip, separated by commas (e.g., 2,4,5)")

    parser.add_argument("pdf_file", help="Input PDF file name")
    parser.add_argument("pptx_file", help="Output PPTX file name")

    args = parser.parse_args()

    # Convert the --skip argument to a list of integers
    if args.skip:
        skip_pages = list(map(int, args.skip.split(',')))
    else:
        skip_pages = []

    # Call the function with the arguments
    pdf_to_pptx(args.pdf_file, args.pptx_file, args.skip_first, skip_pages)


