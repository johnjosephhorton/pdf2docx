import os
import argparse
from pdf2image import convert_from_path

from PIL import Image
from docx import Document
from docx.shared import Inches


def add_images_to_docx(image_dir, docx_path):
    # Get a list of all image files in the directory
    image_files = [f for f in os.listdir(image_dir) if f.endswith(".png")]

    # Sort the list of image files in numerical order
    image_files.sort(key=lambda x: int(os.path.splitext(x)[0]))

    # Create a new Word document
    document = Document()

    # Add each image to the Word document
    for image_file in image_files:
        # Open the image file
        image_path = os.path.join(image_dir, image_file)
        image = Image.open(image_path)

        # Add the image to the Word document
        document.add_picture(image_path, width=Inches(6))

        document.add_page_break()

    # Save the Word document
    document.save(docx_path)


def pdf_to_png(pdf_path):
    # Convert PDF to PNG images
    images = convert_from_path(pdf_path)

    # Create a directory to store the PNG images
    png_dir = os.path.splitext(pdf_path)[0]
    os.makedirs(png_dir, exist_ok=True)

    # Save the PNG images to the directory
    for i, image in enumerate(images):
        image_path = os.path.join(png_dir, f"{i+1}.png")
        image.save(image_path, "PNG")


def main():
    # Parse the command-line arguments
    parser = argparse.ArgumentParser(description="Convert a PDF file to PNG images")
    parser.add_argument("pdf_path", help="Path to the PDF file")
    args = parser.parse_args()

    # Convert the PDF file to PNG images
    pdf_to_png(args.pdf_path)

    # Convert the PNG images to a Word document
    image_path = os.path.splitext(args.pdf_path)[0]
    word_doc = os.path.splitext(args.pdf_path)[0] + ".docx"
    add_images_to_docx(image_path, word_doc)


if __name__ == "__main__":
    main()
