import platform
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import pytesseract
from pytesseract import Output
import io
import tkinter as tk
from tkinter import filedialog

def rip_slides(path, file_handle, format_type) -> str:
    print(f"Opening {path}...")
    pres = Presentation(path)

    file_handle.write(get_header(f"Source: {os.path.basename(path)}", 1, format_type))

    for i, slide in enumerate(pres.slides):
        slide_number = i + 1
        print(f"Processing Slide {slide_number}...")

        file_handle.write(f"\n\n## SLIDE {slide_number}\n")

        for shape in slide.shapes:
            if shape.has_text_frame:
                clean_text = shape.text_frame.text
                file_handle.write(f"{clean_text}\n")
            
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image_blob = shape.image.blob
                    image = Image.open(io.BytesIO(image_blob))

                    extracted_text = extract_text_with_confidence(image)
                    
                    if extracted_text.strip():
                        file_handle.write(f"\n\n[IMAGE READ]:\n{extracted_text}\n")
                
                except Exception as e:
                    print(f"Could not read image on slide {slide_number}: {e}")

def get_header(text, level, format_type):
    if format_type == "md":
        return f"\n{'#' * level} {text} \n"
    else:
        return f"\n{'=' * 5} {text} {'=' * 5}\n"

# during Tesseract OCR only return text with high confidence
def extract_text_with_confidence(image, threshold=60):
    data = pytesseract.image_to_data(image, output_type=Output.DICT)

    valid_words = []

    n_boxes = len(data['text'])

    for i in range(n_boxes):
        conf = int(data['conf'][i])

        if conf > threshold:
            word = data['text'][i]

            if word.strip():
                valid_words.append(word)
    
    return " ".join(valid_words)

if __name__ == "__main__":

    # for windows users
    platform_name = platform.system()

    if platform_name == "Windows":
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    print("Select your PPTX file...")
    root = tk.Tk()
    root.withdraw()
    selected_files = filedialog.askopenfilenames(
        title="Select PowerPoint Slides",
        filetypes=[("PowerPoint Files", "*.pptx")]
    )
    root.destroy()

    if not selected_files:
        print("No files selected.")
        exit()

    print(f"Selected {len(selected_files)} files")

    while True:
        format_choice = input("Output format? [1] Markdown (.md)  [2] Plain Text (.txt): ").strip()

        if format_choice == '1':
            ext = ".md"
            format_type = "md"
            break
        else:
            ext = ".txt"
            format_type = "txt"
            break
    
    while True:
        merge_choice = input("Output mode? [1] Separate files  [2] Combined file: ").strip()
        if merge_choice in ['1', '2']:
            break
    
    if merge_choice == '1':
        for pptx_path in selected_files:
            base_name = os.path.splitext(pptx_path)[0]
            output_name = f"{base_name}{ext}"
            
            print(f"Processing: {os.path.basename(pptx_path)} -> {output_name}")

            with open(output_name, "w", encoding="utf-8") as f:
                rip_slides(pptx_path, f, format_type)
    
    else:
        master_name = "Combined notes" + ext

        print(f"Merging all into -> {master_name}")

        with open(master_name, "w", encoding="utf-8") as f:
            for pptx_path in selected_files:
                print(f"Appending: {os.path.basename(pptx_path)}...")
                rip_slides(pptx_path, f, format_type)
        
    print("\nAll files processed.")
