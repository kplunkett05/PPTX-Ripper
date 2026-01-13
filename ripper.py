import platform
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import pytesseract
from pytesseract import Output
import io
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import threading

def rip_slides(path, file_handle, format_type, progress_callback=None, save_images=False) -> str:
    if not os.path.exists(path):
        return
    
    print(f"Opening {path}...")
    pres = Presentation(path)

    filename = os.path.basename(path)
    file_name_no_ext = os.path.splitext(filename)[0]
    file_handle.write(get_header(f"Source: {filename}", 1, format_type))

    media_folder_name = f"{file_name_no_ext}_Media"
    media_folder = os.path.join(os.path.dirname(path), media_folder_name)

    if save_images:
        if not os.path.exists(media_folder):
            print("making dir.")
            print(f"{media_folder}")
            os.makedirs(media_folder)

    total_slides = len(pres.slides)
    for i, slide in enumerate(pres.slides):
        slide_number = i + 1
        print(f"Processing Slide {slide_number}...")


        file_handle.write(get_header(f"Slide {slide_number}", 2, format_type))

        img_count = 0

        for shape in slide.shapes:
            if shape.has_text_frame:
                clean_text = shape.text_frame.text
                file_handle.write(f"{clean_text}\n")
            
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                img_count += 1

                try:
                    image_blob = shape.image.blob
                    ext = shape.image.ext

                    if save_images:
                        print("saving! i hope")
                        
                        img_filename = f"Slide_{slide_number}_Image_{img_count}.{ext}"
                        img_path = os.path.join(media_folder, img_filename)

                        with open(img_path, "wb") as img_file:
                            img_file.write(image_blob)
                            print(f"saved to {img_path}")

                    image = Image.open(io.BytesIO(image_blob))

                    extracted_text = extract_text_with_confidence(image)
                    
                    if extracted_text.strip():
                        file_handle.write(f"\n\n[IMAGE READ]:\n{extracted_text}\n")
                
                except Exception as e:
                    print(f"Could not read image on slide {slide_number}: {e}")
        
        if progress_callback:
            progress_callback(slide_number, total_slides)

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



    # for windows users
    #platform_name = platform.system()

    #if platform_name == "Windows":
    #    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class SlideRipperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Slide Ripper")
        self.root.geometry("600x500")

        # local storage. default to txt and separate modes
        self.selected_files = []
        self.format_var = tk.StringVar(value="txt")
        self.merge_var = tk.StringVar(value="separate")
        self.save_images_var = tk.BooleanVar(value=False)

        # UI section 1
        frame_files = tk.LabelFrame(root, text="1. Source Files", padx=10, pady=10)
        frame_files.pack(fill="x", padx=10, pady=5)
        
        btn_browse = tk.Button(frame_files, text="Add PPTX Files", command=self.browse_files)
        btn_browse.pack(pady=5)

        btn_clear = tk.Button(frame_files, text="Clear added files", command=self.clear_files)
        btn_clear.pack()
        
        self.lbl_file_count = tk.Label(frame_files, text="No files selected")
        self.lbl_file_count.pack()

        # UI section 2
        frame_opts = tk.LabelFrame(root, text="2. Configuration", padx=10, pady=10)
        frame_opts.pack(fill="x", padx=10, pady=5)

        # format
        tk.Label(frame_opts, text="Output Format:").grid(row=0, column=0, sticky="w")
        tk.Radiobutton(frame_opts, text="Text (.txt)", variable=self.format_var, value="txt").grid(row=1, column=0, sticky="w")
        tk.Radiobutton(frame_opts, text="Markdown (.md)", variable=self.format_var, value="md").grid(row=2, column=0, sticky="w")

        # merge mode
        tk.Label(frame_opts, text="Output Files:").grid(row=0, column=1, sticky="w", padx=20)
        tk.Radiobutton(frame_opts, text="Separate Files", variable=self.merge_var, value="separate").grid(row=1, column=1, sticky="w", padx=20)
        tk.Radiobutton(frame_opts, text="Combine All", variable=self.merge_var, value="combine").grid(row=2, column=1, sticky="w", padx=20)

        # image extraction
        tk.Label(frame_opts, text="Extras:").grid(row=0, column=2, sticky="w", padx=20)
        tk.Checkbutton(frame_opts, text="Extract Images to Folder", variable=self.save_images_var, onvalue=True, offvalue=False).grid(row=1, column=2, sticky="w", padx=20)
        
        # UI section 3
        self.btn_run = tk.Button(root, text="Start", bg="green", font=("Arial", 12, "bold"), command=self.start_thread)
        self.btn_run.pack(pady=5)
        
        # Progress bar
        self.lbl_progress = tk.Label(root, text="Ready", font=("Arial, 10"))
        self.lbl_progress.pack(pady=(0, 0))

        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=5)

        self.log_area = scrolledtext.ScrolledText(root, height=10, state='disabled')
        self.log_area.pack(padx=10, pady=5, fill="both", expand=True)

    def browse_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PowerPoint", "*.pptx")])
        self.root.lift()
        self.root.focus_force()
        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.log(f"Queued {os.path.basename(file)}")
            
            total_files = len(self.selected_files)
            
            if total_files == 1:
                display_str = "1 file ready to process"
            else:
                display_str = f"{total_files} files ready to process"
            self.lbl_file_count.config(text=display_str)
            self.log(f"{total_files} files selected.")
    
    def clear_files(self):
        self.selected_files = []
        self.lbl_file_count.config(text="0 files ready to process")
        self.log(f"Cleared selected files.")

    def log(self, message):
        """Helper to write to the scrolling text box"""
        self.log_area.config(state='normal') # Unlocking the box
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END) # Scroll to bottom
        self.log_area.config(state='disabled') # Locking it again

    def start_thread(self):
        if not self.selected_files:
            messagebox.showwarning("No Files Selected", "Please select files first!")
            return
            
        self.btn_run.config(state="disabled", text="Processing...")
        # run in separate thread so app doesnt freeze
        threading.Thread(target=self.run_process).start()

    def run_process(self):
        fmt = self.format_var.get()
        mode = self.merge_var.get()
        ext = "." + fmt

        self.progress_bar['value'] = 0

        should_save_images = self.save_images_var.get()

        def update_ui_progress(current_slide, total_slides):
            percent = (current_slide / total_slides) * 100
            self.progress_bar['value'] = percent
            self.lbl_progress.config(text=f"Processing Slide {current_slide}/{total_slides}")
            self.root.update_idletasks()

        try:
            if mode == "separate":
                for pptx_path in self.selected_files:
                    base_name = os.path.splitext(pptx_path)[0]
                    output_name = base_name + ext
                    
                    self.log(f"Processing: {os.path.basename(pptx_path)}...")
                    with open(output_name, "w", encoding="utf-8") as f:
                        rip_slides(pptx_path, f, fmt, progress_callback=update_ui_progress, save_images=should_save_images)
            else:
                # Combined Mode
                master_name = "Combined_Notes" + ext
                self.log(f"Merging all into {master_name}...")
                with open(master_name, "w", encoding="utf-8") as f:
                    for pptx_path in self.selected_files:
                        self.log(f"Appending {os.path.basename(pptx_path)}...")
                        rip_slides(pptx_path, f, fmt, progress_callback=update_ui_progress, save_images=should_save_images)

            self.lbl_progress.config(text="Done!")
            self.progress_bar['value'] = 100
            self.log("--- COMPLETED ---")
            messagebox.showinfo("Done", "Processing Complete!")
            print("Completed.")

        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))
        
        finally:
            self.btn_run.config(state="normal", text="Start")


if __name__ == "__main__":
    root = tk.Tk()
    app = SlideRipperApp(root)
    root.mainloop()