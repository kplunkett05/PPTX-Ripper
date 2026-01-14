import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import threading
import os

from ripper import rip_slides


class CreateToolTip(object):
    def __init__(self, widget, text="widget info", delay=300):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.id = None
        self.tw = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
    
    def enter(self, event=None):
        self.schedule()
    
    def leave(self, event=None):
        self.unschedule()
        self.hidetip()
        
    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.delay, self.showtip)

    def unschedule(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def showtip(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20

        self.tw = tk.Toplevel(self.widget)

        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x,y))

        label = tk.Label(self.tw, text=self.text, justify="left", background="light grey", foreground="black", relief="solid", borderwidth=1, font=("Arial", "12", "normal"), wraplength=250)
        label.pack(ipadx=3, ipady=1)
    
    def hidetip(self):
        if self.tw:
            self.tw.destroy()
            self.tw = None


class SlideRipperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Slide Ripper")
        self.root.geometry("600x500")

        # local storage. default to txt and separate modes
        self.selected_files = []
        self.format_var = tk.StringVar(value="txt")
        self.merge_var = tk.StringVar(value="combine")
        self.save_images_var = tk.BooleanVar(value=False)
        self.ocr_var = tk.BooleanVar(value=False)



        # ---UI section 1---
        frame_files = tk.LabelFrame(root, text="1. Source Files", padx=10, pady=10)
        frame_files.pack(fill="x", padx=10, pady=5)
        
        btn_browse = tk.Button(frame_files, text="Add PPTX Files", command=self.browse_files)
        btn_browse.pack(pady=5)

        btn_clear = tk.Button(frame_files, text="Clear added files", command=self.clear_files)
        btn_clear.pack()
        
        self.lbl_file_count = tk.Label(frame_files, text="No files selected")
        self.lbl_file_count.pack()



        # ---UI section 2---
        frame_opts = tk.LabelFrame(root, text="2. Configuration", padx=10, pady=10)
        frame_opts.pack(fill="x", padx=10, pady=5)

        # format
        tk.Label(frame_opts, text="Output Format:").grid(row=0, column=0, sticky="w")

        text_format_button = tk.Radiobutton(frame_opts, text="Text (.txt)", variable=self.format_var, value="txt")
        text_format_button.grid(row=1, column=0, sticky="w")
        CreateToolTip(text_format_button, "Output notes into .txt format.")

        markdown_format_button = tk.Radiobutton(frame_opts, text="Markdown (.md)", variable=self.format_var, value="md")
        markdown_format_button.grid(row=2, column=0, sticky="w")
        CreateToolTip(markdown_format_button, "Output notes into .md format. May require a program to read.")

        # merge mode
        tk.Label(frame_opts, text="Output Files:").grid(row=0, column=1, sticky="w", padx=20)
        
        combine_files_button = tk.Radiobutton(frame_opts, text="Single Note", variable=self.merge_var, value="combine")
        combine_files_button.grid(row=1, column=1, sticky="w", padx=20)
        CreateToolTip(combine_files_button, "Combine notes for every presentation into a single output file.")
        
        separate_files_button = tk.Radiobutton(frame_opts, text="Separate Notes", variable=self.merge_var, value="separate")
        separate_files_button.grid(row=2, column=1, sticky="w", padx=20)
        CreateToolTip(separate_files_button, "Create separate notes for every presentation, if multiple are selected.")

        # image extraction
        tk.Label(frame_opts, text="Extras:").grid(row=0, column=2, sticky="w", padx=20)
        extract_images_button = tk.Checkbutton(frame_opts, text="Extract Images to Folder", variable=self.save_images_var, onvalue=True, offvalue=False)
        extract_images_button.grid(row=1, column=2, sticky="w", padx=20)
        CreateToolTip(extract_images_button, "Extract all images found in the presentation to a separate folder within the same directory.")
        
        ocr_button = tk.Checkbutton(frame_opts, text="OCR (Requires Tesseract)", variable=self.ocr_var, onvalue=True, offvalue=False)
        ocr_button.grid(row=2, column=2, sticky="w", padx=20)
        CreateToolTip(ocr_button, "Use optical character recognition to extract text from image files that contain text. Requires a valid installation of the Tesseract OCR engine.")
        


        # ---UI section 3---
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
        use_ocr = self.ocr_var.get()

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
                        rip_slides(pptx_path, f, fmt, progress_callback=update_ui_progress, save_images=should_save_images, use_ocr=use_ocr)
            else:
                # Combined Mode

                first_file_path = self.selected_files[0]
                directory = os.path.dirname(first_file_path)

                master_name = "Combined_Notes" + ext
                master_path = os.path.join(directory, master_name)
                
                self.log(f"Merging all into {master_name}...")
                
                with open(master_path, "w", encoding="utf-8") as f:
                    for pptx_path in self.selected_files:
                        self.log(f"Appending {os.path.basename(pptx_path)}...")
                        rip_slides(pptx_path, f, fmt, progress_callback=update_ui_progress, save_images=should_save_images, use_ocr=use_ocr)

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