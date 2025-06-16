import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import pandas as pd
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from io import BytesIO
from datetime import datetime
import customtkinter as ctk
import webbrowser
import qrcode
from qrcode.image.styledpil import StyledPilImage
from qrcode.image.styles.moduledrawers import RoundedModuleDrawer

class CertificateGenerator(ctk.CTk):
    """Modern certificate generator application"""
    
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Accredify Suite - Certificate Generator")
        self.iconbitmap(r"C:\Users\User\Desktop\assets\icon.ico")
        self.geometry("1200x800")
        self.minsize(1000, 700)
        
        # Set appearance
        ctk.set_appearance_mode("System Theme")  # Light/Dark mode
        ctk.set_default_color_theme("blue")  # Built-in theme
        
        # Initialize variables
        self.templates = {
            "Classic Elegance": self.generate_classic_certificate,
            "Modern Professional": self.generate_modern_certificate,
            # "Clean Minimalist": self.generate_minimalist_certificate,
            "Academic Diploma": self.generate_academic_diploma,
            "Corporate Achievement": self.generate_corporate_certificate,
            "Workshop Completion": self.generate_workshop_certificate
        }
        self.logo_path = ""
        self.signature_path = ""
        self.batch_mode = False
        self.batch_file_path = ""
        self.preview_zoom = 1.0

        self.interaction_mode = None  # 'move', 'resize', 'rotate'
        self.current_element = None
        self.cursor_over_element = False
        self.cursors = {
            'move': 'fleur',
            'resize': 'sizing',
            'rotate': 'exchange'
        }
        
        # Setup UI
        self.setup_ui()
        # self.setup_customization_ui()
        
    def setup_ui(self):
        """Initialize all UI components"""
        # Configure grid layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Create sidebar frame
        self.sidebar_frame = ctk.CTkFrame(self, width=350, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)
        
        # Logo label
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Certificate Generator",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Template selection
        self.template_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Template Style:",
            anchor="w"
        )
        self.template_label.grid(row=1, column=0, padx=20, pady=(10, 0))
        
        self.template_var = ctk.StringVar(value="Modern Professional")
        self.template_dropdown = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=list(self.templates.keys()),
            variable=self.template_var,
            command=lambda e: self.generate_preview()
        )
        self.template_dropdown.grid(row=2, column=0, padx=20, pady=(0, 10))
        
        # Create tabview for settings
        self.tabview = ctk.CTkTabview(self.sidebar_frame)
        self.tabview.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="nsew")
        
        # Add tabs
        self.tabview.add("Content")
        self.tabview.add("Design")
        self.tabview.add("Batch")
        
        # Content tab
        self.name_label = ctk.CTkLabel(
            self.tabview.tab("Content"), 
            text="Recipient Name:",
            anchor="w"
        )
        self.name_label.grid(row=0, column=0, padx=10, pady=(10, 0))
        
        self.name_var = ctk.StringVar()
        self.name_entry = ctk.CTkEntry(
            self.tabview.tab("Content"),
            textvariable=self.name_var,
            placeholder_text="Full Name"
        )
        self.name_entry.grid(row=1, column=0, padx=10, pady=(0, 10))
        self.name_entry.bind("<KeyRelease>", lambda e: self.generate_preview())
        
        self.course_label = ctk.CTkLabel(
            self.tabview.tab("Content"), 
            text="Course Title:",
            anchor="w"
        )
        self.course_label.grid(row=2, column=0, padx=10, pady=(0, 0))
        
        self.course_var = ctk.StringVar()
        self.course_entry = ctk.CTkEntry(
            self.tabview.tab("Content"),
            textvariable=self.course_var,
            placeholder_text="Course/Program Name"
        )
        self.course_entry.grid(row=3, column=0, padx=10, pady=(0, 10))
        self.course_entry.bind("<KeyRelease>", lambda e: self.generate_preview())
        
        self.date_label = ctk.CTkLabel(
            self.tabview.tab("Content"), 
            text="Completion Date:",
            anchor="w"
        )
        self.date_label.grid(row=4, column=0, padx=10, pady=(0, 0))
        
        self.date_var = ctk.StringVar(value=datetime.now().strftime("%B %d, %Y"))
        self.date_entry = DateEntry(
            self.tabview.tab("Content"),
            textvariable=self.date_var,
            date_pattern="yyyy-mm-dd",
            background="darkblue",
            foreground="white",
            borderwidth=2
        )
        self.date_entry.grid(row=5, column=0, padx=10, pady=(0, 10))
        self.date_entry.bind("<<DateEntrySelected>>", lambda e: self.generate_preview())
        
        self.desc_label = ctk.CTkLabel(
            self.tabview.tab("Content"), 
            text="Description:",
            anchor="w"
        )
        self.desc_label.grid(row=6, column=0, padx=10, pady=(0, 0))
        
        self.desc_var = ctk.StringVar()
        self.desc_entry = ctk.CTkEntry(
            self.tabview.tab("Content"),
            textvariable=self.desc_var,
            placeholder_text="Optional description or credits"
        )
        self.desc_entry.grid(row=7, column=0, padx=10, pady=(0, 10))
        self.desc_entry.bind("<KeyRelease>", lambda e: self.generate_preview())
        

        # Auto-fill dummy data
        self.name_var.set("John Doe")
        self.course_var.set("Python Programming 101")
        self.date_var.set(datetime.now().strftime("%Y-%m-%d"))  # Match the date_pattern you used
        self.desc_var.set("with distinction and excellence")

        # Design tab
        self.logo_label = ctk.CTkLabel(
            self.tabview.tab("Design"), 
            text="Organization Logo:",
            anchor="w"
        )
        self.logo_label.grid(row=0, column=0, padx=10, pady=(10, 0))
        
        self.logo_button = ctk.CTkButton(
            self.tabview.tab("Design"),
            text="Upload Logo",
            command=self.upload_logo
        )
        self.logo_button.grid(row=1, column=0, padx=10, pady=(0, 10))
        
        self.logo_preview = ctk.CTkLabel(
            self.tabview.tab("Design"),
            text="No logo selected",
            fg_color=("gray70", "gray30"),
            corner_radius=6
        )
        self.logo_preview.grid(row=2, column=0, padx=10, pady=(0, 20))
        
        self.signature_label = ctk.CTkLabel(
            self.tabview.tab("Design"), 
            text="Signature Image:",
            anchor="w"
        )
        self.signature_label.grid(row=3, column=0, padx=10, pady=(0, 0))
        
        self.signature_button = ctk.CTkButton(
            self.tabview.tab("Design"),
            text="Upload Signature",
            command=self.upload_signature
        )
        self.signature_button.grid(row=4, column=0, padx=10, pady=(0, 10))
        
        self.signature_preview = ctk.CTkLabel(
            self.tabview.tab("Design"),
            text="No signature selected",
            fg_color=("gray70", "gray30"),
            corner_radius=6
        )
        self.signature_preview.grid(row=5, column=0, padx=10, pady=(0, 20))
        
        # Batch tab
        self.batch_var = ctk.BooleanVar(value=False)
        self.batch_check = ctk.CTkSwitch(
            self.tabview.tab("Batch"),
            text="Enable Batch Mode",
            variable=self.batch_var,
            command=self.toggle_batch_mode
        )
        self.batch_check.grid(row=0, column=0, padx=10, pady=(10, 10))
        
        self.batch_file_button = ctk.CTkButton(
            self.tabview.tab("Batch"),
            text="Select Data File",
            command=self.select_batch_file,
            state="disabled"
        )
        self.batch_file_button.grid(row=1, column=0, padx=10, pady=(0, 10))
        
        self.batch_status = ctk.CTkLabel(
            self.tabview.tab("Batch"),
            text="No file selected",
            text_color="gray50"
        )
        self.batch_status.grid(row=2, column=0, padx=10, pady=(0, 20))
        
        # Action buttons
        self.generate_preview_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Generate Preview",
            command=self.generate_preview
        )
        self.generate_preview_btn.grid(row=4, column=0, padx=20, pady=10)
        
        self.generate_pdf_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Generate PDF",
            command=self.generate_pdf,
            fg_color="green",
            hover_color="dark green"
        )
        self.generate_pdf_btn.grid(row=5, column=0, padx=20, pady=10)
        
        # Appearance mode
        self.appearance_mode_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Appearance Mode:",
            anchor="w"
        )
        self.appearance_mode_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=[ "System Theme", "Dark", "Light"],
            command=self.change_appearance_mode
        )
        self.appearance_mode_optionemenu.grid(row=8, column=0, padx=20, pady=(0, 20))
        
        # Create main preview area
        self.preview_frame = ctk.CTkFrame(self, corner_radius=0)
        self.preview_frame.grid(row=0, column=1, sticky="nsew")
        self.preview_frame.grid_rowconfigure(1, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        
        # Preview label and controls
        self.preview_header = ctk.CTkFrame(self.preview_frame, height=30)
        self.preview_header.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        self.preview_title = ctk.CTkLabel(
            self.preview_header,
            text="Certificate Preview",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.preview_title.pack(side="left", padx=10)
        
        # Zoom controls
        self.zoom_frame = ctk.CTkFrame(self.preview_header, fg_color="transparent")
        self.zoom_frame.pack(side="right", padx=10)
        
        self.zoom_out_btn = ctk.CTkButton(
            self.zoom_frame,
            text="-",
            width=30,
            command=lambda: self.adjust_zoom(-0.1)
        )
        self.zoom_out_btn.pack(side="left", padx=(0, 5))
        
        self.zoom_label = ctk.CTkLabel(
            self.zoom_frame,
            text="100%",
            width=50
        )
        self.zoom_label.pack(side="left")
        
        self.zoom_in_btn = ctk.CTkButton(
            self.zoom_frame,
            text="+",
            width=30,
            command=lambda: self.adjust_zoom(0.1)
        )
        self.zoom_in_btn.pack(side="left", padx=(5, 0))
        
        # Preview canvas
        self.preview_container = ctk.CTkFrame(self.preview_frame, fg_color="transparent")
        self.preview_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.preview_container.grid_rowconfigure(0, weight=1)
        self.preview_container.grid_columnconfigure(0, weight=1)
        
        self.preview_canvas = tk.Canvas(
            self.preview_container,
            bg='white' if ctk.get_appearance_mode() == "Light" else "#2b2b2b",
            highlightthickness=0
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbars
        self.scroll_y = ctk.CTkScrollbar(
            self.preview_container,
            orientation="vertical",
            command=self.preview_canvas.yview
        )
        self.scroll_y.grid(row=0, column=1, sticky="ns")
        
        self.scroll_x = ctk.CTkScrollbar(
            self.preview_container,
            orientation="horizontal",
            command=self.preview_canvas.xview
        )
        self.scroll_x.grid(row=1, column=0, sticky="ew")
        
        self.preview_canvas.configure(
            yscrollcommand=self.scroll_y.set,
            xscrollcommand=self.scroll_x.set
        )
            
        # Status bar
        self.status_bar = ctk.CTkLabel(
            self,
            text="Ready",
            fg_color=("gray70", "gray30"),
            corner_radius=0,
            anchor="w"
        )
        self.status_bar.grid(row=1, column=0, columnspan=2, sticky="ew")
        
        # # Add hover event binding
        # self.preview_canvas.bind("<Motion>", self.check_hover)
        # self.preview_canvas.bind("<Button-1>", self.start_drag)
        # self.preview_canvas.bind("<B1-Motion>", self.on_drag)
        # self.preview_canvas.bind("<ButtonRelease-1>", self.end_drag)

        
        # Generate initial preview
        self.generate_preview()
    
    def change_appearance_mode(self, new_appearance_mode):
        """Change appearance mode (light/dark)"""
        ctk.set_appearance_mode(new_appearance_mode)
        self.preview_canvas.configure(
            bg='white' if new_appearance_mode == "Light" else "#2b2b2b"
        )
        self.generate_preview()
    
    def adjust_zoom(self, change):
        """Adjust preview zoom level"""
        self.preview_zoom = max(0.5, min(2.0, self.preview_zoom + change))
        self.zoom_label.configure(text=f"{self.preview_zoom * 100}%")
        self.generate_preview()

        # Center the preview after zoom
        self.preview_canvas.xview_moveto(0.5)
        self.preview_canvas.yview_moveto(0.5)
    
    def upload_logo(self):
        """Handle logo file upload"""
        file_path = filedialog.askopenfilename(
            title="Select Organization Logo",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
        )
        if file_path:
            try:
                img = Image.open(file_path)
                img.verify()
                self.logo_path = file_path
                
                # Display preview
                preview_img = Image.open(file_path)
                preview_img.thumbnail((200, 100))
                photo = ImageTk.PhotoImage(preview_img)
                
                self.logo_preview.configure(image=photo, text="")
                self.logo_preview.image = photo
                
                self.status_bar.configure(text="Logo uploaded successfully")
                self.generate_preview()
            except Exception as e:
                messagebox.showerror("Error", f"Invalid image file: {str(e)}")
                self.status_bar.configure(text="Logo upload failed")
    
    def upload_signature(self):
        """Handle signature file upload"""
        file_path = filedialog.askopenfilename(
            title="Select Signature Image",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
        )
        if file_path:
            try:
                img = Image.open(file_path)
                img.verify()
                self.signature_path = file_path
                
                # Display preview
                preview_img = Image.open(file_path)
                preview_img.thumbnail((200, 60))
                photo = ImageTk.PhotoImage(preview_img)
                
                self.signature_preview.configure(image=photo, text="")
                self.signature_preview.image = photo
                
                self.status_bar.configure(text="Signature uploaded successfully")
                self.generate_preview()
            except Exception as e:
                messagebox.showerror("Error", f"Invalid image file: {str(e)}")
                self.status_bar.configure(text="Signature upload failed")
    
    def toggle_batch_mode(self):
        """Toggle between single and batch mode"""
        self.batch_mode = self.batch_var.get()
        state = "normal" if self.batch_mode else "disabled"
        self.batch_file_button.configure(state=state)
        self.status_bar.configure(text="Batch mode " + ("enabled" if self.batch_mode else "disabled"))
    
    def select_batch_file(self):
        """Select CSV file for batch processing"""
        file_path = filedialog.askopenfilename(
            title="Select Participant Data File",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            self.batch_file_path = file_path
            filename = os.path.basename(file_path)
            self.batch_status.configure(text=f"Loaded: {filename}")
            self.status_bar.configure(text=f"Batch file loaded: {filename}")
    
    def validate_fields(self):
        """Validate all required fields"""
        if not self.name_var.get().strip():
            messagebox.showerror("Error", "Recipient name is required!")
            self.status_bar.configure(text="Error: Missing recipient name")
            return False
        if not self.course_var.get().strip():
            messagebox.showerror("Error", "Course title is required!")
            self.status_bar.configure(text="Error: Missing course title")
            return False
        if not self.date_var.get().strip():
            messagebox.showerror("Error", "Completion date is required!")
            self.status_bar.configure(text="Error: Missing completion date")
            return False
        return True
    
    def generate_qr_code(self, data, size=100):
        """Generate a styled QR code image"""
        try:
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(data)
            qr.make(fit=True)
            
            # Create a styled QR code
            img = qr.make_image(
                image_factory=StyledPilImage,
                module_drawer=RoundedModuleDrawer(),
                eye_drawer=qrcode.image.styles.moduledrawers.RoundedModuleDrawer(),
                embeded_image_path=self.logo_path if self.logo_path else None
            )
            
            # Resize if needed
            if size:
                img = img.resize((size, size), Image.Resampling.LANCZOS)
                
            return img
        except Exception as e:
            print(f"QR code generation error: {str(e)}")
            return None
    
    def generate_preview(self):
        """Generate a preview of the certificate"""
        print("\n=== Starting preview generation ===")
        if not self.validate_fields():
            print("Validation failed - returning")
            return
            
        try:
            print("Creating temporary PDF in memory...")
            buffer = BytesIO()
            template_func = self.templates[self.template_var.get()]
            print(f"Using template: {self.template_var.get()}")
            
            print("Generating PDF content...")
            template_func(buffer, preview=True)
            
            # Save the buffer to a temporary file
            temp_pdf = "temp_preview.pdf"
            print(f"Saving temporary PDF to {temp_pdf}...")
            with open(temp_pdf, "wb") as f:
                f.write(buffer.getvalue())
            
            # Convert first page of PDF to image
            print("Converting PDF to image...")
            from pdf2image import convert_from_path
            try:
                images = convert_from_path(temp_pdf, dpi=100)
                if not images:
                    raise ValueError("Failed to convert PDF to image - no images returned")
                print("PDF conversion successful")
            except Exception as e:
                print(f"PDF conversion failed: {str(e)}")
                raise
                
            img = images[0]  # Get first page
            print(f"Original image dimensions: {img.width}x{img.height}")
            print(f"Current zoom level: {self.preview_zoom}")
            
            # Safely calculate new dimensions as integers
            try:
                new_width = int(round(img.width * float(self.preview_zoom)))
                new_height = int(round(img.height * float(self.preview_zoom)))
                print(f"Calculated dimensions: {new_width}x{new_height}")
            except (TypeError, ValueError) as e:
                print(f"Dimension calculation error: {str(e)} - using original dimensions")
                new_width = img.width
                new_height = img.height
                self.preview_zoom = 1.0
                self.zoom_label.configure(text="100%")
            
            # Ensure reasonable dimensions
            new_width = max(10, min(new_width, 5000))  # Between 10 and 5000 pixels
            new_height = max(10, min(new_height, 5000))
            print(f"Final dimensions after constraints: {new_width}x{new_height}")
            
            # Resize the image with error handling
            try:
                print("Attempting image resize...")
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                print("Image resize successful")
            except Exception as e:
                print(f"Image resize failed: {str(e)}")
                raise ValueError(f"Image resize failed: {str(e)}")
            
            try:
                print("Creating PhotoImage...")
                photo = ImageTk.PhotoImage(img)
                print("PhotoImage created successfully")
            except Exception as e:
                print(f"PhotoImage creation failed: {str(e)}")
                raise
                
            # Update preview canvas
            print("Updating canvas...")
            self.preview_canvas.delete("all")
            img_width = photo.width()
            img_height = photo.height()
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            # Calculate centered position
            x = max(0, (canvas_width - img_width) // 2)
            y = max(0, (canvas_height - img_height) // 2)
            
            self.preview_canvas.create_image(x, y, anchor=tk.NW, image=photo)
            self.preview_canvas.image = photo  # Keep reference
            
            # Update scroll region
            print("Updating scroll region...")
            self.preview_canvas.configure(
                scrollregion=self.preview_canvas.bbox("all")
            )
            
            # Clean up
            try:
                print("Cleaning up temporary files...")
                os.remove(temp_pdf)
                print("Temporary files removed")
            except Exception as e:
                print(f"Error removing temp file: {str(e)}")
                
            print("Preview generation complete")
            self.status_bar.configure(text="Preview generated successfully")
            
        except Exception as e:
            print(f"!!! ERROR in preview generation: {str(e)}")
            messagebox.showerror("Error", f"Failed to generate preview: {str(e)}")
            self.status_bar.configure(text="Preview generation failed")
        print("=== Preview generation process ended ===\n")
    
    def generate_pdf(self):
        """Generate the final PDF certificate(s)"""
        if self.batch_mode:
            self.process_batch()
        else:
            self.process_single()
    
    def process_single(self):
        """Process a single certificate"""
        if not self.validate_fields():
            return
            
        default_filename = f"Certificate_{self.name_var.get().replace(' ', '_')}.pdf"
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=default_filename,
            title="Save Certificate As")
            
        if output_path:
            try:
                template_func = self.templates[self.template_var.get()]
                with open(output_path, 'wb') as f:
                    template_func(f)
                messagebox.showinfo("Success", f"Certificate saved to:\n{output_path}")
                self.status_bar.configure(text=f"Certificate saved: {os.path.basename(output_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate PDF: {str(e)}")
                self.status_bar.configure(text="PDF generation failed")
    
    def process_batch(self):
        """Process certificates in batch mode from CSV"""
        if not self.batch_file_path:
            messagebox.showerror("Error", "Please select a batch file first!")
            self.status_bar.configure(text="Error: No batch file selected")
            return
            
        try:
            # Read the batch file
            if self.batch_file_path.endswith('.csv'):
                df = pd.read_csv(self.batch_file_path)
            else:  # Excel
                df = pd.read_excel(self.batch_file_path)
                
            # Validate required columns
            required_cols = ['Name', 'Course', 'Date']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_cols)}")
                self.status_bar.configure(text=f"Error: Missing columns {', '.join(missing_cols)}")
                return
                
            # Select output directory
            output_dir = filedialog.askdirectory(
                title="Select Output Directory for Batch Certificates")
            if not output_dir:
                self.status_bar.configure(text="Batch processing cancelled")
                return
                
            # Process each row
            template_func = self.templates[self.template_var.get()]
            success_count = 0
            total_rows = len(df)
            
            # Create progress window
            progress_window = ctk.CTkToplevel(self)
            progress_window.title("Batch Processing")
            progress_window.geometry("400x150")
            progress_window.resizable(False, False)
            
            ctk.CTkLabel(
                progress_window, 
                text="Generating Certificates...",
                font=ctk.CTkFont(weight="bold")
            ).pack(pady=10)
            
            progress_var = ctk.DoubleVar()
            progress_bar = ctk.CTkProgressBar(
                progress_window,
                variable=progress_var,
                orientation="horizontal"
            )
            progress_bar.pack(fill="x", padx=20, pady=5)
            progress_bar.set(0)
            
            status_label = ctk.CTkLabel(progress_window, text="Starting...")
            status_label.pack(pady=5)
            
            # Force update the progress window
            progress_window.update()
            
            for idx, row in df.iterrows():
                try:
                    # Update fields from CSV
                    self.name_var.set(row['Name'])
                    self.course_var.set(row.get('Course', ''))
                    self.date_var.set(row.get('Date', ''))
                    self.desc_var.set(row.get('Description', ''))
                    
                    # Generate PDF
                    filename = f"Certificate_{row['Name'].replace(' ', '_')}.pdf"
                    output_path = os.path.join(output_dir, filename)
                    with open(output_path, 'wb') as f:
                        template_func(f)
                    success_count += 1
                    
                    # Update progress
                    progress_var.set((idx + 1) / total_rows)
                    status_label.configure(text=f"Processing {idx + 1} of {total_rows}: {row['Name']}")
                    progress_window.update()
                    
                except Exception as e:
                    print(f"Error processing {row['Name']}: {str(e)}")
                    
            progress_window.destroy()
            
            messagebox.showinfo("Batch Complete", 
                             f"Successfully generated {success_count} of {total_rows} certificates.")
            self.status_bar.configure(text=f"Batch complete: {success_count}/{total_rows} certificates generated")
                              
        except Exception as e:
            messagebox.showerror("Error", f"Batch processing failed: {str(e)}")
            self.status_bar.configure(text="Batch processing failed")
    
    # Certificate template generation functions remain the same as original
    def generate_classic_certificate(self, output, preview=False):
        """Generate a classic-style certificate"""
        # Get field values
        name = self.name_var.get()
        course = self.course_var.get()
        try:
            raw_date = self.date_var.get()
            date_obj = datetime.strptime(raw_date, "%Y-%m-%d")
            date = date_obj.strftime("%B %d, %Y")
        except ValueError:
            date = raw_date
        description = self.desc_var.get()
        
        # Create PDF canvas
        c = canvas.Canvas(output, pagesize=landscape(A4))
        width, height = landscape(A4)
        
        # Background design with subtle texture
        c.setFillColor(HexColor("#F9F5E8"))
        c.rect(0, 0, width, height, fill=True, stroke=False)
        
        # Add decorative border elements
        border_color = HexColor("#8B7355")
        c.setStrokeColor(border_color)
        c.setLineWidth(10)
        c.roundRect(30, 30, width-60, height-60, radius=10, fill=False, stroke=True)
        
        # Add subtle watermark
        c.setFillColor(HexColor("#F0E6D2"))
        c.setFont("Helvetica-Bold", 120)
        c.drawCentredString(width//2, height//2-60, "CERTIFICATE")
        c.setFillColor(HexColor("#333333"))
        
        # Header with elegant typography
        c.setFont("Helvetica-Bold", 42)
        c.setFillColor(HexColor("#2C3E50"))
        c.drawCentredString(width//2, height-120, "CERTIFICATE OF ACHIEVEMENT")
        
        # Decorative line with accent
        c.setStrokeColor(HexColor("#E74C3C"))
        c.setLineWidth(3)
        c.line(width//2-180, height-150, width//2+180, height-150)
        
        # Main content
        c.setFont("Helvetica", 20)
        c.drawCentredString(width//2, height-200, "This is to certify that")
        
        # Recipient name with elegant styling
        c.setFont("Times-BoldItalic", 36)
        c.setFillColor(HexColor("#8B7355"))
        c.drawCentredString(width//2, height-260, name.upper())
        
        # Course description
        c.setFont("Helvetica", 18)
        c.setFillColor(HexColor("#333333"))
        
        text = f"has successfully completed the course of study in"
        p = Paragraph(text, ParagraphStyle(name="Normal", fontSize=18, leading=22, alignment=1))
        p.wrapOn(c, width-200, 50)
        p.drawOn(c, 100, height-320)
        
        c.setFont("Helvetica-Bold", 22)
        c.setFillColor(HexColor("#2C3E50"))
        c.drawCentredString(width//2, height-360, f"«{course}»")
        
        if description:
            c.setFont("Helvetica", 16)
            c.setFillColor(HexColor("#555555"))
            p = Paragraph(description, ParagraphStyle(name="Normal", fontSize=16, leading=20, alignment=1))
            p.wrapOn(c, width-200, 100)
            p.drawOn(c, 100, height-400)
        
        # Date
        c.setFont("Helvetica-Oblique", 16)
        c.drawCentredString(width//2, height-450, f"Awarded this {date}")
        
        # Logo and signature area
        y_pos = height-550
        if self.logo_path:
            try:
                logo = ImageReader(self.logo_path)
                c.drawImage(logo, 100, y_pos, width=150, height=100, preserveAspectRatio=True, mask='auto')
                c.setStrokeColor(border_color)
                c.setLineWidth(0.5)
                c.line(100, y_pos-10, 250, y_pos-10)
                c.setFont("Helvetica", 10)
                c.drawString(100, y_pos-25, "Official Seal")
            except Exception as e:
                print(f"Error loading logo: {str(e)}")
                
        if self.signature_path:
            try:
                signature = ImageReader(self.signature_path)
                c.drawImage(signature, width-250, y_pos, width=150, height=80, preserveAspectRatio=True, mask='auto')
                c.setStrokeColor(border_color)
                c.setLineWidth(0.5)
                c.line(width-250, y_pos-10, width-100, y_pos-10)
                c.setFont("Helvetica", 10)
                c.drawCentredString(width-175, y_pos-25, "Authorized Signature")
            except Exception as e:
                print(f"Error loading signature: {str(e)}")

        # Generate verification data
        verification_data = f"""
        Certificate Verification
        Name: {name}
        Course: {course}
        Date: {date}
        ID: {datetime.now().strftime('%Y%m%d')}-{hash(name) % 10000:04d}
        """
        
        # Generate and add QR code
        qr_img = self.generate_qr_code(verification_data, size=80)
        if qr_img:
            # Save QR code to a temporary buffer
            qr_buffer = BytesIO()
            qr_img.save(qr_buffer, format='PNG')
            qr_buffer.seek(0)
            
            # Add QR code to PDF
            qr_reader = ImageReader(qr_buffer)
            c.drawImage(qr_reader, width-120, 50, width=80, height=80, preserveAspectRatio=True, mask='auto')
        
        # Footer
        c.setFont("Helvetica", 10)
        c.setFillColor(HexColor("#666666"))
        c.drawCentredString(width//2, 50, "This certificate is awarded as recognition of professional achievement")
        
        if preview:
            c.showPage()
            c.save()
            return output
        
        c.save()
    
    # def generate_minimalist_certificate(self, output, preview=False):
    #     """Generate a minimalist-style certificate with customizable elements"""
    #     # Get field values
    #     name = self.name_var.get()
    #     course = self.course_var.get()
    #     try:
    #         raw_date = self.date_var.get()
    #         date_obj = datetime.strptime(raw_date, "%Y-%m-%d")
    #         date = date_obj.strftime("%B %d, %Y")
    #     except ValueError:
    #         date = raw_date
    #     description = self.desc_var.get()
        
    #     # Create a dialog for customizing the minimalist certificate
    #     customizer = ctk.CTkToplevel(self)
    #     customizer.title("Customize Minimalist Certificate")
    #     customizer.geometry("1000x800")
        
    #     # Variables for customization
    #     orientation_var = ctk.StringVar(value="portrait")
    #     custom_elements = []  # Stores (image_path, x, y, scale)
        
    #     # Create preview canvas
    #     preview_canvas = tk.Canvas(
    #         customizer,
    #         bg='white',
    #         highlightthickness=0
    #     )
    #     preview_canvas.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        
    #     # Sidebar controls
    #     sidebar = ctk.CTkFrame(customizer)
    #     sidebar.pack(side="left", fill="y", padx=10, pady=10)
        
    #     # Orientation selection
    #     ctk.CTkLabel(sidebar, text="Orientation:").pack(pady=(10,0))
    #     orientation_menu = ctk.CTkOptionMenu(
    #         sidebar,
    #         values=["portrait", "landscape"],
    #         variable=orientation_var,
    #         command=lambda _: self.update_minimalist_preview(preview_canvas, orientation_var.get(), custom_elements)
    #     )
    #     orientation_menu.pack(pady=(0,10))
        
    #     # Add element button
    #     def add_element():
    #         file_path = filedialog.askopenfilename(
    #             title="Select Design Element",
    #             filetypes=[("PNG Files", "*.png")]
    #         )
    #         if file_path:
    #             try:
    #                 img = Image.open(file_path)
    #                 img.verify()
    #                 # Add to elements list with default position (center)
    #                 custom_elements.append((file_path, 0.5, 0.5, 1.0))
    #                 self.update_minimalist_preview(preview_canvas, orientation_var.get(), custom_elements)
    #             except Exception as e:
    #                 messagebox.showerror("Error", f"Invalid image file: {str(e)}")
        
    #     add_btn = ctk.CTkButton(
    #         sidebar,
    #         text="Add Design Element",
    #         command=add_element
    #     )
    #     add_btn.pack(pady=10)
        
    #     # Element list and controls
    #     element_frame = ctk.CTkScrollableFrame(sidebar)
    #     element_frame.pack(fill="both", expand=True)
        
    #     # Generate PDF button
    #     def generate_custom_pdf():
    #         # Create PDF with current settings
    #         pagesize = A4 if orientation_var.get() == "portrait" else landscape(A4)
    #         c = canvas.Canvas(output, pagesize=pagesize)
    #         width, height = pagesize
            
    #         # Clean white background
    #         c.setFillColor(HexColor("#FFFFFF"))
    #         c.rect(0, 0, width, height, fill=True, stroke=False)
            
    #         # Add all custom elements
    #         for elem in custom_elements:
    #             img_path, rel_x, rel_y, scale = elem
    #             try:
    #                 img = ImageReader(img_path)
    #                 # Convert relative to absolute coordinates
    #                 abs_x = width * rel_x
    #                 abs_y = height * rel_y
    #                 # Draw image with scaling
    #                 img_width = 100 * scale  # Base size 100pt, scaled
    #                 c.drawImage(img, abs_x, abs_y, width=img_width, preserveAspectRatio=True, mask='auto')
    #             except Exception as e:
    #                 print(f"Error drawing custom element: {str(e)}")
            
    #         # Main content
    #         c.setFont("Helvetica", 24)
    #         c.setFillColor(HexColor("#333333"))
    #         c.drawCentredString(width//2, height-100, "Certificate of Completion")
            
    #         c.setFont("Helvetica", 20)
    #         c.drawCentredString(width//2, height-160, name)
            
    #         c.setFont("Helvetica", 16)
    #         text = f"For successfully completing:"
    #         p = Paragraph(text, ParagraphStyle(
    #             name="Normal", 
    #             fontSize=16, 
    #             leading=20,
    #             alignment=1
    #         ))
    #         p.wrapOn(c, width-200, 50)
    #         p.drawOn(c, 100, height-220)
            
    #         c.setFont("Helvetica", 18)
    #         c.setFillColor(HexColor("#2C3E50"))
    #         c.drawCentredString(width//2, height-260, course)
            
    #         if description:
    #             c.setFont("Helvetica", 14)
    #             c.setFillColor(HexColor("#555555"))
    #             p = Paragraph(description, ParagraphStyle(
    #                 name="Normal", 
    #                 fontSize=14, 
    #                 leading=18,
    #                 alignment=1
    #             ))
    #             p.wrapOn(c, width-200, 100)
    #             p.drawOn(c, 100, height-300)
            
    #         # Date
    #         c.setFont("Helvetica", 14)
    #         c.setFillColor(HexColor("#777777"))
    #         c.drawCentredString(width//2, height-370, f"Completed on {date}")
            
    #         if preview:
    #             c.showPage()
    #             c.save()
    #             return output
            
    #         c.save()
    #         customizer.destroy()
    #         messagebox.showinfo("Success", "Custom certificate generated!")
        
    #     gen_btn = ctk.CTkButton(
    #         sidebar,
    #         text="Generate PDF",
    #         command=generate_custom_pdf,
    #         fg_color="green"
    #     )
    #     gen_btn.pack(pady=10)
        
    #     # Initial preview update
    #     self.update_minimalist_preview(preview_canvas, orientation_var.get(), custom_elements)

    # def update_minimalist_preview(self, canvas, orientation, elements):
    #     """Update the preview canvas with current settings"""
    #     canvas.delete("all")
        
    #     # Determine dimensions based on orientation
    #     if orientation == "portrait":
    #         width, height = 595, 842  # A4 portrait at 72dpi
    #     else:
    #         width, height = 842, 595  # A4 landscape at 72dpi
        
    #     # Draw background
    #     canvas.create_rectangle(0, 0, width, height, fill="white", outline="")
        
    #     # Draw all elements
    #     for elem in elements:
    #         img_path, rel_x, rel_y, scale = elem
    #         try:
    #             img = Image.open(img_path)
    #             # Scale image (base size 100px, scaled)
    #             img_width = int(100 * scale)
    #             img.thumbnail((img_width, img_width))
                
    #             # Convert to PhotoImage
    #             photo = ImageTk.PhotoImage(img)
                
    #             # Convert relative to absolute coordinates
    #             abs_x = width * rel_x
    #             abs_y = height * rel_y
                
    #             # Draw on canvas
    #             img_id = canvas.create_image(abs_x, abs_y, image=photo)
    #             canvas.image = photo  # Keep reference
                
    #             # Make draggable
    #             canvas.tag_bind(img_id, "<Button1-Motion>", 
    #                 lambda e, i=img_id: self.move_element(e, canvas, i))
                
    #         except Exception as e:
    #             print(f"Error loading preview element: {str(e)}")
        
    #     # Main content preview
    #     canvas.create_text(width//2, 100, 
    #                     text="Certificate of Completion", 
    #                     font=("Helvetica", 24), 
    #                     fill="#333333")
        
    #     canvas.create_text(width//2, 160, 
    #                     text=self.name_var.get(), 
    #                     font=("Helvetica", 20), 
    #                     fill="#333333")
        
    #     canvas.create_text(width//2, 220, 
    #                     text=f"For successfully completing:", 
    #                     font=("Helvetica", 16), 
    #                     fill="#333333")
        
    #     canvas.create_text(width//2, 260, 
    #                     text=self.course_var.get(), 
    #                     font=("Helvetica", 18), 
    #                     fill="#2C3E50")
        
    #     if self.desc_var.get():
    #         canvas.create_text(width//2, 300, 
    #                         text=self.desc_var.get(), 
    #                         font=("Helvetica", 14), 
    #                         fill="#555555")
        
    #     canvas.create_text(width//2, 370, 
    #                     text=f"Completed on {self.date_var.get()}", 
    #                     font=("Helvetica", 14), 
    #                     fill="#777777")

    # def move_element(self, event, canvas, element_id):
    #     """Handle dragging of elements on the canvas"""
    #     # Get current canvas dimensions
    #     canvas_width = canvas.winfo_width()
    #     canvas_height = canvas.winfo_height()
        
    #     # Constrain position to canvas
    #     x = max(0, min(event.x, canvas_width))
    #     y = max(0, min(event.y, canvas_height))
        
    #     # Move the element
    #     canvas.coords(element_id, x, y)

    # # Add these methods to the CertificateGenerator class to support the customization
    # def setup_customization_ui(self):
    #     """Add UI elements for customizing the certificate"""
    #     # Add a new tab for customization
    #     self.tabview.add("Customize")
        
    #     # Orientation selection
    #     self.orientation_var = ctk.StringVar(value="Portrait")
    #     ctk.CTkLabel(
    #         self.tabview.tab("Customize"),
    #         text="Page Orientation:"
    #     ).grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        
    #     ctk.CTkOptionMenu(
    #         self.tabview.tab("Customize"),
    #         values=["Portrait", "Landscape"],
    #         variable=self.orientation_var,
    #         command=lambda _: self.generate_preview()
    #     ).grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")
        
    #     # Toggle default content
    #     self.show_default_content = ctk.BooleanVar(value=True)
    #     ctk.CTkCheckBox(
    #         self.tabview.tab("Customize"),
    #         text="Show Default Content",
    #         variable=self.show_default_content,
    #         command=self.generate_preview
    #     ).grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")
        
    #     # Add element buttons
    #     ctk.CTkButton(
    #         self.tabview.tab("Customize"),
    #         text="Add Text",
    #         command=self.add_text_element
    #     ).grid(row=3, column=0, padx=10, pady=(10, 5))
        
    #     ctk.CTkButton(
    #         self.tabview.tab("Customize"),
    #         text="Add Image",
    #         command=self.add_image_element
    #     ).grid(row=4, column=0, padx=10, pady=5)
        
    #     ctk.CTkButton(
    #         self.tabview.tab("Customize"),
    #         text="Add Watermark",
    #         command=self.add_watermark_element
    #     ).grid(row=5, column=0, padx=10, pady=5)
        
    #     ctk.CTkButton(
    #         self.tabview.tab("Customize"),
    #         text="Add Shape",
    #         command=self.add_shape_element
    #     ).grid(row=6, column=0, padx=10, pady=5)
        
    #     # Element list
    #     self.element_list_frame = ctk.CTkFrame(self.tabview.tab("Customize"))
    #     self.element_list_frame.grid(row=7, column=0, padx=10, pady=10, sticky="nsew")
        
    #     self.element_list_label = ctk.CTkLabel(
    #         self.element_list_frame,
    #         text="Added Elements:"
    #     )
    #     self.element_list_label.pack(pady=(0, 10))
        
    #     self.element_listbox = tk.Listbox(
    #         self.element_list_frame,
    #         width=30,
    #         height=8,
    #         bg="#f0f0f0" if ctk.get_appearance_mode() == "Light" else "#2b2b2b",
    #         fg="black" if ctk.get_appearance_mode() == "Light" else "white"
    #     )
    #     self.element_listbox.pack(fill="both", expand=True)
        
    #     # Element controls
    #     self.element_controls_frame = ctk.CTkFrame(self.tabview.tab("Customize"))
    #     self.element_controls_frame.grid(row=8, column=0, padx=10, pady=(0, 10), sticky="ew")
        
    #     ctk.CTkButton(
    #         self.element_controls_frame,
    #         text="Edit",
    #         width=80,
    #         command=self.edit_element
    #     ).pack(side="left", padx=5)
        
    #     ctk.CTkButton(
    #         self.element_controls_frame,
    #         text="Delete",
    #         width=80,
    #         command=self.delete_element
    #     ).pack(side="left", padx=5)
        
    #     # Initialize custom elements list
    #     self.custom_elements = []
        
    #     # Make the preview canvas draggable
    #     self.preview_canvas.bind("<Button-1>", self.start_drag)
    #     self.preview_canvas.bind("<B1-Motion>", self.on_drag)
    #     self.preview_canvas.bind("<ButtonRelease-1>", self.end_drag)
    #     self.dragging = False
    #     self.current_element = None

    # def add_text_element(self):
    #     """Add a new text element to the certificate"""
    #     dialog = ctk.CTkToplevel(self)
    #     dialog.title("Add Text Element")
    #     dialog.geometry("400x400")
    #     dialog.resizable(False, False)
        
    #     # Text content
    #     ctk.CTkLabel(dialog, text="Text:").pack(pady=(10, 0))
    #     text_entry = ctk.CTkEntry(dialog)
    #     text_entry.pack(pady=(0, 10))
        
    #     # Font selection
    #     ctk.CTkLabel(dialog, text="Font:").pack()
    #     font_var = ctk.StringVar(value="Helvetica")
    #     font_menu = ctk.CTkOptionMenu(
    #         dialog,
    #         values=["Helvetica", "Times-Roman", "Courier", "Helvetica-Bold", "Times-Bold"],
    #         variable=font_var
    #     )
    #     font_menu.pack(pady=(0, 10))
        
    #     # Font size
    #     ctk.CTkLabel(dialog, text="Font Size:").pack()
    #     size_slider = ctk.CTkSlider(dialog, from_=8, to=72, number_of_steps=32)
    #     size_slider.set(16)
    #     size_slider.pack(pady=(0, 10))
        
    #     # Color
    #     ctk.CTkLabel(dialog, text="Color:").pack()
    #     color_entry = ctk.CTkEntry(dialog, placeholder_text="#000000")
    #     color_entry.pack(pady=(0, 10))
        
    #     # Position
    #     ctk.CTkLabel(dialog, text="Position (x, y):").pack()
    #     pos_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    #     pos_frame.pack(pady=(0, 10))
        
    #     ctk.CTkLabel(pos_frame, text="X:").pack(side="left")
    #     x_entry = ctk.CTkEntry(pos_frame, width=60)
    #     x_entry.pack(side="left", padx=5)
    #     x_entry.insert(0, "100")
        
    #     ctk.CTkLabel(pos_frame, text="Y:").pack(side="left")
    #     y_entry = ctk.CTkEntry(pos_frame, width=60)
    #     y_entry.pack(side="left", padx=5)
    #     y_entry.insert(0, "100")
        
    #     def save_element():
    #         self.custom_elements.append({
    #             'type': 'text',
    #             'text': text_entry.get(),
    #             'font': font_var.get(),
    #             'size': int(size_slider.get()),
    #             'color': color_entry.get() or "#000000",
    #             'x': int(x_entry.get()),
    #             'y': int(y_entry.get()),
    #             'id': str(len(self.custom_elements) + 1)
    #         })
    #         self.update_element_list()
    #         dialog.destroy()
    #         self.generate_preview()
        
    #     ctk.CTkButton(
    #         dialog,
    #         text="Add Text",
    #         command=save_element
    #     ).pack(pady=10)

    # def add_image_element(self):
    #     """Add a new image element to the certificate"""
    #     file_path = filedialog.askopenfilename(
    #         title="Select Image",
    #         filetypes=[("Image Files", "*.png *.jpg *.jpeg *.gif")]
    #     )
        
    #     if file_path:
    #         dialog = ctk.CTkToplevel(self)
    #         dialog.title("Add Image Element")
    #         dialog.geometry("400x300")
    #         dialog.resizable(False, False)
            
    #         # Preview
    #         try:
    #             img = Image.open(file_path)
    #             img.thumbnail((200, 200))
    #             photo = ImageTk.PhotoImage(img)
                
    #             preview_label = ctk.CTkLabel(dialog, text="", image=photo)
    #             preview_label.image = photo
    #             preview_label.pack(pady=10)
    #         except Exception as e:
    #             ctk.CTkLabel(dialog, text="Image preview not available").pack(pady=10)
            
    #         # Dimensions
    #         ctk.CTkLabel(dialog, text="Dimensions:").pack()
    #         dim_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    #         dim_frame.pack(pady=(0, 10))
            
    #         ctk.CTkLabel(dim_frame, text="Width:").pack(side="left")
    #         width_entry = ctk.CTkEntry(dim_frame, width=60)
    #         width_entry.pack(side="left", padx=5)
    #         width_entry.insert(0, "200")
            
    #         ctk.CTkLabel(dim_frame, text="Height:").pack(side="left")
    #         height_entry = ctk.CTkEntry(dim_frame, width=60)
    #         height_entry.pack(side="left", padx=5)
    #         height_entry.insert(0, "100")
            
    #         # Position
    #         ctk.CTkLabel(dialog, text="Position (x, y):").pack()
    #         pos_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    #         pos_frame.pack(pady=(0, 10))
            
    #         ctk.CTkLabel(pos_frame, text="X:").pack(side="left")
    #         x_entry = ctk.CTkEntry(pos_frame, width=60)
    #         x_entry.pack(side="left", padx=5)
    #         x_entry.insert(0, "100")
            
    #         ctk.CTkLabel(pos_frame, text="Y:").pack(side="left")
    #         y_entry = ctk.CTkEntry(pos_frame, width=60)
    #         y_entry.pack(side="left", padx=5)
    #         y_entry.insert(0, "100")
            
    #         def save_element():
    #             self.custom_elements.append({
    #                 'type': 'image',
    #                 'path': file_path,
    #                 'width': int(width_entry.get()),
    #                 'height': int(height_entry.get()),
    #                 'x': int(x_entry.get()),
    #                 'y': int(y_entry.get()),
    #                 'id': str(len(self.custom_elements) + 1)
    #             })
    #             self.update_element_list()
    #             dialog.destroy()
    #             self.generate_preview()
            
    #         ctk.CTkButton(
    #             dialog,
    #             text="Add Image",
    #             command=save_element
    #         ).pack(pady=10)

    # def add_watermark_element(self):
    #     """Add a watermark text element"""
    #     dialog = ctk.CTkToplevel(self)
    #     dialog.title("Add Watermark")
    #     dialog.geometry("400x400")
    #     dialog.resizable(False, False)
        
    #     # Text content
    #     ctk.CTkLabel(dialog, text="Watermark Text:").pack(pady=(10, 0))
    #     text_entry = ctk.CTkEntry(dialog)
    #     text_entry.pack(pady=(0, 10))
        
    #     # Font selection
    #     ctk.CTkLabel(dialog, text="Font:").pack()
    #     font_var = ctk.StringVar(value="Helvetica")
    #     font_menu = ctk.CTkOptionMenu(
    #         dialog,
    #         values=["Helvetica", "Times-Roman", "Courier", "Helvetica-Bold", "Times-Bold"],
    #         variable=font_var
    #     )
    #     font_menu.pack(pady=(0, 10))
        
    #     # Font size
    #     ctk.CTkLabel(dialog, text="Font Size:").pack()
    #     size_slider = ctk.CTkSlider(dialog, from_=8, to=120, number_of_steps=56)
    #     size_slider.set(48)
    #     size_slider.pack(pady=(0, 10))
        
    #     # Color
    #     ctk.CTkLabel(dialog, text="Color (with alpha, e.g. #00000080):").pack()
    #     color_entry = ctk.CTkEntry(dialog, placeholder_text="#00000080")
    #     color_entry.pack(pady=(0, 10))
        
    #     # Rotation
    #     ctk.CTkLabel(dialog, text="Rotation (degrees):").pack()
    #     rotation_slider = ctk.CTkSlider(dialog, from_=0, to=360, number_of_steps=36)
    #     rotation_slider.set(45)
    #     rotation_slider.pack(pady=(0, 10))
        
    #     # Position
    #     ctk.CTkLabel(dialog, text="Position (x, y):").pack()
    #     pos_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    #     pos_frame.pack(pady=(0, 10))
        
    #     ctk.CTkLabel(pos_frame, text="X:").pack(side="left")
    #     x_entry = ctk.CTkEntry(pos_frame, width=60)
    #     x_entry.pack(side="left", padx=5)
    #     x_entry.insert(0, "200")
        
    #     ctk.CTkLabel(pos_frame, text="Y:").pack(side="left")
    #     y_entry = ctk.CTkEntry(pos_frame, width=60)
    #     y_entry.pack(side="left", padx=5)
    #     y_entry.insert(0, "300")
        
    #     def save_element():
    #         self.custom_elements.append({
    #             'type': 'watermark',
    #             'text': text_entry.get(),
    #             'font': font_var.get(),
    #             'size': int(size_slider.get()),
    #             'color': color_entry.get() or "#00000080",
    #             'rotation': int(rotation_slider.get()),
    #             'x': int(x_entry.get()),
    #             'y': int(y_entry.get()),
    #             'id': str(len(self.custom_elements) + 1)
    #         })
    #         self.update_element_list()
    #         dialog.destroy()
    #         self.generate_preview()
        
    #     ctk.CTkButton(
    #         dialog,
    #         text="Add Watermark",
    #         command=save_element
    #     ).pack(pady=10)

    # def add_shape_element(self):
    #     """Add a shape element (rectangle or circle)"""
    #     dialog = ctk.CTkToplevel(self)
    #     dialog.title("Add Shape")
    #     dialog.geometry("400x500")
    #     dialog.resizable(False, False)
        
    #     # Shape type
    #     ctk.CTkLabel(dialog, text="Shape Type:").pack(pady=(10, 0))
    #     shape_var = ctk.StringVar(value="rectangle")
    #     ctk.CTkOptionMenu(
    #         dialog,
    #         values=["rectangle", "circle"],
    #         variable=shape_var
    #     ).pack(pady=(0, 10))
        
    #     # Fill color
    #     ctk.CTkLabel(dialog, text="Fill Color:").pack()
    #     fill_color_entry = ctk.CTkEntry(dialog, placeholder_text="#FFFFFF00")
    #     fill_color_entry.pack(pady=(0, 10))
        
    #     # Stroke color
    #     ctk.CTkLabel(dialog, text="Stroke Color:").pack()
    #     stroke_color_entry = ctk.CTkEntry(dialog, placeholder_text="#000000")
    #     stroke_color_entry.pack(pady=(0, 10))
        
    #     # Stroke width
    #     ctk.CTkLabel(dialog, text="Stroke Width:").pack()
    #     stroke_slider = ctk.CTkSlider(dialog, from_=0, to=10, number_of_steps=10)
    #     stroke_slider.set(1)
    #     stroke_slider.pack(pady=(0, 10))
        
    #     # Dimensions
    #     ctk.CTkLabel(dialog, text="Dimensions:").pack()
    #     dim_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    #     dim_frame.pack(pady=(0, 10))
        
    #     ctk.CTkLabel(dim_frame, text="Width/Radius:").pack(side="left")
    #     width_entry = ctk.CTkEntry(dim_frame, width=80)
    #     width_entry.pack(side="left", padx=5)
    #     width_entry.insert(0, "100")
        
    #     ctk.CTkLabel(dim_frame, text="Height:").pack(side="left")
    #     height_entry = ctk.CTkEntry(dim_frame, width=80)
    #     height_entry.pack(side="left", padx=5)
    #     height_entry.insert(0, "50")
        
    #     # Position
    #     ctk.CTkLabel(dialog, text="Position (x, y):").pack()
    #     pos_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    #     pos_frame.pack(pady=(0, 10))
        
    #     ctk.CTkLabel(pos_frame, text="X:").pack(side="left")
    #     x_entry = ctk.CTkEntry(pos_frame, width=60)
    #     x_entry.pack(side="left", padx=5)
    #     x_entry.insert(0, "100")
        
    #     ctk.CTkLabel(pos_frame, text="Y:").pack(side="left")
    #     y_entry = ctk.CTkEntry(pos_frame, width=60)
    #     y_entry.pack(side="left", padx=5)
    #     y_entry.insert(0, "100")
        
    #     def save_element():
    #         self.custom_elements.append({
    #             'type': 'shape',
    #             'shape': shape_var.get(),
    #             'fill_color': fill_color_entry.get() or "#FFFFFF00",
    #             'stroke_color': stroke_color_entry.get() or "#000000",
    #             'stroke_width': int(stroke_slider.get()),
    #             'width': int(width_entry.get()),
    #             'height': int(height_entry.get()),
    #             'radius': int(width_entry.get()),  # For circles
    #             'x': int(x_entry.get()),
    #             'y': int(y_entry.get()),
    #             'fill': fill_color_entry.get() != "#FFFFFF00",
    #             'stroke': stroke_slider.get() > 0,
    #             'id': str(len(self.custom_elements) + 1)
    #         })
    #         self.update_element_list()
    #         dialog.destroy()
    #         self.generate_preview()
        
    #     ctk.CTkButton(
    #         dialog,
    #         text="Add Shape",
    #         command=save_element
    #     ).pack(pady=10)

    # def update_element_list(self):
    #     """Update the listbox with current elements"""
    #     self.element_listbox.delete(0, tk.END)
    #     for element in self.custom_elements:
    #         self.element_listbox.insert(tk.END, f"{element['type'].title()}: {element.get('text', element.get('path', element['shape'] if element['type'] == 'shape' else ''))}")

    # def edit_element(self):
    #     """Edit the selected element"""
    #     selection = self.element_listbox.curselection()
    #     if not selection:
    #         return
        
    #     index = selection[0]
    #     if 0 <= index < len(self.custom_elements):
    #         element = self.custom_elements[index]
            
    #         if element['type'] == 'text':
    #             self.edit_text_element(index)
    #         elif element['type'] == 'image':
    #             self.edit_image_element(index)
    #         elif element['type'] == 'watermark':
    #             self.edit_watermark_element(index)
    #         elif element['type'] == 'shape':
    #             self.edit_shape_element(index)

    # def delete_element(self):
    #     """Delete the selected element"""
    #     selection = self.element_listbox.curselection()
    #     if not selection:
    #         return
        
    #     index = selection[0]
    #     if 0 <= index < len(self.custom_elements):
    #         self.custom_elements.pop(index)
    #         self.update_element_list()
    #         self.generate_preview()

    # def start_drag(self, event):
    #     """Start interacting with an element"""
    #     # Find which element was clicked
    #     for i, element in enumerate(self.custom_elements):
    #         # Simple bounding box check
    #         if (element['x'] <= event.x <= element['x'] + element.get('width', 100) and 
    #             element['y'] <= event.y <= element['y'] + element.get('height', 50)):
    #             self.current_element = i
    #             self.drag_start_x = event.x
    #             self.drag_start_y = event.y
                
    #             # Determine interaction mode based on cursor position
    #             if (event.x > element['x'] + element.get('width', 100) - 15 and 
    #                 event.y > element['y'] + element.get('height', 50) - 15):
    #                 self.interaction_mode = 'resize'
    #             elif (event.x > element['x'] + element.get('width', 100)//2 - 15 and 
    #                 event.y < element['y'] + 15):
    #                 self.interaction_mode = 'rotate'
    #             else:
    #                 self.interaction_mode = 'move'
                
    #             self.config(cursor=self.cursors[self.interaction_mode])
    #             break

    # def on_drag(self, event):
    #     """Handle element interaction"""
    #     if self.current_element is not None and self.interaction_mode:
    #         element = self.custom_elements[self.current_element]
            
    #         dx = event.x - self.drag_start_x
    #         dy = event.y - self.drag_start_y
            
    #         if self.interaction_mode == 'move':
    #             element['x'] += dx
    #             element['y'] += dy
    #         elif self.interaction_mode == 'resize':
    #             element['width'] = max(20, element.get('width', 100) + dx)
    #             element['height'] = max(20, element.get('height', 50) + dy)
    #         elif self.interaction_mode == 'rotate':
    #             center_x = element['x'] + element.get('width', 100)/2
    #             center_y = element['y'] + element.get('height', 50)/2
    #             angle = math.degrees(math.atan2(event.y - center_y, event.x - center_x))
    #             element['rotation'] = angle
            
    #         self.drag_start_x = event.x
    #         self.drag_start_y = event.y
    #         self.generate_preview()

    # def end_drag(self, event):
    #     """End interaction operation"""
    #     self.current_element = None
    #     self.interaction_mode = None
    #     self.config(cursor='')
    #     self.generate_preview()

    # def check_hover(self, event):
    #     """Check if cursor is over an element and change cursor accordingly"""
    #     self.cursor_over_element = False
    #     for element in self.custom_elements:
    #         if (element['x'] <= event.x <= element['x'] + element.get('width', 100) and 
    #             element['y'] <= event.y <= element['y'] + element.get('height', 50)):
    #             self.cursor_over_element = True
                
    #             # Check if near resize handle
    #             if (event.x > element['x'] + element.get('width', 100) - 15 and 
    #                 event.y > element['y'] + element.get('height', 50) - 15):
    #                 self.config(cursor=self.cursors['resize'])
    #             # Check if near rotate handle
    #             elif (event.x > element['x'] + element.get('width', 100)//2 - 15 and 
    #                 event.y < element['y'] + 15):
    #                 self.config(cursor=self.cursors['rotate'])
    #             else:
    #                 self.config(cursor=self.cursors['move'])
    #             break
        
    #     if not self.cursor_over_element:
    #         self.config(cursor='')

    # def draw_handles(self, element):
    #     """Draw resize and rotation handles for selected element"""
    #     # Resize handle (bottom-right corner)
    #     x1 = element['x'] + element['width'] - 5
    #     y1 = element['y'] + element['height'] - 5
    #     x2 = element['x'] + element['width'] + 5
    #     y2 = element['y'] + element['height'] + 5
    #     self.preview_canvas.create_rectangle(x1, y1, x2, y2, fill="red", outline="black")
        
    #     # Rotation handle (top-center)
    #     rot_x = element['x'] + element['width']//2
    #     rot_y = element['y'] - 15
    #     self.preview_canvas.create_oval(
    #         rot_x-5, rot_y-5, 
    #         rot_x+5, rot_y+5, 
    #         fill="blue", outline="black"
    #     )

    def generate_modern_certificate(self, output, preview=False):
        """Generate a modern-style certificate"""
        # Get field values
        name = self.name_var.get()
        course = self.course_var.get()
        try:
            raw_date = self.date_var.get()
            date_obj = datetime.strptime(raw_date, "%Y-%m-%d")
            date = date_obj.strftime("%B %d, %Y")
        except ValueError:
            date = raw_date
        description = self.desc_var.get()
        
        # Create PDF canvas
        c = canvas.Canvas(output, pagesize=A4)
        width, height = A4
        
        # Convert all dimensions to integers
        width = int(width)
        height = int(height)
        
        # Modern gradient background
        c.linearGradient(
            x0=0, y0=0, 
            x1=width, y1=height,
            colors=[HexColor("#FFFFFF"), HexColor("#F8F9FA")],
            positions=[0, 1]
        )
        
        # Header with color block - ensure integer dimensions
        header_height = 150
        c.setFillColor(HexColor("#2C3E50"))
        c.rect(0, height-header_height, width, header_height, fill=True, stroke=False)
        
        # Decorative accent
        c.setFillColor(HexColor("#E74C3C"))
        c.rect(0, height-header_height, width, 8, fill=True, stroke=False)
        
        # Title with modern typography
        c.setFont("Helvetica-Bold", 34)
        c.setFillColor(HexColor("#FFFFFF"))
        c.drawCentredString(width//2, height-80, "CERTIFICATE")
        
        # Subtitle
        c.setFont("Helvetica", 14)
        c.setFillColor(HexColor("#BDC3C7"))
        c.drawCentredString(width//2, height-110, "OF PROFESSIONAL ACHIEVEMENT")
        
        # Main content area with subtle shadow - ensure integer dimensions
        content_x = 40
        content_y = 180
        content_width = width-80
        content_height = height-380
        c.setFillColor(HexColor("#FFFFFF"))
        c.rect(content_x, content_y, content_width, content_height, fill=True, stroke=False)
        
        # Recipient name
        c.setFillColor(HexColor("#2C3E50"))
        c.setFont("Helvetica-Bold", 28)
        c.drawCentredString(width//2, height-200, name)
        
        # Course description
        c.setFont("Helvetica", 16)
        c.setFillColor(HexColor("#333333"))
        
        text = f"has successfully completed the {course} program"
        if description:
            text += f" with demonstrated excellence in {description}"
            
        p = Paragraph(text, ParagraphStyle(
            name="Normal", 
            fontSize=16, 
            leading=22,
            alignment=1
        ))
        p.wrapOn(c, width-200, 100)
        p.drawOn(c, 100, height-280)
        
        # Achievement statement
        c.setFont("Helvetica", 14)
        c.setFillColor(HexColor("#7F8C8D"))
        c.drawCentredString(width//2, height-350, "in recognition of outstanding performance and dedication")
        
        # Date
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(HexColor("#E74C3C"))
        c.drawCentredString(width//2, height-390, f"Completed on: {date}")
        
        # Logo and signature
        y_pos = height-480
        if self.logo_path:
            try:
                logo = ImageReader(self.logo_path)
                c.drawImage(logo, 100, y_pos, width=120, height=80, preserveAspectRatio=True, mask='auto')
                c.setStrokeColor(HexColor("#BDC3C7"))
                c.setLineWidth(0.5)
                c.line(100, y_pos-10, 220, y_pos-10)
                c.setFont("Helvetica", 10)
                c.drawString(100, y_pos-25, "Issuing Organization")
            except Exception as e:
                print(f"Error loading logo: {str(e)}")
                
        if self.signature_path:
            try:
                signature = ImageReader(self.signature_path)
                c.drawImage(signature, width-250, y_pos, width=150, height=60, preserveAspectRatio=True, mask='auto')
                c.setStrokeColor(HexColor("#BDC3C7"))
                c.setLineWidth(0.5)
                c.line(width-250, y_pos-10, width-100, y_pos-10)
                c.setFont("Helvetica", 10)
                c.drawCentredString(width-175, y_pos-25, "Authorized Signatory")
            except Exception as e:
                print(f"Error loading signature: {str(e)}")
        
        # Verification ID
        c.setFont("Helvetica", 8)
        c.setFillColor(HexColor("#95A5A6"))
        verification_id = f"ID: {datetime.now().strftime('%Y%m%d')}-{hash(name) % 10000:04d}"
        c.drawRightString(width-40, 40, verification_id)

        # Generate verification data
        verification_data = f"""
        Certificate Verification
        Name: {name}
        Course: {course}
        Date: {date}
        ID: {datetime.now().strftime('%Y%m%d')}-{hash(name) % 10000:04d}
        """
        
        # Generate and add QR code
        qr_img = self.generate_qr_code(verification_data, size=100)
        if qr_img:
            # Save QR code to a temporary buffer
            qr_buffer = BytesIO()
            qr_img.save(qr_buffer, format='PNG')
            qr_buffer.seek(0)
            
            # Add QR code to PDF
            qr_reader = ImageReader(qr_buffer)
            c.drawImage(qr_reader, width-120, 50, width=80, height=80, preserveAspectRatio=True, mask='auto')
        
        # Footer
        c.setFont("Helvetica", 9)
        c.setFillColor(HexColor("#7F8C8D"))
        c.drawCentredString(width//2, 30, "© " + datetime.now().strftime("%Y") + " Professional Certification Board. All rights reserved.")
        
        if preview:
            c.showPage()
            c.save()
            return output
        
        c.save()

    def generate_academic_diploma(self, output, preview=False):
        """Generate an academic diploma-style certificate"""
        name = self.name_var.get()
        course = self.course_var.get()
        try:
            raw_date = self.date_var.get()
            date_obj = datetime.strptime(raw_date, "%Y-%m-%d")
            date = date_obj.strftime("%B %d, %Y")
        except ValueError:
            date = raw_date
        description = self.desc_var.get()
        
        # Create PDF canvas
        c = canvas.Canvas(output, pagesize=landscape(A4))
        width, height = landscape(A4)
        
        # Parchment-style background
        c.setFillColor(HexColor("#FDF5E6"))
        c.rect(0, 0, width, height, fill=True, stroke=False)
        
        # Add subtle texture
        c.setFillColor(HexColor("#FAEBD7"))
        for i in range(0, width, 3):
            for j in range(0, height, 3):
                if (i + j) % 6 == 0:
                    c.rect(i, j, 2, 2, fill=True, stroke=False)
        
        # Ornate border
        border_color = HexColor("#8B4513")
        c.setStrokeColor(border_color)
        c.setLineWidth(8)
        c.roundRect(40, 40, width-80, height-80, radius=15, fill=False, stroke=True)
        
        # University-style seal at top
        c.setFillColor(HexColor("#8B4513"))
        c.circle(width//2, height-100, 50, fill=True, stroke=False)
        c.setFillColor(HexColor("#FDF5E6"))
        c.setFont("Times-Bold", 16)
        c.drawCentredString(width//2, height-100, "SEAL")
        
        # Title with academic styling
        c.setFillColor(HexColor("#8B4513"))
        c.setFont("Times-Bold", 36)
        c.drawCentredString(width//2, height-180, "DIPLOMA")
        
        # Latin motto
        c.setFont("Times-Italic", 12)
        c.drawCentredString(width//2, height-210, "Scientia est potentia")
        
        # Main content
        c.setFont("Times-Roman", 18)
        c.setFillColor(HexColor("#000000"))
        c.drawCentredString(width//2, height-270, "This certifies that")
        
        c.setFont("Times-Bold", 28)
        c.setFillColor(HexColor("#8B4513"))
        c.drawCentredString(width//2, height-320, name.upper())
        
        c.setFont("Times-Roman", 18)
        c.setFillColor(HexColor("#000000"))
        text = f"has satisfactorily completed all requirements for"
        p = Paragraph(text, ParagraphStyle(name="Normal", fontSize=18, leading=22, alignment=1))
        p.wrapOn(c, width-200, 50)
        p.drawOn(c, 100, height-370)
        
        c.setFont("Times-Bold", 22)
        c.drawCentredString(width//2, height-410, course)
        
        if description:
            c.setFont("Times-Roman", 16)
            p = Paragraph(description, ParagraphStyle(name="Normal", fontSize=16, leading=20, alignment=1))
            p.wrapOn(c, width-200, 100)
            p.drawOn(c, 100, height-450)
        
        # Date and signatures
        c.setFont("Times-Roman", 16)
        c.drawCentredString(width//2, height-500, f"Given this {date}")
        
        # Signature lines
        c.setStrokeColor(border_color)
        c.setLineWidth(1)
        c.line(width//4, height-550, width//4+200, height-550)
        c.line(3*width//4-200, height-550, 3*width//4, height-550)
        
        c.setFont("Times-Roman", 12)
        c.drawCentredString(width//4+100, height-570, "Dean of Studies")
        c.drawCentredString(3*width//4-100, height-570, "University President")
        
        # Generate verification data
        verification_data = f"""
        Diploma Verification
        Name: {name}
        Program: {course}
        Date: {date}
        ID: {datetime.now().strftime('%Y%m%d')}-{hash(name) % 10000:04d}
        """
        
        # Generate and add QR code
        qr_img = self.generate_qr_code(verification_data, size=80)
        if qr_img:
            qr_buffer = BytesIO()
            qr_img.save(qr_buffer, format='PNG')
            qr_buffer.seek(0)
            qr_reader = ImageReader(qr_buffer)
            c.drawImage(qr_reader, width-120, 50, width=80, height=80, preserveAspectRatio=True, mask='auto')

        if preview:
            c.showPage()
            c.save()
            return output
        
        c.save()

    def generate_corporate_certificate(self, output, preview=False):
        """Generate a corporate training certificate"""
        name = self.name_var.get()
        course = self.course_var.get()
        try:
            raw_date = self.date_var.get()
            date_obj = datetime.strptime(raw_date, "%Y-%m-%d")
            date = date_obj.strftime("%B %d, %Y")
        except ValueError:
            date = raw_date
        description = self.desc_var.get()
        
        # Create PDF canvas
        c = canvas.Canvas(output, pagesize=A4)
        width, height = A4
        
        # Corporate blue background
        c.setFillColor(HexColor("#E6F2FF"))
        c.rect(0, 0, width, height, fill=True, stroke=False)
        
        # Header with company branding
        c.setFillColor(HexColor("#003366"))
        c.rect(0, height-100, width, 100, fill=True, stroke=False)
        
        # Logo area
        if self.logo_path:
            try:
                logo = ImageReader(self.logo_path)
                c.drawImage(logo, width-150, height-90, width=120, height=80, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                print(f"Error loading logo: {str(e)}")
        
        c.setFont("Helvetica-Bold", 24)
        c.setFillColor(HexColor("#FFFFFF"))
        c.drawString(50, height-60, "CORPORATE TRAINING CERTIFICATION")
        
        # Certificate number
        c.setFont("Helvetica", 10)
        cert_id = f"CERT-{datetime.now().strftime('%Y%m%d')}-{hash(name) % 10000:04d}"
        c.drawRightString(width-50, height-70, cert_id)
        
        # Main content
        c.setFillColor(HexColor("#000000"))
        c.setFont("Helvetica", 16)
        c.drawCentredString(width//2, height-180, "This is to certify that")
        
        c.setFont("Helvetica-Bold", 24)
        c.setFillColor(HexColor("#003366"))
        c.drawCentredString(width//2, height-230, name.upper())
        
        c.setFont("Helvetica", 16)
        c.setFillColor(HexColor("#000000"))
        text = f"has successfully completed the corporate training program:"
        p = Paragraph(text, ParagraphStyle(name="Normal", fontSize=16, leading=20, alignment=1))
        p.wrapOn(c, width-200, 50)
        p.drawOn(c, 100, height-280)
        
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(width//2, height-320, course)
        
        if description:
            c.setFont("Helvetica", 14)
            p = Paragraph(description, ParagraphStyle(name="Normal", fontSize=14, leading=18, alignment=1))
            p.wrapOn(c, width-200, 100)
            p.drawOn(c, 100, height-360)
        
        # Completion details
        c.setFont("Helvetica", 14)
        c.drawCentredString(width//2, height-420, f"Date of Completion: {date}")
        
        # Signature area
        c.setStrokeColor(HexColor("#003366"))
        c.setLineWidth(1)
        c.line(width//3, height-480, width//3+200, height-480)
        c.line(2*width//3-200, height-480, 2*width//3, height-480)
        
        c.setFont("Helvetica", 12)
        c.drawCentredString(width//3+100, height-500, "Training Manager")
        c.drawCentredString(2*width//3-100, height-500, "HR Director")
        
        # Generate verification data
        verification_data = f"""
        Corporate Certification
        Name: {name}
        Training: {course}
        Completed: {date}
        ID: {datetime.now().strftime('%Y%m%d')}-{hash(name) % 10000:04d}
        """
        
        # Generate and add QR code
        qr_img = self.generate_qr_code(verification_data, size=80)
        if qr_img:
            qr_buffer = BytesIO()
            qr_img.save(qr_buffer, format='PNG')
            qr_buffer.seek(0)
            qr_reader = ImageReader(qr_buffer)
            c.drawImage(qr_reader, width-120, 50, width=80, height=80, preserveAspectRatio=True, mask='auto')

        # Footer
        c.setFont("Helvetica", 10)
        c.setFillColor(HexColor("#666666"))
        c.drawCentredString(width//2, 50, "This certificate verifies completion of required training hours")
        c.drawCentredString(width//2, 30, "and demonstration of competency in the subject matter.")
        
        if preview:
            c.showPage()
            c.save()
            return output
        
        c.save()

    def generate_workshop_certificate(self, output, preview=False):
        """Generate a workshop participation certificate"""
        name = self.name_var.get()
        course = self.course_var.get()
        try:
            raw_date = self.date_var.get()
            date_obj = datetime.strptime(raw_date, "%Y-%m-%d")
            date = date_obj.strftime("%B %d, %Y")
        except ValueError:
            date = raw_date
        description = self.desc_var.get()
        
        # Create PDF canvas
        c = canvas.Canvas(output, pagesize=A4)
        width, height = A4
        
        # Colorful modern background
        colors = [HexColor("#FF9AA2"), HexColor("#FFB7B2"), HexColor("#FFDAC1"), 
                 HexColor("#E2F0CB"), HexColor("#B5EAD7"), HexColor("#C7CEEA")]
        
        for i in range(6):
            c.setFillColor(colors[i])
            c.rect(0, height//6*i, width, height//6, fill=True, stroke=False)
        
        # White content area
        c.setFillColor(HexColor("#FFFFFF"))
        c.setStrokeColor(HexColor("#DDDDDD"))
        c.setLineWidth(1)
        c.roundRect(40, 40, width-80, height-80, 10, fill=True, stroke=True)
        
        # Title
        c.setFillColor(HexColor("#333333"))
        c.setFont("Helvetica-Bold", 28)
        c.drawCentredString(width//2, height-100, "WORKSHOP PARTICIPATION")
        
        # Main content
        c.setFont("Helvetica", 16)
        c.drawCentredString(width//2, height-160, "This certificate is presented to")
        
        c.setFont("Helvetica-Bold", 24)
        c.setFillColor(HexColor("#FF6B6B"))
        c.drawCentredString(width//2, height-210, name)
        
        c.setFont("Helvetica", 16)
        c.setFillColor(HexColor("#333333"))
        text = f"for active participation in the workshop:"
        p = Paragraph(text, ParagraphStyle(name="Normal", fontSize=16, leading=20, alignment=1))
        p.wrapOn(c, width-200, 50)
        p.drawOn(c, 100, height-260)
        
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(width//2, height-300, course)
        
        if description:
            c.setFont("Helvetica", 14)
            p = Paragraph(description, ParagraphStyle(name="Normal", fontSize=14, leading=18, alignment=1))
            p.wrapOn(c, width-200, 100)
            p.drawOn(c, 100, height-340)
        
        # Date and location
        c.setFont("Helvetica", 14)
        c.drawCentredString(width//2, height-390, f"Completed on {date}")
        
        # Signature area
        if self.signature_path:
            try:
                signature = ImageReader(self.signature_path)
                c.drawImage(signature, width//2-75, height-450, width=150, height=60, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                print(f"Error loading signature: {str(e)}")
        
        c.setStrokeColor(HexColor("#AAAAAA"))
        c.setLineWidth(0.5)
        c.line(width//2-100, height-450, width//2+100, height-450)
        c.setFont("Helvetica", 12)
        c.drawCentredString(width//2, height-480, "Workshop Facilitator")
        
        # Verification QR code placeholder
        c.setFillColor(HexColor("#EEEEEE"))
        c.rect(width-100, 50, 80, 80, fill=True, stroke=False)
        c.setFillColor(HexColor("#999999"))
        c.setFont("Helvetica", 8)
        c.drawCentredString(width-60, 70, "VERIFICATION")
        c.drawCentredString(width-60, 60, "QR CODE")

        # Generate and add actual QR code
        verification_data = f"""
        Workshop Certificate Verification
        Name: {name}
        Workshop: {course}
        Date: {date}
        """
        qr_img = self.generate_qr_code(verification_data, size=80)
        if qr_img:
            qr_buffer = BytesIO()
            qr_img.save(qr_buffer, format='PNG')
            qr_buffer.seek(0)
            qr_reader = ImageReader(qr_buffer)
            c.drawImage(qr_reader, width-100, 50, width=80, height=80, preserveAspectRatio=True, mask='auto')
        
        if preview:
            c.showPage()
            c.save()
            return output
        
        c.save()

if __name__ == "__main__":
    app = CertificateGenerator()
    app.mainloop()