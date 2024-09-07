import os
import pandas as pd
from photoshop import Session
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, colorchooser
import tkinter.font as tkfont
import qrcode
from PIL import Image
import photoshop.api as ps

class PhotoshopAutomationApp:
    def __init__(self, master):
        self.master = master
        master.title("Photoshop Automation")
        master.geometry("550x700")

        style = ttk.Style("darkly")
        
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        self.tab1 = ttk.Frame(self.notebook, padding="20")
        self.tab2 = ttk.Frame(self.notebook, padding="20")

        self.notebook.add(self.tab1, text="File Selection")
        self.notebook.add(self.tab2, text="Settings")

        self.create_tab1()
        self.create_tab2()

    def create_tab1(self):

        # Title
        ttk.Label(self.tab1, text="Photoshop Automation", font=("", 24, "bold")).pack(pady=(0, 20))

        # File Selection Frame
        file_frame = ttk.LabelFrame(self.tab1, text="File Selection", padding="10")
        file_frame.pack(fill=X, pady=10)

        # CSV File Selection
        self.csv_path = ttk.StringVar()
        ttk.Label(file_frame, text="CSV File:").grid(row=0, column=0, sticky=W, pady=5)
        ttk.Entry(file_frame, textvariable=self.csv_path, width=30).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.select_csv, style='Outline.TButton').grid(row=0, column=2, pady=5)

        # PSD File Selection
        self.psd_path = ttk.StringVar()
        ttk.Label(file_frame, text="PSD File:").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(file_frame, textvariable=self.psd_path, width=30).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.select_psd, style='Outline.TButton').grid(row=1, column=2, pady=5)

        # Output Folder Selection
        self.output_path = ttk.StringVar()
        ttk.Label(file_frame, text="Output Folder:").grid(row=2, column=0, sticky=W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_path, width=30).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.select_output_folder, style='Outline.TButton').grid(row=2, column=2, pady=5)

        # Row Range Selection Frame
        row_range_frame = ttk.LabelFrame(self.tab1, text="Row Range Selection", padding="10")
        row_range_frame.pack(fill=X, pady=10)

        # Start Row Selection
        ttk.Label(row_range_frame, text="Start Row:").grid(row=0, column=0, sticky=W, pady=5)
        self.start_row = ttk.IntVar(value=1)
        ttk.Spinbox(row_range_frame, from_=1, to=10000, textvariable=self.start_row, width=5).grid(row=0, column=1, padx=5, pady=5)

        # End Row Selection
        ttk.Label(row_range_frame, text="End Row:").grid(row=1, column=0, sticky=W, pady=5)
        self.end_row = ttk.IntVar(value=10000)
        ttk.Spinbox(row_range_frame, from_=1, to=10000, textvariable=self.end_row, width=5).grid(row=1, column=1, padx=5, pady=5)

        # Next Button
        ttk.Button(self.tab1, text="Next", command=lambda: self.notebook.select(1), style='primary.TButton').pack(pady=20)
        
    def create_tab2(self):
        # Text Settings Frame
        text_settings_frame = ttk.LabelFrame(self.tab2, text="Text Settings", padding="10")
        text_settings_frame.pack(fill=X, pady=10)

        # Output Format Selection
        ttk.Label(text_settings_frame, text="Output Format:").grid(row=0, column=0, sticky=W, pady=5)
        self.output_format = ttk.StringVar(value="PSD")
        format_combo = ttk.Combobox(text_settings_frame, textvariable=self.output_format, 
                                    values=["PSD", "PDF", "JPEG"], width=10)
        format_combo.grid(row=0, column=1, padx=5, pady=5)

        # Font Selection
        ttk.Label(text_settings_frame, text="Font:").grid(row=1, column=0, sticky=W, pady=5)
        self.font_var = ttk.StringVar()
        self.font_combo = ttk.Combobox(text_settings_frame, textvariable=self.font_var, width=20)
        self.font_combo.grid(row=1, column=1, padx=5, pady=5)
        self.populate_fonts()

        # Font Size Selection
        ttk.Label(text_settings_frame, text="Font Size:").grid(row=2, column=0, sticky=W, pady=5)
        self.font_size_var = ttk.IntVar(value=26)
        font_size_spin = ttk.Spinbox(text_settings_frame, from_=8, to=72, textvariable=self.font_size_var, width=5)
        font_size_spin.grid(row=2, column=1, padx=5, pady=5, sticky=W)

        # Font Style Selection (Bold, Italic, etc.)
        ttk.Label(text_settings_frame, text="Font Style:").grid(row=3, column=0, sticky=W, pady=5)
        self.font_style_var = ttk.StringVar(value="Bold")
        font_style_combo = ttk.Combobox(text_settings_frame, textvariable=self.font_style_var, 
                                        values=["Regular", "Bold", "Italic", "Bold Italic"], width=15)
        font_style_combo.grid(row=3, column=1, padx=5, pady=5)

        # Color Selection
        ttk.Label(text_settings_frame, text="Text Color:").grid(row=4, column=0, sticky=W, pady=5)
        self.text_color_var = ttk.StringVar(value="#000000")
        ttk.Entry(text_settings_frame, textvariable=self.text_color_var, width=10).grid(row=4, column=1, padx=5, pady=5, sticky=W)
        ttk.Button(text_settings_frame, text="Choose", command=self.select_color, style='Outline.TButton').grid(row=4, column=2, padx=5, pady=5)

        # Text Location Selection
        ttk.Label(text_settings_frame, text="Text Location:").grid(row=5, column=0, sticky=W, pady=5)
        location_frame = ttk.Frame(text_settings_frame)
        location_frame.grid(row=5, column=1, sticky=W, pady=5)
        
        self.text_x_var = ttk.IntVar(value=4000)
        ttk.Label(location_frame, text="X:").pack(side=LEFT)
        ttk.Spinbox(location_frame, from_=0, to=10000, textvariable=self.text_x_var, width=5).pack(side=LEFT, padx=(0, 10))
        
        self.text_y_var = ttk.IntVar(value=2635)
        ttk.Label(location_frame, text="Y:").pack(side=LEFT)
        ttk.Spinbox(location_frame, from_=0, to=10000, textvariable=self.text_y_var, width=5).pack(side=LEFT)

        # QR Settings Frame
        qr_settings_frame = ttk.LabelFrame(self.tab2, text="QR Settings", padding="10")
        qr_settings_frame.pack(fill=X, pady=10)

        # QR Size Selection
        ttk.Label(qr_settings_frame, text="QR Size:").grid(row=0, column=0, sticky=W, pady=5)
        self.qr_size_var = ttk.IntVar(value=1750)
        qr_size_spin = ttk.Spinbox(qr_settings_frame, from_=50, to=500, textvariable=self.qr_size_var, width=5)
        qr_size_spin.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        # QR Location Selection
        ttk.Label(qr_settings_frame, text="QR Location:").grid(row=1, column=0, sticky=W, pady=5)
        qr_location_frame = ttk.Frame(qr_settings_frame)
        qr_location_frame.grid(row=1, column=1, sticky=W, pady=5)
        
        self.qr_x_var = ttk.IntVar(value=140)
        ttk.Label(qr_location_frame, text="X:").pack(side=LEFT)
        ttk.Spinbox(qr_location_frame, from_=0, to=10000, textvariable=self.qr_x_var, width=5).pack(side=LEFT, padx=(0, 10))
        
        self.qr_y_var = ttk.IntVar(value=1625)
        ttk.Label(qr_location_frame, text="Y:").pack(side=LEFT)
        ttk.Spinbox(qr_location_frame, from_=0, to=10000, textvariable=self.qr_y_var, width=5).pack(side=LEFT)

        # PNG QR without background option
        self.png_qr_var = ttk.BooleanVar(value=False)
        ttk.Checkbutton(qr_settings_frame, text="PNG QR without background", variable=self.png_qr_var).grid(row=2, column=0, columnspan=2, sticky=W, pady=5)

        
        # Navigation buttons
        button_frame = ttk.Frame(self.tab2)
        button_frame.pack(fill=X, pady=20)
        ttk.Button(button_frame, text="Back", command=lambda: self.notebook.select(0), style='secondary.TButton').pack(side=LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Run Automation", command=self.run_automation, style='success.TButton').pack(side=LEFT)

        # Status
        self.status_var = ttk.StringVar(value="Ready")
        ttk.Label(self.tab2, textvariable=self.status_var).pack()

    def populate_fonts(self):
        fonts = sorted(tkfont.families())
        self.font_combo['values'] = fonts
        if fonts:
            self.font_var.set(fonts[0])

    def select_color(self):
        color_code = colorchooser.askcolor(title="Choose text color")[1]
        if color_code:
            self.text_color_var.set(color_code)

    def select_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.csv_path.set(file_path)
   
    def select_psd(self):
        file_path = filedialog.askopenfilename(filetypes=[("PSD files", "*.psd")])
        if file_path:
            self.psd_path.set(file_path)

    def select_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_path.set(folder_path)

    def generate_qr_code(self, data, size, transparent=False):
        qr = qrcode.QRCode(version=1, box_size=10, border=0)
        qr.add_data(data)
        qr.make(fit=True)
        
        if transparent:
            img = qr.make_image(fill_color="black", back_color="transparent")
            img = img.convert("RGBA")
        else:
            img = qr.make_image(fill_color="black", back_color="white")
        
        img = img.resize((size, size))
        return img

    def run_automation(self):
        csv_file_path = self.csv_path.get()
        photoshop_file_path = self.psd_path.get()
        output_folder_path = self.output_path.get()
        output_format = self.output_format.get()
        font = self.font_var.get()
        font_size = self.font_size_var.get()
        font_style = self.font_style_var.get()
        color_hex = self.text_color_var.get()
        qr_size = self.qr_size_var.get()
        qr_x = self.qr_x_var.get()
        qr_y = self.qr_y_var.get()
        start_row = self.start_row.get()
        end_row = self.end_row.get()

        if not csv_file_path or not photoshop_file_path or not output_folder_path:
            messagebox.showerror("Error", "Please select CSV file, PSD file, and output folder.")
            return

        self.status_var.set("Processing...")
        self.master.update_idletasks()

        try:
            df = pd.read_csv(csv_file_path)

            # Adjust row range based on user input
            start_index = start_row - 1  # Convert to 0-based index
            end_index = min(end_row, len(df))  # Ensure we don't exceed the dataframe length
            df_subset = df.iloc[start_index:end_index]

            with Session() as ps:
                doc = ps.app.open(photoshop_file_path)

                for index, row in df_subset.iterrows():
                    # Add text layer
                    text_layer = doc.artLayers.add()
                    text_layer.kind = ps.LayerKind.TextLayer
                    text_layer.textItem.contents = row['doctor']
                    text_layer.textItem.position = [self.text_x_var.get(), self.text_y_var.get()]
                    text_layer.name = f'TextLayer_{index}'
                    
                    # Apply the selected font, size, and style
                    text_layer.textItem.font = font
                    text_layer.textItem.size = font_size

                    # Applying font style (bold, italic, etc.)
                    if font_style == "Bold":
                        text_layer.textItem.fauxBold = True
                    elif font_style == "Italic":
                        text_layer.textItem.fauxItalic = True
                    elif font_style == "Bold Italic":
                        text_layer.textItem.fauxBold = True
                        text_layer.textItem.fauxItalic = True
                    else:
                        text_layer.textItem.fauxBold = False
                        text_layer.textItem.fauxItalic = False

                    text_layer.textItem.leading = 12
                    text_layer.textItem.tracking = 10

                    # Set text color based on user selection
                    color = ps.SolidColor()
                    color.rgb.red = int(color_hex[1:3], 16)
                    color.rgb.green = int(color_hex[3:5], 16)
                    color.rgb.blue = int(color_hex[5:7], 16)
                    text_layer.textItem.color = color

                    text_layer.textItem.justification = ps.Justification.Center

                    output_folder_path = os.path.join(os.path.dirname(csv_file_path), 'UpdatedFile')
                    os.makedirs(output_folder_path, exist_ok=True)

                    # Generate and add QR code
                    qr_img = self.generate_qr_code(row['link'], qr_size, self.png_qr_var.get())
                    temp_qr_path = os.path.join(output_folder_path, f'temp_qr_{index}.png')
                    qr_img.save(temp_qr_path, "PNG")

                    # Place QR code in Photoshop
                    doc.artLayers.add()
                    ps.app.load(temp_qr_path)
                    ps.app.activeDocument.selection.selectAll()
                    ps.app.activeDocument.selection.copy()
                    ps.app.activeDocument.close(ps.SaveOptions.DoNotSaveChanges)
                    doc.paste()
                    qr_layer = doc.activeLayer
                    qr_layer.name = f'QRLayer_{index}'
                    qr_layer.translate(qr_x - qr_layer.bounds[0], qr_y - qr_layer.bounds[1])
                    
                    os.remove(temp_qr_path)  # Remove temporary QR code file

                    output_file_path = os.path.join(output_folder_path, f'{row["doctor"]}.{output_format.lower()}')

                    # Save based on the selected format
                    if output_format == "PSD":
                        save_options = ps.PhotoshopSaveOptions()
                    elif output_format == "PDF":
                        save_options = ps.PDFSaveOptions()
                    elif output_format == "JPEG":
                        save_options = ps.JPEGSaveOptions(quality=12)

                    doc.saveAs(output_file_path, save_options, True)
                    text_layer.remove()
                    qr_layer.remove()

                doc.close(ps.SaveOptions.DoNotSaveChanges)

            self.status_var.set("Completed successfully!")
            messagebox.showinfo("Success", "Photoshop automation completed successfully.")
        except Exception as e:
            self.status_var.set("Error occurred")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    root = ttk.Window()
    app = PhotoshopAutomationApp(root)
    root.mainloop()