import os
import pandas as pd
from photoshop import Session
import tkinter as tk
from tkinter import filedialog, messagebox

#Select folder function
def select_folder(title):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title=title)
    return folder_path
folder_path = select_folder(title="Select the PhotoshopAutomation folder")

def select_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return file_path

#Input CSV and Photoshop file
def automate_photoshop():
    csv_file_path = select_file("Select CSV File", [("CSV files", "*.csv")])
    if not csv_file_path:
        messagebox.showerror("Error", "No CSV file selected")
        return
    photoshop_file_path = select_file("Select Photoshop File", [("PSD files", "*.psd")])
    if not photoshop_file_path:
        messagebox.showerror("Error", "No Photoshop file selected")
        return

    #Read data from CSV
    df = pd.read_csv(csv_file_path)

    with Session() as ps:
        doc = ps.app.open(photoshop_file_path)

        for index, row in df.iterrows():
            text_layer = doc.artLayers.add()
            text_layer.kind = ps.LayerKind.TextLayer
            text_layer.textItem.contents = row['doctor']
            text_layer.textItem.position = [2235, 950]
            text_layer.name = f'TextLayer_{index}'
            text_layer.textItem.font = "IRYekan"
            text_layer.textItem.size = 24
            text_layer.textItem.leading = 12
            text_layer.textItem.tracking = 10
            color = ps.SolidColor()
            color.rgb.red = 63
            color.rgb.green = 162
            color.rgb.blue = 207
            text_layer.textItem.color = color
            text_layer.textItem.justification = ps.Justification.Center

            #Save PSD files
            output_folder_path = os.path.join(os.path.dirname(csv_file_path), 'UpdatedFile')
            os.makedirs(output_folder_path, exist_ok=True)
            output_file_path = os.path.join(output_folder_path, f'{row["doctor"]}.psd')
            doc.saveAs(output_file_path, ps.PhotoshopSaveOptions(), True)

            #Remove text layer
            text_layer.remove()

        #Close the document
        doc.close(ps.SaveOptions.DoNotSaveChanges)

    messagebox.showinfo("Success", "Photoshop automation completed successfully.")

if __name__ == "__main__":
    automate_photoshop()
