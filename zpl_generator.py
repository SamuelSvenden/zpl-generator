import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import serial
import serial.tools.list_ports
import socket
import json
import os
import requests
import base64
from PIL import Image, ImageTk
from io import BytesIO
import webbrowser
import urllib.parse
import win32print

class ZPLGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("ZPL Barcode Generator")
        self.root.geometry("900x800")  # Reduced width from 1200 to 900
        
        # Create main canvas with scrollbar
        self.main_canvas = tk.Canvas(root)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.main_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.main_canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        )
        
        # Bind mouse wheel events
        self.main_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.main_canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.main_canvas.bind_all("<Button-5>", self._on_mousewheel)
        
        self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack the canvas and scrollbar
        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Create main frame
        self.main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Barcode Parameters
        self.create_barcode_section()
        
        # Label Parameters
        self.create_label_section()
        
        # Printer Settings
        self.create_printer_section()
        
        # Configuration Management
        self.create_config_section()
        
        # Generate Button
        self.create_generate_button()
        
        # Preview Area
        self.create_preview_section()
        
        # Image Preview
        self.create_image_preview_section()
        
        # Load defaults if they exist
        self.load_defaults()

    def _on_mousewheel(self, event):
        if event.num == 5 or event.delta < 0:  # scroll down
            self.main_canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:  # scroll up
            self.main_canvas.yview_scroll(-1, "units")

    def create_barcode_section(self):
        barcode_frame = ttk.LabelFrame(self.main_frame, text="Barcode Parameters", padding="5")
        barcode_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Start Number
        ttk.Label(barcode_frame, text="Start Number:").grid(row=0, column=0, sticky=tk.W)
        self.start_number = ttk.Entry(barcode_frame, width=10)
        self.start_number.insert(0, "4300012")
        self.start_number.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Barcode Height
        ttk.Label(barcode_frame, text="Height:").grid(row=1, column=0, sticky=tk.W)
        self.barcode_height = ttk.Entry(barcode_frame, width=10)
        self.barcode_height.insert(0, "30")
        self.barcode_height.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Module Width
        ttk.Label(barcode_frame, text="Module Width:").grid(row=2, column=0, sticky=tk.W)
        self.module_width = ttk.Entry(barcode_frame, width=10)
        self.module_width.insert(0, "4")
        self.module_width.grid(row=2, column=1, sticky=tk.W, padx=5)

    def create_label_section(self):
        label_frame = ttk.LabelFrame(self.main_frame, text="Label Parameters", padding="5")
        label_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Number of Labels
        ttk.Label(label_frame, text="Number of Labels:").grid(row=0, column=0, sticky=tk.W)
        self.num_labels = ttk.Entry(label_frame, width=10)
        self.num_labels.insert(0, "10")
        self.num_labels.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # X Position
        ttk.Label(label_frame, text="X Position:").grid(row=1, column=0, sticky=tk.W)
        self.x_pos = ttk.Entry(label_frame, width=10)
        self.x_pos.insert(0, "20")
        self.x_pos.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Y Position
        ttk.Label(label_frame, text="Y Position:").grid(row=2, column=0, sticky=tk.W)
        self.y_pos = ttk.Entry(label_frame, width=10)
        self.y_pos.insert(0, "20")
        self.y_pos.grid(row=2, column=1, sticky=tk.W, padx=5)

    def create_printer_section(self):
        printer_frame = ttk.LabelFrame(self.main_frame, text="Printer Settings", padding="5")
        printer_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Get installed printers
        printers = ["Preview"]
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            printers.append(printer[2])
        
        # Printer Selection
        ttk.Label(printer_frame, text="Printer:").grid(row=0, column=0, sticky=tk.W)
        self.printer = ttk.Combobox(printer_frame, values=printers, width=25)  # Reduced width from 30 to 25
        self.printer.set("Preview")
        self.printer.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Refresh Button
        ttk.Button(printer_frame, text="Refresh Printers", command=self.refresh_printers).grid(row=0, column=2, padx=5)

    def create_config_section(self):
        config_frame = ttk.LabelFrame(self.main_frame, text="Configuration Management", padding="5")
        config_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Save Configuration Button
        ttk.Button(config_frame, text="Save Configuration", command=self.save_config).grid(row=0, column=0, padx=5)
        
        # Load Configuration Button
        ttk.Button(config_frame, text="Load Configuration", command=self.load_config).grid(row=0, column=1, padx=5)
        
        # Set Defaults Button
        ttk.Button(config_frame, text="Set Defaults", command=self.save_defaults).grid(row=0, column=2, padx=5)

    def create_generate_button(self):
        self.generate_btn = ttk.Button(self.main_frame, text="Generate ZPL", command=self.generate_zpl)
        self.generate_btn.grid(row=4, column=0, columnspan=2, pady=10)

    def create_preview_section(self):
        preview_frame = ttk.LabelFrame(self.main_frame, text="ZPL Preview", padding="5")
        preview_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=5)
        
        self.preview_text = tk.Text(preview_frame, height=20, width=40)  # Reduced width from 50 to 40
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Open in Labelary Button
        ttk.Button(preview_frame, text="Open in Labelary", command=self.open_in_labelary).grid(row=1, column=0, pady=5)

    def create_image_preview_section(self):
        image_frame = ttk.LabelFrame(self.main_frame, text="Label Preview", padding="5")
        image_frame.grid(row=5, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=5)
        
        self.image_label = ttk.Label(image_frame)
        self.image_label.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Save Image Button
        ttk.Button(image_frame, text="Save Preview Image", command=self.save_preview_image).grid(row=1, column=0, pady=5)

    def save_defaults(self):
        try:
            defaults = {
                'barcode': {
                    'start_number': self.start_number.get(),
                    'height': self.barcode_height.get(),
                    'module_width': self.module_width.get()
                },
                'label': {
                    'num_labels': self.num_labels.get(),
                    'x_pos': self.x_pos.get(),
                    'y_pos': self.y_pos.get()
                },
                'printer': {
                    'name': self.printer.get()
                }
            }
            
            # Save to defaults.json in the same directory as the script
            with open('defaults.json', 'w') as f:
                json.dump(defaults, f, indent=4)
            messagebox.showinfo("Success", "Default values saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save default values: {str(e)}")

    def load_defaults(self):
        try:
            if os.path.exists('defaults.json'):
                with open('defaults.json', 'r') as f:
                    defaults = json.load(f)
                
                # Load barcode settings
                self.start_number.delete(0, tk.END)
                self.start_number.insert(0, defaults['barcode']['start_number'])
                self.barcode_height.delete(0, tk.END)
                self.barcode_height.insert(0, defaults['barcode']['height'])
                self.module_width.delete(0, tk.END)
                self.module_width.insert(0, defaults['barcode']['module_width'])
                
                # Load label settings
                self.num_labels.delete(0, tk.END)
                self.num_labels.insert(0, defaults['label']['num_labels'])
                self.x_pos.delete(0, tk.END)
                self.x_pos.insert(0, defaults['label']['x_pos'])
                self.y_pos.delete(0, tk.END)
                self.y_pos.insert(0, defaults['label']['y_pos'])
                
                # Load printer settings
                self.printer.set(defaults['printer']['name'])
        except Exception as e:
            print(f"Failed to load defaults: {str(e)}")

    def refresh_printers(self):
        printers = ["Preview"]
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            printers.append(printer[2])
        self.printer['values'] = printers

    def open_in_labelary(self):
        try:
            # Get the current ZPL code
            zpl = self.preview_text.get(1.0, tk.END).strip()
            if not zpl:
                messagebox.showwarning("Warning", "No ZPL code to open in Labelary")
                return
            
            # Encode the ZPL code for URL
            encoded_zpl = urllib.parse.quote(zpl)
            
            # Open in default browser
            url = f"http://labelary.com/viewer.html?density=8&width=4&height=6&zpl={encoded_zpl}"
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open in Labelary: {str(e)}")

    def save_config(self):
        config = {
            'barcode': {
                'start_number': self.start_number.get(),
                'height': self.barcode_height.get(),
                'module_width': self.module_width.get()
            },
            'label': {
                'num_labels': self.num_labels.get(),
                'x_pos': self.x_pos.get(),
                'y_pos': self.y_pos.get()
            },
            'printer': {
                'name': self.printer.get()
            }
        }
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    json.dump(config, f, indent=4)
                messagebox.showinfo("Success", "Configuration saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")

    def load_config(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    config = json.load(f)
                
                # Load barcode settings
                self.start_number.delete(0, tk.END)
                self.start_number.insert(0, config['barcode']['start_number'])
                self.barcode_height.delete(0, tk.END)
                self.barcode_height.insert(0, config['barcode']['height'])
                self.module_width.delete(0, tk.END)
                self.module_width.insert(0, config['barcode']['module_width'])
                
                # Load label settings
                self.num_labels.delete(0, tk.END)
                self.num_labels.insert(0, config['label']['num_labels'])
                self.x_pos.delete(0, tk.END)
                self.x_pos.insert(0, config['label']['x_pos'])
                self.y_pos.delete(0, tk.END)
                self.y_pos.insert(0, config['label']['y_pos'])
                
                # Load printer settings
                self.printer.set(config['printer']['name'])
                
                messagebox.showinfo("Success", "Configuration loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")

    def get_labelary_preview(self, zpl):
        try:
            # Labelary API endpoint
            url = "http://api.labelary.com/v1/printers/8dpmm/labels/4x6/0/"
            
            # Send ZPL to Labelary
            response = requests.post(url, data=zpl.encode())
            
            if response.status_code == 200:
                # Convert response to image
                image = Image.open(BytesIO(response.content))
                return image
            else:
                messagebox.showerror("Error", f"Failed to get preview: {response.text}")
                return None
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get preview: {str(e)}")
            return None

    def save_preview_image(self):
        if hasattr(self, 'current_preview_image'):
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
            )
            
            if file_path:
                try:
                    self.current_preview_image.save(file_path)
                    messagebox.showinfo("Success", "Preview image saved successfully!")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save image: {str(e)}")

    def generate_zpl(self):
        try:
            # Get values from entries
            start_num = self.start_number.get()
            height = self.barcode_height.get()
            module_width = self.module_width.get()
            num_labels = self.num_labels.get()
            x_pos = self.x_pos.get()
            y_pos = self.y_pos.get()
            
            # Generate ZPL code
            zpl = f"""^XA
^LH5,5
^FO{x_pos},{y_pos}
^BY1,{module_width}
^BCN,{height},Y,N,N
^SN{start_num},1,Y
^PQ{num_labels}
^XZ"""
            
            # Update preview
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, zpl)
            
            # Get preview image
            preview_image = self.get_labelary_preview(zpl)
            if preview_image:
                # Resize image to fit in the window
                preview_image.thumbnail((400, 400))
                self.current_preview_image = preview_image
                
                # Convert to PhotoImage
                photo = ImageTk.PhotoImage(preview_image)
                self.image_label.configure(image=photo)
                self.image_label.image = photo  # Keep a reference
            
            # Send to printer if not in preview mode
            if self.printer.get() != "Preview":
                try:
                    # Get the printer handle
                    printer_handle = win32print.OpenPrinter(self.printer.get())
                    try:
                        # Start a print job
                        job = win32print.StartDocPrinter(printer_handle, 1, ("ZPL Label", None, "RAW"))
                        try:
                            # Start a page
                            win32print.StartPagePrinter(printer_handle)
                            # Write the ZPL data
                            win32print.WritePrinter(printer_handle, zpl.encode())
                            # End the page
                            win32print.EndPagePrinter(printer_handle)
                        finally:
                            # End the print job
                            win32print.EndDocPrinter(printer_handle)
                    finally:
                        # Close the printer
                        win32print.ClosePrinter(printer_handle)
                    messagebox.showinfo("Success", f"ZPL code sent to {self.printer.get()} successfully!")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send to printer: {str(e)}")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ZPLGenerator(root)
    root.mainloop()
