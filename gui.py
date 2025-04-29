import tkinter as tk
from tkinter import filedialog
import json
from oft_parser import extract_html_from_oft  # Import the function

class OFTConverterGUI:
    def __init__(self, master):
        self.master = master
        master.title("OFT to HTML Converter")

        self.font_styles = self.load_font_styles("font_styles.json")

        self.select_oft_button = tk.Button(master, text="Select OFT File", command=self.select_oft_file)
        self.select_oft_button.pack()

        self.convert_button = tk.Button(master, text="Convert", command=self.convert_oft, state=tk.DISABLED)  # Initially disabled
        self.convert_button.pack()

        self.output_text = tk.Text(master, height=20, width=80)  # Increased size
        self.output_text.pack()

        self.oft_file_path = None

    def load_font_styles(self, filepath):
        try:
            with open(filepath, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Error: Font style file not found at {filepath}")
            return {}  # Return an empty dictionary if the file is not found
        except json.JSONDecodeError:
            print(f"Error: Invalid JSON format in {filepath}")
            return {} # Return an empty dictionary if the JSON is invalid

    def select_oft_file(self):
        self.oft_file_path = filedialog.askopenfilename(filetypes=[("Outlook Template Files", "*.oft")])
        if self.oft_file_path:
            print(f"Selected OFT file: {self.oft_file_path}")
            self.convert_button["state"] = tk.NORMAL  # Enable the convert button
        else:
            self.convert_button["state"] = tk.DISABLED # Disable the convert button

    def convert_oft(self):
        self.output_text.delete("1.0", tk.END)  # Clear previous output

        if self.oft_file_path:
            html_content = extract_html_from_oft(self.oft_file_path)
            if html_content:
                self.output_text.insert(tk.END, "HTML Content:\n")
                self.output_text.insert(tk.END, html_content)
            else:
                self.output_text.insert(tk.END, "Failed to extract HTML from OFT file.\n")
        else:
            self.output_text.insert(tk.END, "No OFT file selected.\n")

        self.output_text.insert(tk.END, f"\nFont Styles: {self.font_styles}\n")
        self.output_text.insert(tk.END, f"OFT File Path: {self.oft_file_path}\n")

if __name__ == '__main__':
    root = tk.Tk()
    gui = OFTConverterGUI(root)
    root.mainloop()