import tkinter as tk
from gui import OFTConverterGUI

def main():
    root = tk.Tk()
    gui = OFTConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()