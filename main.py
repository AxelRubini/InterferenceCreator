# main.py

import tkinter as tk
from gui import App

def main():
    root = tk.Tk()
    app = App(master=root)
    app.mainloop()

if __name__ == "__main__":
    main()
