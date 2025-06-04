# gui.py

import tkinter as tk
from tkinter import filedialog, messagebox
from processor import InterferenceProcessor, DEFAULT_ZONE_ORDER

class App(tk.Frame):
    """
    Interfaccia grafica che permette di:
    - selezionare il file Excel
    - inserire il nome del foglio
    - definire l'ordine delle zone (inserire nomi arbitrari e riordinare)
    - selezionare i percorsi di output per chart_config.txt e interferences_summary.txt
    - avviare la generazione dei file
    """

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Interference Generator")
        self.grid(padx=10, pady=10)
        self.create_widgets()

    def create_widgets(self):
        # Excel file
        tk.Label(self, text="Excel file:").grid(row=0, column=0, sticky="e")
        self.excel_entry = tk.Entry(self, width=50)
        self.excel_entry.grid(row=0, column=1, padx=5, pady=2)
        tk.Button(self, text="Browse…", command=self.browse_excel).grid(row=0, column=2, padx=5)

        # Sheet name
        tk.Label(self, text="Sheet name:").grid(row=1, column=0, sticky="e")
        self.sheet_entry = tk.Entry(self, width=30)
        self.sheet_entry.insert(0, "Variabili")
        self.sheet_entry.grid(row=1, column=1, padx=5, pady=2, sticky="w")

        # Zone order: Entry + Add + Listbox + Up/Down/Remove
        tk.Label(self, text="Zone order:").grid(row=2, column=0, sticky="ne")
        self.zone_entry = tk.Entry(self, width=20)
        self.zone_entry.grid(row=2, column=1, sticky="w", padx=5, pady=2)

        tk.Button(self, text="Add Zone", command=self.add_zone).grid(row=2, column=2, padx=5)

        self.zone_listbox = tk.Listbox(self, height=8, width=25)
        self.zone_listbox.grid(row=3, column=1, sticky="w", padx=5)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=self.zone_listbox.yview)
        scrollbar.grid(row=3, column=1, sticky="nse", padx=(0,5))
        self.zone_listbox.config(yscrollcommand=scrollbar.set)

        # Up / Down / Remove buttons
        btn_frame = tk.Frame(self)
        btn_frame.grid(row=3, column=2, sticky="n", padx=5)
        tk.Button(btn_frame, text="Up", width=8, command=self.move_up).pack(pady=(0,5))
        tk.Button(btn_frame, text="Down", width=8, command=self.move_down).pack(pady=(0,5))
        tk.Button(btn_frame, text="Remove", width=8, command=self.remove_zone).pack()

        # Prepopola listbox con DEFAULT_ZONE_ORDER
        for z in DEFAULT_ZONE_ORDER:
            self.zone_listbox.insert("end", z)

        # ChartConfig output
        tk.Label(self, text="ChartConfig output:").grid(row=4, column=0, sticky="e")
        self.chart_entry = tk.Entry(self, width=50)
        self.chart_entry.insert(0, "chart_config.txt")
        self.chart_entry.grid(row=4, column=1, padx=5, pady=2)
        tk.Button(self, text="Browse…", command=self.browse_chart).grid(row=4, column=2, padx=5)

        # Summary output
        tk.Label(self, text="Summary output:").grid(row=5, column=0, sticky="e")
        self.summary_entry = tk.Entry(self, width=50)
        self.summary_entry.insert(0, "interferences_summary.txt")
        self.summary_entry.grid(row=5, column=1, padx=5, pady=2)
        tk.Button(self, text="Browse…", command=self.browse_summary).grid(row=5, column=2, padx=5)

        # Pulsante "Generate"
        self.generate_btn = tk.Button(self, text="Generate Files", command=self.generate_files)
        self.generate_btn.grid(row=6, column=0, columnspan=3, pady=10)

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)

    def browse_chart(self):
        path = filedialog.asksaveasfilename(
            title="Save chart_config as…",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")]
        )
        if path:
            self.chart_entry.delete(0, tk.END)
            self.chart_entry.insert(0, path)

    def browse_summary(self):
        path = filedialog.asksaveasfilename(
            title="Save summary as…",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")]
        )
        if path:
            self.summary_entry.delete(0, tk.END)
            self.summary_entry.insert(0, path)

    def add_zone(self):
        """Aggiunge il testo di zone_entry nella listbox, se non vuoto e non duplicato."""
        zone = self.zone_entry.get().strip()
        if not zone:
            return
        existing = self.zone_listbox.get(0, "end")
        if zone in existing:
            messagebox.showwarning("Warning", f"Zona '{zone}' già presente.")
        else:
            self.zone_listbox.insert("end", zone)
        self.zone_entry.delete(0, tk.END)

    def move_up(self):
        """Sposta la voce selezionata su di un posto."""
        sel = self.zone_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == 0:
            return
        text = self.zone_listbox.get(idx)
        self.zone_listbox.delete(idx)
        self.zone_listbox.insert(idx - 1, text)
        self.zone_listbox.selection_set(idx - 1)

    def move_down(self):
        """Sposta la voce selezionata giù di un posto."""
        sel = self.zone_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        last = self.zone_listbox.size() - 1
        if idx == last:
            return
        text = self.zone_listbox.get(idx)
        self.zone_listbox.delete(idx)
        self.zone_listbox.insert(idx + 1, text)
        self.zone_listbox.selection_set(idx + 1)

    def remove_zone(self):
        """Rimuove la voce selezionata dalla listbox."""
        sel = self.zone_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.zone_listbox.delete(idx)

    def get_zone_order(self):
        """Restituisce la lista di zone dall’alto in basso nella listbox."""
        return list(self.zone_listbox.get(0, "end"))

    def generate_files(self):
        excel_path = self.excel_entry.get().strip()
        sheet_name = self.sheet_entry.get().strip()
        chart_out = self.chart_entry.get().strip()
        summary_out = self.summary_entry.get().strip()
        zone_order = self.get_zone_order()

        if not excel_path or not sheet_name or not chart_out or not summary_out:
            messagebox.showerror("Error", "Tutti i campi devono essere compilati.")
            return

        if not zone_order:
            messagebox.showerror("Error", "Devi definire almeno una zona.")
            return

        try:
            processor = InterferenceProcessor(
                excel_path=excel_path,
                sheet_name=sheet_name,
                output_chart=chart_out,
                output_summary=summary_out,
                zone_order=zone_order,
            )
            processor.run()
            messagebox.showinfo("Success", "File generati correttamente.")
        except Exception as e:
            messagebox.showerror("Generation Error", str(e))
