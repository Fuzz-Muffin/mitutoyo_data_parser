import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import pandas as pd


def gendic(label=None, row=-1, col=-1, x=0.0, y=0.0, z=0.0,
           x_nom=0.0, y_nom=0.0, z_nom=0.0, dx=0.0, dy=0.0, dz=0.0,
           rad=0.0, dia=0.0, drad=0.0, ddia=0.0, rad_nom=0.0, dia_nom=0.0):
    return {
        'label': label,
        'row': row,
        'col': col,
        'x': x,
        'y': y,
        'z': z,
        'x_nom': x_nom,
        'y_nom': y_nom,
        'z_nom': z_nom,
        'dx': dx,
        'dy': dy,
        'dz': dz,
        'radius': rad,
        'diameter': dia,
        'radius_nom': rad_nom,
        'diameter_nom': dia_nom,
        'd_radius' : drad,
        'd_diameter': ddia,
    }


def parse_files(targets_fpath, data_fpath, log_func=print):
    log_func("Reading targets file...")
    possible_objects = ['Circle',]
    targets_list = []
    if targets_fpath:
        log_func("Reading provided target file to build list of desired objects...")
        with open(targets_fpath, 'r', encoding='utf-8') as fi:
            for line in fi:
                blah = line.strip()
                if blah:
                    targets_list.append(blah)
    else:
        log_func("No target file provided, will scan data file and capture all objects...")
        with open(data_fpath, 'r', encoding='utf-8') as fi:
            for i,line in enumerate(fi.readlines()):   
                if i>0: # skip header
                    if any([x in line for x in possible_objects]):
                        tmp = line.strip().split(':')[1].split('[')[0].strip()
                        log_func(f"found object {tmp}")
                        targets_list.append(tmp)
        targets_list = list(dict.fromkeys(targets_list))

    log_func(f"Loaded {len(targets_list)} targets.")
    log_func("Reading data file...")
    extracted_data = []

    with open(data_fpath, 'r', encoding='utf-8') as fi:
        lines = fi.readlines()

    for i, line in enumerate(lines):
        which_target = [target in line for target in targets_list]
        if any(which_target):
            lab = targets_list[[j for j, x in enumerate(which_target) if x][0]]
            log_func(f"Found relevant data for {lab} at line {i+1}")

            actual_pos = [0.0, 0.0, 0.0]
            nominal_pos = [0.0, 0.0, 0.0]
            deviation = [0.0, 0.0, 0.0]
            radius_actual= 0.0
            radius_nominal= 0.0
            radius_deviation = 0.0            
            diam_actual= 0.0
            diam_nominal= 0.0
            diam_deviation = 0.0

            try:
                tmp = line.split('[', 1)[1].split(']', 1)[0]
                if ',' in tmp:
                    tmp = tmp.split(',', 1)
                    row = int(tmp[0].strip())
                    col = int(tmp[1].strip())
                else:
                    row = int(tmp.strip())
                    col = -1
            except Exception:
                row = -1
                col = -1

            for next_line in lines[i + 1:]:
                if next_line.strip() == "":
                    break

                data_in_list = next_line.split('=')[1].strip().split()
                # find out which coord we are on
                idx = -1
                if 'Coord. X' in next_line:
                    idx=0
                elif 'Coord. Y' in next_line:
                    idx=1
                elif 'Coord. Z' in next_line:
                    idx=2
                elif 'Radius' in next_line:
                    idx=3
                elif 'Diam' in next_line:
                    idx=4

                if idx > -1:
                    if idx <=2:
                        # actual position should always be there if the coordinate is present
                        actual_pos[idx]  = float(data_in_list[0])
                        try:
                            nominal_pos[idx] = float(data_in_list[1])
                            deviation[idx] = float(data_in_list[2])
                        except Exception:
                            pass
                    elif idx == 3:
                        radius_actual  = float(data_in_list[0])
                        try:
                            radius_nominal = float(data_in_list[1])
                            radius_deviation = float(data_in_list[2])
                        except Exception:
                            pass
                    elif idx == 4:
                        diam  = float(data_in_list[0])
                        try:
                            diam_nominal = float(data_in_list[1])
                            diam_deviation = float(data_in_list[2])
                        except Exception:
                            pass

            extracted_data.append(gendic(label=lab,
                                         row=row,
                                         col=col, 
                                         x=actual_pos[0],
                                         y=actual_pos[1],
                                         z=actual_pos[2],
                                         x_nom=nominal_pos[0],
                                         y_nom=nominal_pos[1],
                                         z_nom=nominal_pos[2],
                                         dx= deviation[0],
                                         dy= deviation[1],
                                         dz= deviation[2],
                                         rad= radius_actual,
                                         rad_nom= radius_nominal,
                                         drad= radius_deviation,
                                         dia= diam_actual,
                                         dia_nom= diam_nominal,
                                         ddiam= diam_deviation,
                                         ))

    df = pd.DataFrame(extracted_data)
    log_func(f"Parsing complete. Extracted {len(df)} rows.")
    return df


def save_formatted_excel(df, output_path):
    """
    Saves an additional Excel file with some basic formatting.
    """
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ParsedData")
        ws = writer.book["ParsedData"]

        # Freeze top row
        ws.freeze_panes = "A2"

        # Auto-filter
        ws.auto_filter.ref = ws.dimensions

        # Set widths
        widths = {
            "A": 18, "B": 10, "C": 10, "D": 12, "E": 12, "F": 12,
            "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12
        }
        for col_letter, width in widths.items():
            ws.column_dimensions[col_letter].width = width


class ParserGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Parser GUI")
        self.geometry("1200x700")
        self.minsize(1100, 650)

        self.targets_path = tk.StringVar()
        self.data_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.save_formatted_copy = tk.BooleanVar(value=True)

        self.df = None

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Label(top, text="targets.txt:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(top, textvariable=self.targets_path, width=90).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(top, text="Browse", command=self.browse_targets).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(top, text="test_data.txt:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(top, textvariable=self.data_path, width=90).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(top, text="Browse", command=self.browse_data).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(top, text="Output file:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(top, textvariable=self.output_path, width=90).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(top, text="Browse", command=self.browse_output).grid(row=2, column=2, padx=5, pady=5)

        ttk.Checkbutton(
            top,
            text="Also save formatted Excel copy",
            variable=self.save_formatted_copy
        ).grid(row=3, column=1, sticky="w", padx=5, pady=5)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(button_frame, text="Run", command=self.run).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side="left", padx=5)

        # Main area with preview and log
        main = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # Dataframe preview
        preview_frame = ttk.Labelframe(main, text="DataFrame Preview")
        main.add(preview_frame, weight=3)

        self.tree = ttk.Treeview(preview_frame, show="headings")
        self.tree.pack(fill="both", expand=True, side="left")

        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        # Log window
        log_frame = ttk.Labelframe(main, text="Log")
        main.add(log_frame, weight=2)

        self.log_text = tk.Text(log_frame, wrap="word", height=20)
        self.log_text.pack(fill="both", expand=True, side="left")

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.update_idletasks()

    def clear_log(self):
        self.log_text.delete("1.0", "end")

    def browse_targets(self):
        path = filedialog.askopenfilename(
            title="Select targets.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if path:
            self.targets_path.set(path)

    def browse_data(self):
        path = filedialog.askopenfilename(
            title="Select test_data.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if path:
            self.data_path.set(path)

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save output file as",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.output_path.set(path)

    def preview_dataframe(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row.astype(str)))

    def run(self):
        targets_path = self.targets_path.get().strip()
        data_path = self.data_path.get().strip()
        output_path = self.output_path.get().strip()

        if not data_path or not output_path:
            messagebox.showerror("Missing input", "Please select data in and output paths.")
            return

        try:
            self.log("Starting parse...")
            self.log(f"targets path: {targets_path}")
            df = parse_files(targets_path, data_path, log_func=self.log)
            self.df = df

            self.preview_dataframe(df)

            out_ext = Path(output_path).suffix.lower()
            if out_ext == ".xlsx":
                self.log(f"Saving main output to Excel: {output_path}")
                df.to_excel(output_path, index=False)
            else:
                self.log(f"Saving main output to CSV: {output_path}")
                df.to_csv(output_path, index=False)

            if self.save_formatted_copy.get():
                formatted_path = str(Path(output_path).with_name(Path(output_path).stem + "_formatted.xlsx"))
                self.log(f"Saving formatted Excel copy: {formatted_path}")
                save_formatted_excel(df, formatted_path)

            self.log("Done.")
            messagebox.showinfo("Success", "Parsing complete and output saved.")

        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", f"An error occurred:\n{e}")


if __name__ == "__main__":
    app = ParserGUI()
    app.mainloop()