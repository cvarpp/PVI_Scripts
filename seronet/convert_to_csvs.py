import pandas as pd
import sys
import glob
import os
import argparse
import customtkinter as tk

class Converter(tk.CTk):
    def __init__(self, fg_color = None, **kwargs):

        # Setup
        super().__init__(fg_color, **kwargs)
        self.minsize(700, 300)
        ##self.geometry(##HERE)
        self.resizable(True, True)
        self.title("CSV Converter")
        tk.set_appearance_mode("System")
        tk.set_default_color_theme("blue")
        
        # Instance Variables (Specifically StringVars for updating labels easily)
        self.infilevar = tk.StringVar(master=self, value="None Selected")
        self.outfilevar = tk.StringVar(master=self, value="None Selected")
        self.statusvar = tk.StringVar(master=self, value="Please select file paths, then press convert.")

        # Widgets
        self.ask_in_button = tk.CTkButton(master=self, text="Select Input", command = self.get_infilename)
        self.ask_out_button = tk.CTkButton(master=self, text="Select Output", command = self.get_outfilename)
        self.clear_button = tk.CTkButton(master=self, text="Clear selection", command = self.clear_inputs, fg_color = "#753A36", text_color = "white")
        self.convert_button = tk.CTkButton(master=self, text="Convert!", command = self.do_conversion)
        self.indir_label = tk.CTkLabel(master=self, textvariable = self.infilevar)
        self.outdir_label = tk.CTkLabel(master=self, textvariable = self.outfilevar)
        self.status_label = tk.CTkLabel(master=self, textvariable = self.statusvar)
        self.indir_label_label = tk.CTkLabel(master=self, text="Selected Input Folder:")
        self.outdir_label_label = tk.CTkLabel(master=self, text="Selected Output Folder:")
        self.status_label_label = tk.CTkLabel(master=self, text="Status:")

        # Layout
        self.ask_in_button.grid(row=0, column=0, padx=20, pady=20, sticky="w")
        self.ask_out_button.grid(row=1, column=0, padx=20, pady=20, sticky="w")
        self.clear_button.grid(row=2, column=0, padx=20, pady=20)
        self.convert_button.grid(row=2, column=2, padx=20, pady=20)
        self.indir_label.grid(row=0, column=2, padx=20, pady=20, sticky="ew")
        self.outdir_label.grid(row=1, column=2, padx=20, pady=20, sticky="ew")
        self.status_label.grid(row=4, column=2, padx=20, pady=20, sticky="ew")
        self.indir_label_label.grid(row=0, column=1, padx=20, pady=20, sticky="w")
        self.outdir_label_label.grid(row=1, column=1, padx=20, pady=20, sticky="w")
        self.status_label_label.grid(row=4, column=1, padx=20, pady=20, sticky="e")

    # Methods
    def get_infilename(self):
        self.infilevar.set(tk.filedialog.askdirectory())

    def get_outfilename(self):
        self.outfilevar.set(tk.filedialog.askdirectory())

    def clear_inputs(self):
        self.infilevar.set("None Selected")
        self.outfilevar.set("None Selected")

    def do_conversion(self):
        counter = 0
        for fname in glob.glob('{}/*xls*'.format(self.infilevar.get())):
            if '~' in fname:
                continue
            pd.read_excel(fname, na_filter=False, keep_default_na=False).to_csv('{}/{}.csv'.format(self.outfilevar.get(), fname.split(os.sep)[-1].split('.')[0]), index=False)
            counter += 1
            self.statusvar.set("Successfully converted " + str(counter) + " files...")
        self.statusvar.set("Done! Converted " + str(counter) + " files.")



if __name__ == '__main__':
    if len(sys.argv) != 1:
        argParser = argparse.ArgumentParser(description='Read Excel files in input_dir and create corresponding csv files in output_dir')
        argParser.add_argument('-i', '--input_dir', action='store', required=True)
        argParser.add_argument('-o', '--output_dir', action='store', required=True)
        args = argParser.parse_args()
        for fname in glob.glob('{}/*xls*'.format(args.input_dir)):
            if '~' in fname:
                continue
            print("Converting", fname)
            pd.read_excel(fname, na_filter=False, keep_default_na=False).to_csv('{}/{}.csv'.format(args.output_dir, fname.split(os.sep)[-1].split('.')[0]), index=False)
            print(fname, "converted!")
    else:
        app = Converter()
        app.attributes('-topmost', True)
        app.lift()
        app.attributes('-topmost', False)
        app.mainloop()
