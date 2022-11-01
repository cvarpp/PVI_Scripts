import pandas as pd
import sys
import glob
import os

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("usage: python convert_to_csvs.py input_directory output_directory")
        exit(1)
    input_dirname = sys.argv[1]
    output_dirname = sys.argv[2]
    for fname in glob.glob('{}/*xls*'.format(input_dirname)):
        print("Converting", fname)
        pd.read_excel(fname, na_filter=False, keep_default_na=False).to_csv('{}/{}.csv'.format(output_dirname, fname.split(os.sep)[-1][:-5]), index=False)
        print(fname, "converted!")
    print("Done!")
