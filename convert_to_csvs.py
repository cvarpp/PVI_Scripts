import pandas as pd
import sys
import glob
import os
import argparse

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Read Excel files in input_dir and create corresponding csv files in output_dir')
    argParser.add_argument('-i', '--input_dir', action='store', required=True)
    argParser.add_argument('-o', '--output_dir', action='store', required=True)
    args = argParser.parse_args()
    for fname in glob.glob('{}/*xls*'.format(args.input_dir)):
        print("Converting", fname)
        pd.read_excel(fname, na_filter=False, keep_default_na=False).to_csv('{}/{}.csv'.format(args.output_dir, fname.split(os.sep)[-1].split('.')[0]), index=False)
        print(fname, "converted!")
    print("Done!")
