import pandas as pd
import datetime
import argparse
import util
import os


if __name__ == '__main__':
    templates = util.cross_d4 + 'Template Files/'
    _dirpath, _dirnames, valid_fnames = next(os.walk(templates))
    dfs_to_concat = {fname: [] for fname in valid_fnames}
    for dirpath, _dirnames, fnames in os.walk(util.cross_d4 + 'Data/Accepted Submissions/'):
        for fname in fnames:
            if fname in valid_fnames:
                dfs_to_concat[fname].append(pd.read_excel(os.path.join(dirpath, fname), keep_default_na=False))
    for fname, dfs in dfs_to_concat.items():
        if len(dfs) > 0:
            df = pd.concat(dfs)
            df.to_excel(util.cross_d4 + 'Local Database/' + fname, index=False,  na_rep='N/A')
