import numpy as np
import pandas as pd
import argparse
import util

def clean_strings(s):
    return s.astype(str).str.upper().str.strip()

if __name__ == '__main__':
    argParser = argparse.ArgumentParser(description='Find next pal racks.')
    args = argParser.parse_args()
    df = pd.read_excel(util.psp + 'NPS Rack Status.xlsx').dropna(subset=['Rack number']).set_index('Rack number')
    new_idx = sorted([str(idx).strip().upper() for idx in df.index.to_numpy()])
    labels = ['NPS {}'.format(n) for n in range(1,1000)]
    new_idx += [label for label in labels if label not in new_idx]
    df = df.reindex(index=new_idx).reset_index()
    next_unlinked = df[df['Linked?'].pipe(clean_strings) != 'YES'].head(6)['Rack number'].to_numpy()
    next_unprinted = df[df['Printed?'].pipe(clean_strings) != 'YES'].head(6)['Rack number'].to_numpy()
    next_unaliquoted = df[df['Aliquoted?'].pipe(clean_strings) != 'YES'].head(6)['Rack number'].to_numpy()
    all_todo = np.concatenate((next_unlinked, next_unprinted, next_unaliquoted))
    out_df = pd.DataFrame({'Rack': all_todo}).drop_duplicates().set_index('Rack')
    for col in ['Linking Initials', 'Linking Weekday',
                'Printing Initials', 'Printing Weekday',
                'Aliquoting Initials', 'Aliquoting Weekday']:
        out_df[col] = 'TBD'
    for rname, row in out_df.iterrows():
        for col in ['Linking Initials', 'Linking Weekday']:
            if rname not in next_unlinked:
                row.loc[col] = 'N/A'
        for col in ['Printing Initials', 'Printing Weekday']:
            if rname not in next_unprinted:
                row.loc[col] = 'N/A'
        for col in ['Aliquoting Initials', 'Aliquoting Weekday']:
            if rname not in next_unaliquoted:
                row.loc[col] = 'N/A'
    output_fname = util.psp + 'PAL Plan/PAL {}.xlsx'.format(pd.to_datetime('today').date())
    out_df = out_df.style.apply(
        lambda x: ["background-color: orange" if v == 'TBD'
                   else "background-color: black" for v in x], axis=1)
    with pd.ExcelWriter(output_fname) as writer:
        out_df.to_excel(writer, na_rep='N/A')
