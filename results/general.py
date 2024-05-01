import pandas as pd
from datetime import date
from dateutil import parser
import argparse
import util
from helpers import query_intake

if __name__ == '__main__':
    argparser = argparse.ArgumentParser(description='Generate report based on participant IDs in results_of_interest. Legacy script')
    argparser.add_argument('-i', '--input_file', action='store', default=util.script_input + 'results_of_interest.xlsx', help="Path to input file (defaults to script input)")
    argparser.add_argument('-o', '--output_file', action='store', default='tmp', help="Prefix for the output file (in addition to current date)")
    argparser.add_argument('-d', '--debug', action='store_true', help="Print to the command line but do not write to file")
    args = argparser.parse_args()

    print("List of participant ids used from {}".format(args.input_file))
    if 'xls' in args.input_file:
        pid_df = pd.read_excel(args.input_file)
    elif 'csv' in args.input_file:
        pid_df = pd.read_csv(args.input_file)
    else:
        print("Invalid file type. Fatal error, exiting...")
        exit(1)
    participants = [str(x).strip().upper() for x in pid_df['Participant ID']]
    print("Pulling data for {} participants".format(len(participants)))
    samplesClean = query_intake(participants=participants, include_research=True).rename(
        columns={'Date Collected': 'Date',
                 "Visit Type / Samples Needed": 'Visit Type'})
    baselines = samplesClean.drop_duplicates(subset='participant_id').set_index('participant_id')
    samplesClean['Post-Baseline'] = samplesClean.apply(lambda row: int((row['Date'] - baselines.loc[row['participant_id'], 'Date']).days), axis=1)
    keep_cols = ['participant_id', 'Date', 'Post-Baseline', 'sample_id', 'Visit Type', 'Qualitative', 'Quantitative', 'COV22', 'Spike endpoint', 'AUC']
    report = samplesClean.reset_index().loc[:, keep_cols]
    output_filename = util.script_output + 'results_{}_{}.xlsx'.format(args.output_file, date.today().strftime("%m.%d.%y"))
    if not args.debug:
        report.to_excel(output_filename, index=False)
        print("Report written to {}".format(output_filename))
    print("Pulled data for {} samples from {} out of {} participants".format(
        report.shape[0],
        report['participant_id'].unique().size,
        len(participants)
    ))
