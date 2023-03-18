import pandas as pd
import argparse
from pathlib import Path

parser = argparse.ArgumentParser()                                               

parser.add_argument("--input", "-i", type=str, required=True)
parser.add_argument("--output", "-o", type=str)
parser.add_argument("--type", "-t", type=str, default="html", help="can convert to 'csv' or 'html'")
parser.add_argument("--sheet", "-s", type=str, help="Sheet Name that you want to convert, program will get all sheets if not specify")
args = parser.parse_args()

def get_outfile_name(end):
  s = args.output
  if s:
    return s if s.endswith(end) else s + end
  else:
    return '.'.join(args.input.split('.')[:-1]) + end

def get_file_type(s):
  return s if s in ['csv', 'html', 'markdown'] else 'csv'

suffix_map = {
  'csv': 'csv',
  'html': 'html',
  'markdown': 'md'
}

inputfile = args.input
filetype = args.type
file_suffix = suffix_map[filetype]
outputfile = get_outfile_name('.%s' % file_suffix)
sheet = args.sheet

def csv_from_excel(input, output, sheet=None, filetype="csv"):
    df = pd.read_excel(input, sheet)
    func_name = 'to_%s' % filetype
    if isinstance(df, pd.DataFrame):
      getattr(df, func_name)(output)
    else:
      sheets = df.keys()
      out_dir = '.'.join(output.split('.')[:-1])
      Path(out_dir).mkdir(parents=True, exist_ok=True)
      for sheet_name in sheets:
          sheet = df[sheet_name]
          getattr(sheet, func_name)("%s/%s.%s" % (out_dir, sheet_name, file_suffix), index=False)
    
# runs the csv_from_excel function:
csv_from_excel(inputfile, outputfile, sheet, filetype)