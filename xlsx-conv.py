from openpyxl import load_workbook
from time import strftime
from itertools import islice
import argparse
import csv
import os

# ---
# Get arguments
# ---
# see https://docs.python.org/3/howto/argparse.html 

parser = argparse.ArgumentParser(description='Convert XLSX file to DSV using openpyxl')
parser.add_argument('-i','--input', help='Path to XLSX file', required=True)
parser.add_argument('-o','--outputDir', help='Path to output directory')
parser.add_argument('--col_index', help='Generate column indices (c1,c2,etc) as first line in output', action="store_true")
parser.add_argument('--delimiter', help='Delimiter used in output, defaults to ,', choices=[',', ';', '|', 'tab'], default=',')
parser.add_argument('--encoding', help='Output encoding, defaults to utf-8', choices=['ascii', 'latin-1', 'utf-8', 'utf-16'], default='utf-8')
parser.add_argument('--extension', help='Extension of output, defaults to csv', default='csv')
parser.add_argument('--linebreak_replacement', help='Replace linebreaks in cells by replacement string')
parser.add_argument('--noprefix', help='Do not prefix ouput with workbook name', action="store_true")
parser.add_argument('--prefix', help='Use specified prefix instead of prefixing output with workbook name')
parser.add_argument('--row_index', help='Write row numbers as first column in output', action="store_true")
parser.add_argument('--quotechar', help='One-character string used to quote fields containing special characters', default='"')
parser.add_argument('--quoting', help='Controls field quoting, defaults to MINIMAL', choices=['ALL', 'MINIMAL', 'NONE', 'NONNUMERIC'], default='MINIMAL')
parser.add_argument('--version', action='version', version="%(prog)s 1.2.0-SNAPSHOT")
args = parser.parse_args()

inputPath               = args.input
inputPath               = args.input
inputPath               = args.input
outputDir               = args.outputDir
outputDir               = args.outputDir

colIndex                = args.col_index 
outputDelimiter         = args.delimiter 
outputEncoding          = args.encoding 
outputExtension         = args.extension 
linebreakReplacement    = args.linebreak_replacement
noPrefix                = args.noprefix
customPrefix            = args.prefix
rowIndex                = args.row_index
outputQuoteChar         = args.quotechar
outputQuoting           = args.quoting

if not os.path.isfile(inputPath):
    print('xlsx-conv: error: No such file or directory:', inputPath)
    parser.print_usage()
    exit(1)

inputPath = os.path.realpath(inputPath)
inputDir, inputFile = os.path.split(inputPath)
inputBaseFn, inputExt = os.path.splitext(inputFile)

if outputDir:
    if not os.path.isdir(outputDir):
        parser.print_usage()
        print('xlsx-conv: error: No such file or directory:', outputDir)
        exit(1)
else:
   outputDir = inputDir

quoteStyle = csv.QUOTE_MINIMAL
if outputQuoting == "ALL":
    quoteStyle = csv.QUOTE_ALL
elif outputQuoting == "NONE":
    quoteStyle = csv.QUOTE_NONE
elif outputQuoting == "NONNUMERIC":
    quoteStyle = csv.QUOTE_NONNUMERIC

# ---
# Go ahead and dump that Workbook
# ---

print(strftime("%Y-%m-%d %H:%M:%S"), "- Converting", inputPath)

# convert sheet function
#  If row index is set, but column index is not then an index is just inserted before each record

def convertSheet(ws,outputPath):
    with open(outputPath, 'w', encoding=outputEncoding) as f:
        c = csv.writer(f, lineterminator='\n', delimiter=outputDelimiter, quotechar=outputQuoteChar, quoting=quoteStyle)
        
        # First check if there is are rows in ws.rows 
        first_row_slice = list(islice(ws.rows,1))
        if len(first_row_slice) == 0:
            return

        if colIndex == True:
            first_row =  first_row_slice[0]
            col_index = []
            if rowIndex == True:
                col_index.append("c0") # if row_index is also set then include additional column
            for i in range(len(first_row)):
                c_val = "c" + str(i+1)
                col_index.append(c_val)
            c.writerow(col_index)

        for index, row in enumerate(ws.rows):
            values = []
            if rowIndex == True:
                values.append(index+1)
            for cell in row:
                value = cell.value
                if linebreakReplacement is not None and isinstance(value, str):
                    value = value.replace('\r\n', linebreakReplacement).replace('\n', linebreakReplacement).replace('\r', linebreakReplacement)
                values.append(value)
            c.writerow(values)

# ---
# load workbook and invoke convertSheet for all sheets in workbook
# ---

try:
    wb = load_workbook(filename=inputPath, read_only=True, data_only=True)
except Exception as e:
    print("xlsx-conv: error: Failed to load workbook")
    print(e)
    exit(1)
    
ws_names = wb.sheetnames

outputPrefix = inputBaseFn + '.'
if customPrefix:
    outputPrefix = customPrefix + '.' # override with custom prefix
if noPrefix == True:
    outputPrefix = '' # noprefix takes precedence over customPrefix

if outputDelimiter == 'tab':
    outputDelimiter = '\t'

for ws_name in ws_names:
    ws = wb[ws_name] # ws is now an IterableWorksheet
    outputPath = outputDir + os.sep + outputPrefix + ws_name + '.' + outputExtension
    print(strftime("%Y-%m-%d %H:%M:%S"), "- Outputting sheet to", outputPath)
    try:
        convertSheet(ws,outputPath)
    except Exception as e:
        print("xlsx-conv: error: Failed to convert sheet")
        print(e)
        exit(1) 

print(strftime("%Y-%m-%d %H:%M:%S"), "- bye!")