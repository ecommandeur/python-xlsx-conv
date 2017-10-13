from openpyxl import load_workbook
from time import strftime
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
parser.add_argument('--delimiter', help='Delimiter used in output, defaults to ,', choices=[',', ';', '|', 'tab'], default=',')
parser.add_argument('--encoding', help='Output encoding, defaults to utf-8', choices=['ascii', 'latin-1', 'utf-8', 'utf-16'], default='utf-8')
parser.add_argument('--extension', help='Extension of output, defaults to csv', default='csv')
parser.add_argument('--noprefix', help='Do not prefix ouput with workbook name', action="store_true")
parser.add_argument('--prefix', help='Use specified prefix instead of prefixing output with workbook name')
parser.add_argument('--quotechar', help='One-character string used to quote fields containing special characters', default='"')
parser.add_argument('--quoting', help='Controls field quoting, defaults to MINIMAL', choices=['ALL', 'MINIMAL', 'NONE', 'NONNUMERIC'], default='MINIMAL')
parser.add_argument('--linebreak_replacement', help='Replace linebreaks in cells by replacement string')
parser.add_argument('--version', action='version', version="%(prog)s 1.1.0dev")
args = parser.parse_args()

inputPath = args.input
outputDir = args.outputDir
noPrefix = args.noprefix
customPrefix = args.prefix
outputExtension = args.extension
outputDelimiter = args.delimiter
outputEncoding = args.encoding
outputQuoteChar = args.quotechar
outputQuoting = args.quoting
linebreakReplacement = args.linebreak_replacement

if not os.path.isfile(inputPath):
    print('xlsx-conv: error: No such file or directory:', inputPath)
    parser.print_usage()
    exit(1)

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

print(strftime("%Y-%m-%d %H:%M:%S"), "- Converting", inputPath)

# ---
# Go ahead and dump that Workbook
# ---

# convert sheet function

def convertSheet(ws,outputPath):
    with open(outputPath, 'w', encoding=outputEncoding) as f:
        c = csv.writer(f, lineterminator='\n', delimiter=outputDelimiter, quotechar=outputQuoteChar, quoting=quoteStyle)
        for row in ws.rows:
            values = []
            for cell in row:
                value = cell.value
                if linebreakReplacement is not None and isinstance(value, str):
                    value = value.replace('\r\n', linebreakReplacement).replace('\n', linebreakReplacement).replace('\r', linebreakReplacement)
                values.append(value)
            c.writerow(values)

# load workbook and invoke convertSheet for all sheets in workbook

try:
    wb = load_workbook(filename=inputPath, read_only=True, data_only=True)
except Exception as e:
    print("xlsx-conv: error: Failed to load workbook")
    print(e)
    exit(1)
    
ws_names = wb.get_sheet_names()

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