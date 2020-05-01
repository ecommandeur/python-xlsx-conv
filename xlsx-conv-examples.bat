echo off
if exist temp\ rmdir /S/Q temp\
mkdir temp
REM -i is the only required argument
copy "resources\Characters.xlsx" "temp"
python xlsx-conv.py -i "temp\Characters.xlsx"
REM -o can be used to specify the output dir
python xlsx-conv.py -i "resources\Empty.xlsx" -o "temp"
python xlsx-conv.py -i "resources\Filtered.xlsx" -o "temp"
python xlsx-conv.py -i "resources\Formulas.xlsx" -o "temp"
python xlsx-conv.py -i "resources\Hidden.xlsx" -o "temp"
python xlsx-conv.py -i "resources\MultipleSheets.xlsx" -o "temp"
python xlsx-conv.py -i "resources\MultipleSheets.xlsx" --sheetnames
python xlsx-conv.py -i "resources\Numbers.xlsx" -o "temp"
python xlsx-conv.py -i "resources\Range.xlsx" -o "temp"
REM other options
python xlsx-conv.py -i "resources\Range.xlsx" -o "temp" --delimiter "|" --extension "txt"
python xlsx-conv.py -i "resources\Range.xlsx" -o "temp" --encoding "ascii" --extension "ascii.txt"
python xlsx-conv.py -i "resources\Range.xlsx" -o "temp" --noprefix
python xlsx-conv.py -i "resources\Range.xlsx" -o "temp" --prefix "R"
python xlsx-conv.py -i "temp\Characters.xlsx" --extension "lr.csv" --linebreak_replacement " "
python xlsx-conv.py -i "temp\Characters.xlsx" --extension "q.csv" --quoting "ALL"
REM loading input from file
python xlsx-conv.py -i "resources\Input_OutputDir.txt" -o "temp"
REM graceful error upon missing input
python xlsx-conv.py -i "resources\NonExistent.txt" -o "temp"
python xlsx-conv.py -i "resources\NonExistent.xlsx" -o "temp"