echo off
REM python xlsx-conv.py -i "resources\Numbers.xlsx"
python xlsx-conv.py -i "resources\Formulas.xlsx" --extension "txt"
REM no prefix
REM python xlsx-conv.py -i "resources\Numbers.xlsx" --noprefix
REM custom prefix
REM python xlsx-conv.py -i "resources\Numbers.xlsx" --prefix "CUSTOM"