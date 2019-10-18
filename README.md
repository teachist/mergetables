# FOR EASY WORK IN EVERYDAY IN YUMEN(EWEY)

## FEATURE #1: AUTO COLLECTS DATAS FROM PARTIAL EXCEL FILES

A excel file consists of rows and columns.  However in a rountines, we usually include a header for a excel file, maybe it also using some merge cells, usually a header can use a lot columns and rows. If we want to combine a lot of excel files in one action we need to carefully handle such issues. In a school as a data analyzer usually I just need to carefuuly defined a good templates file and dispatch it to the class leader, in a good lucky, they will return back to a files with right format and good shape, so if I can recongize a mode then it can be used for all the file. For general purpose, may be it can be good for a user to define where the rows and columns should be ignored and which should be kept.  

Now you can use it to merge files(spreadsheets) with same templates.

- Improve from 2018 app.py
- Using xlwt, xlrd module, because it can operate *.xlsx

## HOW TO USE

COMMANDS

```python3
python appv3.py xpx sample 0 5 1
```

PS:

- `appv3.py` is our script file writen by python3
- `xpx` is the files where you want to be merged
- `smaple` is the output file
- `0` define which sheet you want to merge in each file
- `5` define where the rows you want to start to merge
- `1` define where the verify key you choose
