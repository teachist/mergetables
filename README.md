# Merge any excel spreadsheet with same the template


![Project Logo](https://raw.githubusercontent.com/starttolearning/mergetables/master/assets/mergetables.png)

## Introduction

This script helps me a lot in my daily workload. My main work was to collect many spreadsheets with same templates from class leaders, when all the things were ready, like cleaned, audited the data, then I need to merge all the table to a single file, you konw it is all the repetitive work and it is very boring, so I had this python scrpt to pull me out.

If you are in the same situation, it may help you. Feel free to use it and get your work down.

## Main feature: Auto collects data from partial excel files

A excel file consists of rows and columns.  However in a rountine, we usually include a header for a excel file, maybe it also contains some merged cells, sometimes a header could use a lot columns and rows. If we want to combine a lot of excel files in one action we need to carefully handle such issues. In school as a data analyzer usually I just need to carefuuly defined a template file and dispatch it to the class leader, in a good lucky, they will return back to a file with right format and good shape, so if I can recongize a mode then it could be used for all the files. For general purpose, may be it can be good for a user to define where the rows and columns should be ignored and which should be kept.  

Now after a fewer tries I got it right. You can use it to merge files(spreadsheets) with same templates at easy.

- Template detection
- Improve from 2018 old one time use
- Using xlwt, xlrd module, because it can operate *.xlsx

## HOW TO INSTALL

1. Clone this repository to your computer.

   `git clone https://github/starttolearning/mergetable`

2. Install from the requirements.txt using your pip
  
  `pip install -r requirements.txt`

3. Install this package to your environment, then you can start to use it

  `pip install --editable .`

## HOW TO USE

```shell
Usage: mgtb [OPTIONS] FOLDER_NAME OUTPUT_FILE

  Terminal setup for merge table

Options:
  --sheet-key INTEGER   Which sheet you want to process.
  --start-row INTEGER   What row you want to start to include in you output
                        file.
  --verify-key INTEGER  Which cell you want to make it as a primary key.
  --end-col INTEGER     What col you want to end.
  --help                Show this message and exit.
```

## MORE INFO

This project will group in the future.
Things that I think it should become much better are following:

- [ ] Terminal friendly
- [ ] Documentation impeletation
- [ ] much flexible for any using case