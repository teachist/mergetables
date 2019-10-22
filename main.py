import click

from mgtb.mgtb import mergetable

'''
Usage: mgtb [OPTIONS] FOLDERNAME OUTPUTFILE

  Simple program that greets NAME for a total of COUNT times.

Options:
  --sheet-key INTEGER  Which sheet you want to process.
  --start-row INTEGER  What row you want to start to include in you output file.
  --verify-key INTEGER  Which cell you want to make it as a primary key.
  --end-col INTEGER     What col you want to end.

  --help           Show this message and exit.

'''
@click.command()
@click.argument('folder_name')
@click.argument('output_file')
@click.option('--sheet-key', default=0, help='Which sheet you want to process.')
@click.option('--start-row', default=1, help='What row you want to start to include in you output file.')
@click.option('--verify-key', default=1, help='Which cell you want to make it as a primary key.')
@click.option('--end-col',  help='What col you want to end.', type=int)
def cli(folder_name, output_file, sheet_key, start_row, verify_key, end_col):
    """ Terminal setup for merge table"""
    click.echo('Hello World!')
    # click.echo(folder_name, output_file, sheet_key)
    mergetable(folder_name, output_file, sheet_key=sheet_key,
               start_row=start_row, verify_key=verify_key, end_col=end_col)
