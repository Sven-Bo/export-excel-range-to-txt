from pathlib import Path

import pandas as pd  # pip install pandas
import xlwings as xw  # pip install xlwings as xw

# Define excel file path
BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "data.xlsx"


# Open Excel instance in background
with xw.App(visible=False) as app:
    wb = app.books.open(EXCEL_FILE)  # read in Excel file
    ws = wb.sheets[0]  # use the first worksheet (index position 0)
    df_named_range = (
        ws.range("Germany").options(pd.DataFrame).value
    )  # convert named range to pandas dataframe

# write contents of a DataFrame into a text file
dfAsString = df_named_range.to_string(header=True, index=False)
Path(BASE_DIR / "Germany.txt").write_text(dfAsString)
