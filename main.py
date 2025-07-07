from dbconnector.database import engine
import pandas as pd
from datetime import date
from pyodbc import IntegrityError
from tkinter import *
import calendar

rollup_file = "../Forecast ROLLUP.xlsm"
month_name = date.today().strftime("%B")
month = date.today().month
if len(str(month)) == 1:
    month = "0" + str(month)
else:
    month = str(month)
# month_name = "June"
# month = "06"

# pull the correct columns by month
# range function includes the first number but not the last number
# 0 is the first column!!
month_columns = {"June": list(range(30, 38)),
                 "July": list(range(44, 52)),
                 "August": list(range(52, 60)),
                 "September": list(range(60, 68)),
                 "October": [],
                 "November": [],
                 "December": []}
cols = [5] + (month_columns[month_name])

def run_upload():
    rollup_df = pd.read_excel(io=rollup_file, sheet_name='Forecast ROLLUP', header=3, skiprows=[4])
    rollup_df = rollup_df[rollup_df['SFDC ID'].notna()].reset_index(drop=True)
    rollup_df = rollup_df.iloc[:, cols]
    # add a column at beginning with the snapshot date
    rollup_df.insert(0, "SNAPSHOT_DATE", get_etl_date(radio_state.get()))
    # add a second column with the month
    rollup_df.insert(1, "FORECAST_MONTH", "2025_" + month)

    # insert into tblRollup_Snapshot
    with engine.begin() as conn:
        try:
            table_headers = pd.read_sql_query("select * from tblRollup_Snapshot", conn).columns.tolist()
            rollup_df.columns = table_headers
            print("Inserting rows into tblRollup_Snapshot...")
            rollup_df.to_sql('tblRollup_Snapshot', conn, if_exists='append', index=False)
        except IntegrityError:
            print("Primary key violation: there is already data for the given date.")
        finally:
            window.destroy()

##############################TKINTER########################################
window = Tk()
window.title("Forecast Rollup Capture")
window.config(pady=20, padx=20)

frequency = Label(text="What ETL batch is this?")
frequency.pack()

def get_etl_date(sel):
    if sel == "first":
        return f"2025-{month}-01"
    elif sel == "mid":
        return f"2025-{month}-15"
    else:
        return f"2025-{month}-{calendar.monthrange(2025, int(month))[1] - 6}"

def radio_used():
    if radio_state.get():
        run.config(state="normal")  # Enable GO button

radio_state = StringVar()  # No default = none selected

radiobutton1 = Radiobutton(text="1st of the Month", value="first",
                           variable=radio_state, command=radio_used)
radiobutton2 = Radiobutton(text="Mid-month", value="mid",
                           variable=radio_state, command=radio_used)
radiobutton3 = Radiobutton(text="End of Month", value="end",
                           variable=radio_state, command=radio_used)
radiobutton1.pack()
radiobutton2.pack()
radiobutton3.pack()

run = Button(text="GO", command=run_upload, state="disabled", padx=10)
run.pack()

window.mainloop()


