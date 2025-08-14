from database import engine
import pandas as pd
from datetime import date
from pyodbc import IntegrityError
from tkinter import *
import calendar
from dateutil.relativedelta import relativedelta
from sqlalchemy import text
import math

rollup_file = "../Forecast ROLLUP.xlsm"
last_month_name = (date.today() - relativedelta(months=1)).strftime("%B")
month_name = date.today().strftime("%B")

# get mm for previous month
prev_mm = date.today().month - 1
if prev_mm == 0:
    prev_mm = 12
if len(str(prev_mm)) == 1:
    prev_mm = "0" + str(prev_mm)
else:
    prev_mm = str(prev_mm)

# get mm for current month
mm = date.today().month
if len(str(mm)) == 1:
    mm = "0" + str(mm)
else:
    mm = str(mm)

# pull the correct columns by month
# range function includes the first number but not the last number
# 0 is the first column!!
month_columns = {"June": list(range(30, 38)),
                 "July": list(range(44, 52)),
                 "August": list(range(52, 60)),
                 "September": list(range(60, 68)),
                 "October": list(range(74, 82)),
                 "November": list(range(82, 90)),
                 "December": list(range(90,98))}

# get the columns containing last month's data
prev_cols = [5] + (month_columns[last_month_name])
# get the columns containing this month's data
cols = [5] + (month_columns[month_name])


def insert_into_table(pddf: pd.DataFrame, connection):
    """Inserts the given pddf into tblRollup_Snapshot."""
    try:
        table_headers = pd.read_sql_query("select * from tblRollup_Snapshot", connection).columns.tolist()
        pddf.columns = table_headers
        print("Inserting rows into tblRollup_Snapshot...")
        pddf.to_sql('tblRollup_Snapshot', connection, if_exists='append', index=False)
        print(f"Successfully inserted rows into tblRollup_Snapshot.")
    except IntegrityError:
        print("Primary key violation: there is already data for the given date.")


def update_variance_explanations(pddf: pd.DataFrame, connection):
    """Update variance explanations for the previous month. This is necessary because it takes a few days after
    the start of a new month for reps to enter in accurate variance explanations for the closed month."""
    def none_if_nan(value):
        return None if (value is None or (isinstance(value, float) and math.isnan(value))) else value
    print("Updating variance explanations from last month...")
    for index, row in pddf.iterrows():
        try:
            update_sql = text("""UPDATE tblRollup_Snapshot
                            SET [IMPL_REASON_FOR_VARIANCE] = :impl,
                            [REV_REASON_FOR_VARIANCE] = :rev
                            WHERE [ACT_ID] = :id
                            AND [FORECAST_MONTH] = :month""")

            connection.execute(update_sql, {
                'impl': none_if_nan(row['IMPL_REASON_FOR_VARIANCE']),
                'rev': none_if_nan(row['REV_REASON_FOR_VARIANCE']),
                'id': row['SFDC ID'],
                'month': row['FORECAST_MONTH']
            })

        except Exception as e:
            print(f"An error occurred while updating last month's variance explanations: {e}")

    print(f"Updated last month's variance explanations.")


def get_etl_date(sel):
    if sel == "first":
        return f"2025-{mm}-01"
    elif sel == "mid":
        return f"2025-{mm}-15"
    else:
        return f"2025-{mm}-{calendar.monthrange(2025, int(mm))[1] - 6}"


def run_upload():
    df = pd.read_excel(io=rollup_file, sheet_name='Forecast ROLLUP', header=3, skiprows=[4])
    df = df[df['SFDC ID'].notna()].reset_index(drop=True)
    # get columns for the current month
    rollup_df = df.iloc[:, cols]
    # add a column at beginning with the snapshot date
    rollup_df.insert(0, "SNAPSHOT_DATE", get_etl_date(radio_state.get()))
    # add a second column with the month
    rollup_df.insert(1, "FORECAST_MONTH", "2025_" + mm)

    # create new df, get columns from previous month
    variance_df = df.iloc[:, prev_cols]
    # keep only sfdc_id and variance columns
    variance_df = variance_df.iloc[:, [0, 4, 8]]
    variance_df.insert(0, "FORECAST_MONTH", "2025_" + prev_mm)
    variance_df.columns = ["FORECAST_MONTH", "SFDC ID", "IMPL_REASON_FOR_VARIANCE", "REV_REASON_FOR_VARIANCE"]

    variance_df.to_excel("variance_df.xlsx", index=False)
    rollup_df.to_excel("rollup_df.xlsx", index=False)


    with engine.begin() as conn:
        insert_into_table(rollup_df, conn)
        # if snapshot date is mid-month, update the variance explanations for the prior month
        if radio_state.get() == f"mid" or radio_state.get() == f"first":
            update_variance_explanations(variance_df, conn)

    window.destroy()


##############################TKINTER########################################
window = Tk()
window.title("Forecast Rollup Capture")
window.config(pady=20, padx=20)

frequency = Label(text="What ETL batch is this?")
frequency.pack()

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


