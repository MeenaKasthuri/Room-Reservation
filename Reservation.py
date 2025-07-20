import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import os
from datetime import date
from datetime import datetime

# Database Path
db_path = r"D:\SUMMER 2025\MIS\Python-CodeReservation__database\Sample_ReservationDatabase.mdb"

if not os.path.exists(db_path):
    messagebox.showerror("Database Error", "Database file not found.")
    exit()

#  DB Connection
try:
    con = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        fr'DBQ={db_path};'
    )
    cursor = con.cursor()
except pyodbc.Error as e:
    messagebox.showerror("Connection Error", str(e))
    exit()

# Helper: convert time string like "10am" or "1:30pm" to float hour
def time_str_to_float(time_str):
    try:
        dt = datetime.strptime(time_str.strip().lower(), "%I:%M%p")
    except ValueError:
        dt = datetime.strptime(time_str.strip().lower(), "%I%p")
    return dt.hour + dt.minute / 60

def float_to_ampm(time_float):
    hours = int(time_float)
    minutes = int(round((time_float - hours) * 60))
    dt = datetime.strptime(f"{hours}:{minutes}", "%H:%M")
    return dt.strftime("%I:%M %p").lstrip("0")

def ampm_str_to_float(time_str):
    dt = datetime.strptime(time_str.strip(), "%I:%M %p")
    return dt.hour + dt.minute / 60

# GUI Window
root = tk.Tk()
root.title("Room Reservation")
root.geometry("700x400")
root.resizable(False, False)

# Treeview
style = ttk.Style()
style.configure("Treeview", rowheight=24, font=('Segoe UI', 10), borderwidth=0)
style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'))
style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

cols = ("Room", "Requestor", "Date", "Start", "End", "Duration", "Purpose")
tree = ttk.Treeview(root, columns=cols, show="headings", height=8)
for col in cols:
    tree.heading(col, text=col)
    tree.column(col, width=90, anchor="center")
tree.pack(pady=10, padx=10)

def format_row(row):
    formatted = []
    for i, value in enumerate(row):
        if isinstance(value, date):
            formatted.append(value.strftime("%Y-%m-%d"))
        elif i == 3 or i == 4:  # Start Time or End Time columns
            try:
                formatted.append(float_to_ampm(float(value)))
            except Exception:
                formatted.append(str(value))
        else:
            formatted.append(str(value))
    return tuple(formatted)

def load_data():
    tree.delete(*tree.get_children())
    cursor.execute("SELECT * FROM Schedule")
    for row in cursor.fetchall():
        tree.insert("", "end", values=format_row(row))

def add_reservation():
    try:
        rid = int(entry_room.get())
        reqid = int(entry_req.get())
        date_val = datetime.strptime(entry_date.get(), "%m/%d/%Y").date()

        start = time_str_to_float(entry_start.get())
        end = time_str_to_float(entry_end.get())

        if end <= start:
            messagebox.showerror("Input Error", "End time must be after start time.")
            return

        duration = end - start
        purpose = entry_purpose.get()

        cursor.execute(
            """INSERT INTO Schedule 
            (Sch_Room_No, Sch_Req_ID, Sch_Date, Sch_Start_Time, Sch_End_Time, Duration, Sch_Purpose)
            VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (rid, reqid, date_val, start, end, duration, purpose)
        )
        con.commit()
        load_data()
        clear_entries()
        messagebox.showinfo("Success", "Reservation added.")
    except ValueError as ve:
        messagebox.showerror("Input Error", f"Invalid input:\n{ve}")
    except Exception as e:
        messagebox.showerror("Error", f"Add failed:\n{e}")

def clear_entries():
    entry_room.delete(0, tk.END)
    entry_req.delete(0, tk.END)
    entry_date.delete(0, tk.END)
    entry_start.delete(0, tk.END)
    entry_end.delete(0, tk.END)
    entry_purpose.delete(0, tk.END)

def delete_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning", "No reservation selected to delete.")
        return
    item = tree.item(selected[0])["values"]
    try:
        # Convert date string back to date object
        sch_date = datetime.strptime(item[2], "%Y-%m-%d").date()
        # Convert displayed start time string back to float
        sch_start_time = ampm_str_to_float(item[3])

        cursor.execute(
            "DELETE FROM Schedule WHERE Sch_Room_No=? AND Sch_Req_ID=? AND Sch_Date=? AND Sch_Start_Time=?",
            (int(item[0]), int(item[1]), sch_date, sch_start_time)
        )
        con.commit()
        load_data()
        messagebox.showinfo("Deleted", "Reservation deleted.")
    except Exception as e:
        messagebox.showerror("Error", f"Delete failed:\n{e}")

form = tk.Frame(root)
form.pack()

def make_field(label_text, row, col):
    tk.Label(form, text=label_text).grid(row=row, column=col, sticky="w")
    entry = tk.Entry(form, width=12)
    entry.grid(row=row, column=col + 1, padx=5, pady=2)
    return entry

entry_room = make_field("Room No:", 0, 0)
entry_req = make_field("Requestor ID:", 0, 2)
entry_date = make_field("Date (MM/DD/YYYY):", 1, 0)
entry_start = make_field("Start Time (e.g. 10am):", 1, 2)
entry_end = make_field("End Time (e.g. 4pm):", 2, 0)
entry_purpose = make_field("Purpose:", 2, 2)

btns = tk.Frame(root)
btns.pack(pady=5)

tk.Button(btns, text="Add", width=10, command=add_reservation).pack(side="left", padx=10)
tk.Button(btns, text="Delete", width=10, command=delete_selected).pack(side="left", padx=10)
tk.Button(btns, text="Refresh", width=10, command=load_data).pack(side="left", padx=10)

load_data()
root.mainloop()
