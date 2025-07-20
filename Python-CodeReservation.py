import pyodbc
import os
#print(pyodbc.drivers())
#import os
#print("File exists?", os.path.exists(db_path))


print("---------------  Controlling Reservations Using Password   ----------------")

# Full path to your .mdb file
db_path = r"D:\SUMMER 2025\MIS\Python-CodeReservation__database\Sample_ReservationDatabase.mdb"

print("File exists?", os.path.exists(db_path))
print(pyodbc.drivers())
# Connect using DSN-less connection (no need to configure system DSNs)
try:
   con = pyodbc.connect(
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    fr'DBQ={db_path};'
   )


   cursor = con.cursor()
except pyodbc.Error as e:
   print("❌ Connection failed. Make sure the Access driver is installed and .mdb file exists.")
   print("ODBC Error:", e)
   exit()

# Get user input
input("Press Enter to begin...")
pss = input("Enter password to make reservation (Type 'NO' to view reservations): ")

if pss.lower() == "admin":
    try:
        rid = int(input("Please enter Room No.: "))
        reqid = int(input("Please enter Requestor ID: "))
        Date = input("Please enter date (mm/dd/yyyy): ")
        stime = input("Please enter start time (e.g., 10.0 for 10:00 AM): ")
        etime = input("Please enter end time (e.g., 11.5 for 11:30 AM): ")
        purpose = input("Please enter purpose: ")

        # Duration calculation
        duration = float(etime) - float(stime)

        # Insert into Schedule table
        sql = """INSERT INTO Schedule 
                 (Sch_Room_No, Sch_Req_ID, Sch_Date, Sch_Start_Time, Sch_End_Time, Duration, Sch_Purpose)
                 VALUES (?, ?, ?, ?, ?, ?, ?)"""
        cursor.execute(sql, (rid, reqid, Date, stime, etime, duration, purpose))
        con.commit()

        print("Reservation added successfully.")

    except Exception as e:
        print("Error inserting data:", e)

elif pss.lower() == "no":
    try:
        cursor.execute("SELECT * FROM Schedule")
        results = cursor.fetchall()
        print("\nRoom No | Requestor ID | Schedule Date | Start Time | End Time | Duration | Purpose")
        print("-" * 80)
        for row in results:
            print(f"{row[0]:<8} {row[1]:<13} {row[2]:<15} {row[3]:<12} {row[4]:<10} {row[5]:<9} {row[6]}")
    except Exception as e:
        print(" Error fetching data:", e)

else:
    print(" Invalid input. Please enter 'admin' or 'NO'.")

input("\nPress Enter to exit.")
con.close()
