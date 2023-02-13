from tkinter import *
from tkinter import filedialog
import customtkinter
import os
import mysql.connector
import pandas as pd
from pandas.io import sql as sql
from pyxlsb import open_workbook as open_xlsb
import openpyxl

global store_num
global date

global receiving_data_path
global new_epcs_path
global qb_master_path
global item_file_path

global app
global conn
global cursor

global rec_over_list
rec_over_list = []


def import_receiving_data():
    print("Receiving...")
    pop_up_title = "Select Receiving Data (.csv)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
    global receiving_data_path
    receiving_data_path = filename
    print(receiving_data_path)


def import_new_epcs():
    print("New EPCs...")
    pop_up_title = "Select New EPCs (.xlsx)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
    global new_epcs_path
    new_epcs_path = filename
    print(new_epcs_path)


def import_qb_master_items():
    print("QB Master Items...")
    pop_up_title = "Select QB Master Items (.xlsb)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("xlsb files", "*.xlsb"), ("all files", "*.*")))
    global qb_master_path
    qb_master_path = filename
    print(qb_master_path)


def import_item_file():
    print("Item File...")
    pop_up_title = "Select Item File (GM) (.csv)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
    global item_file_path
    item_file_path = filename
    print(item_file_path)


def open_about():
    print("About...")


def open_settings():
    print("Settings...")


def quit_app():
    print("Quit...")
    app.quit()


def submit_info():
    return_value = True
    for func in [store_number_verification, date_verification, receiving_path_verification, new_epcs_path_verification,
                 qb_path_verification, item_file_path_verification]:
        if not func():
            print(f"{func.__name__} input is invalid or unspecified!")
            return_value = False
    return return_value


def connect_to_mysql():
    try:
        global conn
        conn = mysql.connector.connect(user='root', password='password', host='127.0.0.1', database='receivingsystem',
                                       allow_local_infile=True)
        global cursor
        cursor = conn.cursor()
        stmt00 = "SET GLOBAL local_infile=1;"
        cursor.execute(stmt00)
        print("Connected to MySQL...")
    except:
        print(":: ERROR :: Something went wrong! Unable to connect to MySQL!")


def import_receiving_sql():
    try:
        stmt = "DROP TABLE IF EXISTS ReceivingData"
        cursor.execute(stmt)
        statement_headers = "CREATE TABLE ReceivingData(date_hour text, Event_Timestamp bigint, STORE_NBR int, " \
                            "dept_nbr int, CID bigint, eGTIN bigint, Transaction_Type text, Transaction_QTY int, " \
                            "InventoryState text, Transaction_QTY2 int, Aggregate_Qty int)"
        cursor.execute(statement_headers)
        receiving_corrected = receiving_data_path.replace(" ", "\\ ")

        stmt = "LOAD DATA LOCAL INFILE \'{}\' " \
               "INTO TABLE ReceivingData " \
               "CHARACTER SET latin1 " \
               "FIELDS TERMINATED BY \',\' " \
               "ENCLOSED BY \'\"\' " \
               "LINES TERMINATED BY \'\\r\\n\' " \
               "IGNORE 1 ROWS;".format(receiving_corrected)
        print(" -- Starting Receiving Data import...")
        cursor.execute(stmt)
        print(" -- Receiving Data import complete.")
    except Exception as e:
        print(":: ERROR :: Could not import Receiving Data!")
        print(e)


def import_new_epcs_sql():
    try:
        stmt = "DROP TABLE IF EXISTS NewEPCs"
        cursor.execute(stmt)
        '''
            * Import Active EPC Directory file.
            * Select 2nd sheet in workbook
            * Import entire New EPC Directory sheet to SQL
            * Filter sheet once in SQL to New EPCs only
        '''
        df = []
        print(" -- Converting New EPCs to .csv")
        # df = pd.DataFrame(pd.read_excel(new_epcs_path))
        read_file = pd.read_excel(new_epcs_path, sheet_name=1)
        # df = pd.DataFrame(df[1:], columns=df[0])

        new_epcs_csv_path = os.path.splitext(new_epcs_path)[0]
        new_epcs_csv_path += ".csv"
        count = 1
        while os.path.exists(new_epcs_csv_path):
            new_epcs_csv_path = os.path.splitext(new_epcs_path)[0] + " (" + str(count) + ")" + ".csv"
            count += 1

        read_file.to_csv(new_epcs_csv_path, index=False)
        df = pd.DataFrame(pd.read_csv(new_epcs_csv_path))
        # df.to_csv(new_epcs_csv_path, index=False)  # to generate a .csv file

        with open(new_epcs_csv_path, "rb") as file:
            lines = file.readlines()

        with open(new_epcs_csv_path, "wb") as file:
            for line in lines:
                file.write(line.replace(b"\r\n", b"\n"))
        print(" -- New EPCs file conversion to .csv complete.")

        statement_headers = "CREATE TABLE NewEPCs(EPC text, UPC text, Latest_Date_Seen text, Status text)"
        cursor.execute(statement_headers)

        new_epcs_corrected = new_epcs_csv_path.replace(" ", "\\ ")

        stmt1 = "LOAD DATA LOCAL INFILE \'{}\' " \
                "INTO TABLE NewEPCs " \
                "CHARACTER SET latin1 " \
                "FIELDS TERMINATED BY \',\'" \
                "OPTIONALLY ENCLOSED BY \'\"\' " \
                "LINES TERMINATED BY \'\\n\' " \
                "IGNORE 1 ROWS;".format(new_epcs_corrected)

        print(" -- Starting New EPCs import...")
        cursor.execute(stmt1)
        print(" -- New EPCs import complete.")
        stmt1 = "CREATE TABLE temp AS " \
                "SELECT * " \
                "FROM newepcs " \
                "WHERE newepcs.Status = \"New\";"
        cursor.execute(stmt1)
        stmt2 = "DROP TABLE IF EXISTS NewEPCs;"
        cursor.execute(stmt2)
        stmt3 = "CREATE TABLE NewEPCs AS " \
                "SELECT * FROM temp;"
        cursor.execute(stmt3)
        cursor.execute("DROP TABLE IF EXISTS temp")

        conn.commit()

    except Exception as e:
        print(":: ERROR :: Could not import New EPCs file!")
        print(e)


def import_active_epcs_sql():
    try:
        stmt = "DROP TABLE IF EXISTS ActiveEPCs"
        cursor.execute(stmt)

        df = []
        print(" -- Converting Active EPCs to .csv")
        read_file = pd.read_excel(new_epcs_path, sheet_name=0)

        active_epcs_csv_path = os.path.splitext(new_epcs_path)[0]
        active_epcs_csv_path += "_Active.csv"
        count = 1

        while os.path.exists(active_epcs_csv_path):
            active_epcs_csv_path = os.path.splitext(new_epcs_path)[0] + "_Active (" + str(count) + ")" + ".csv"
            count += 1

        read_file.to_csv(active_epcs_csv_path, index=False)
        df = pd.DataFrame(pd.read_csv(active_epcs_csv_path))

        with open(active_epcs_csv_path, "rb") as file:
            lines = file.readlines()

        with open(active_epcs_csv_path, "wb") as file:
            for line in lines:
                file.write(line.replace(b"\r\n", b"\n"))
        print(" -- Active EPCs file conversion to .csv complete.")

        statement_headers = "CREATE TABLE ActiveEPCs(EPC text, UPC text, Latest_Date_Seen text)"
        cursor.execute(statement_headers)

        active_epcs_corrected = active_epcs_csv_path.replace(" ", "\\ ")

        stmt1 = "LOAD DATA LOCAL INFILE \'{}\' " \
                "INTO TABLE ActiveEPCs " \
                "CHARACTER SET latin1 " \
                "FIELDS TERMINATED BY \',\'" \
                "OPTIONALLY ENCLOSED BY \'\"\' " \
                "LINES TERMINATED BY \'\\n\' " \
                "IGNORE 1 ROWS;".format(active_epcs_corrected)

        print(" -- Starting Active EPCs import...")
        cursor.execute(stmt1)
        print(" -- Active EPCs import complete.")
        conn.commit()

    except Exception as e:
        print(":: ERROR :: Could not import Active EPCs file!")
        print(e)


def import_qb_sql():
    try:
        stmt = "DROP TABLE IF EXISTS QBMasterItems"
        cursor.execute(stmt)
        # df = pd.DataFrame(pd.read_excel(qb_master_path))
        df = []

        print(" -- Converting QB Master Items File to .csv")
        with open_xlsb(qb_master_path) as wb:
            with wb.get_sheet(3) as sheet:
                for row in sheet.rows():
                    df.append([item.v for item in row])

        df = pd.DataFrame(df[1:], columns=df[0])

        qb_csv_path = os.path.splitext(qb_master_path)[0]
        qb_csv_path += ".csv"
        count = 1
        while os.path.exists(qb_csv_path):
            qb_csv_path = os.path.splitext(qb_master_path)[0] + " (" + str(count) + ")" + ".csv"
            count += 1

        df.to_csv(qb_csv_path, index=False)  # to generate a .csv file

        with open(qb_csv_path, "rb") as file:
            lines = file.readlines()

        with open(qb_csv_path, "wb") as file:
            for line in lines:
                file.write(line.replace(b"\r\n", b"\n"))

        print(" -- QB Master Items file conversion to .csv complete.")

        statement_headers = "CREATE TABLE QBMasterItems(Record_ID_NBR text, Items_Record_ID_NBRs text, Item_Validation_Status text, " \
                            "Item_Arrival_Status text, Vendor_Number text, Vendor_Name text, Dept_NBR text, SBU text, UPC text, " \
                            "Item_Description text, Arrival_Month text, Max_Shipped_On_Date text, Offshore text)"
        cursor.execute(statement_headers)

        qb_path_corrected = qb_csv_path.replace(" ", "\\ ")

        stmt1 = "LOAD DATA LOCAL INFILE \'{}\' " \
            "INTO TABLE QBMasterItems " \
            "CHARACTER SET latin1 " \
            "FIELDS TERMINATED BY \',\'" \
            "OPTIONALLY ENCLOSED BY \'\"\' " \
            "LINES TERMINATED BY \'\\n\' " \
            "IGNORE 1 ROWS;".format(qb_path_corrected)

        print(" -- Starting QB Master Items import...")
        cursor.execute(stmt1)
        cursor.execute("UPDATE qbmasteritems SET UPC = REPLACE(UPC, \".0\", \"\");")  # Removes '.0' from end of UPC.
        print(" -- QB Master Items import complete.")
        conn.commit()

    except Exception as e:
        print(":: ERROR :: Could not import QB Master Items file!")
        print(e)


def import_item_file_sql():
    stmt = "DROP TABLE IF EXISTS ItemFile"
    cursor.execute(stmt)
    statement_headers = "CREATE TABLE ItemFile(store_number int, REPL_GROUP_NBR int, gtin bigint, ei_onhand_qty int, " \
                        "SNAPSHOT_DATE text, UPC_NBR bigint, UPC_DESC text, ITEM1_DESC text, dept_nbr int, " \
                        "DEPT_DESC text, MDSE_SEGMENT_DESC text, MDSE_SUBGROUP_DESC text, ACCTG_DEPT_DESC text, " \
                        "DEPT_CATG_GRP_DESC text, DEPT_CATEGORY_DESC text, DEPT_SUBCATG_DESC text, VENDOR_NBR int, " \
                        "VENDOR_NAME text, BRAND_OWNER_NAME text, BRAND_FAMILY_NAME text)"
    cursor.execute(statement_headers)

    # --------------Loads both Item Files into single ItemFile table------------------------------------------------
    item_file_path_corrected = item_file_path.replace(" ", "\\ ")

    stmt = "LOAD DATA LOCAL INFILE \'{}\' " \
           "INTO TABLE ItemFile " \
           "CHARACTER SET latin1 " \
           "FIELDS TERMINATED BY \',\' " \
           "OPTIONALLY ENCLOSED BY \'\"\' " \
           "LINES TERMINATED BY \'\\r\\n\' " \
           "IGNORE 1 ROWS;".format(item_file_path_corrected)

    print(" -- Starting Item File import...")
    cursor.execute(stmt)
    print(" -- Item File import complete.")
    conn.commit()


def create_upc_drop_sql():
    print(" -- Creating UPC Drop table...")
    stmt = "DROP TABLE IF EXISTS UPCDrop"
    cursor.execute(stmt)
    stmt1 = "CREATE TABLE UPCDrop AS " \
            "SELECT activeepcs.UPC " \
            "FROM activeepcs " \
            "WHERE activeepcs.Latest_Date_Seen = (SELECT MAX(activeepcs.Latest_Date_Seen) FROM activeepcs);"
    cursor.execute(stmt1)
    conn.commit()
    print(" -- UPC Drop table creation complete.")


def create_total_items_sql():
    print(" -- Creating Total Items Table...")
    stmt = "DROP TABLE IF EXISTS TotalItems"
    cursor.execute(stmt)
    stmt1 = "CREATE TABLE TotalItems AS " \
            "SELECT itemfile.gtin, " \
            "itemfile.DEPT_CATG_GRP_DESC, " \
            "itemfile.DEPT_CATEGORY_DESC, " \
            "itemfile.VENDOR_NBR, " \
            "itemfile.VENDOR_NAME, " \
            "itemfile.BRAND_FAMILY_NAME, " \
            "itemfile.dept_nbr " \
            "FROM itemfile " \
            "INNER JOIN UPCDrop ON UPCDrop.UPC = itemfile.gtin " \
            "WHERE UPCDrop.UPC = itemfile.gtin"
    cursor.execute(stmt1)
    conn.commit()
    print(" -- Total Items Table creation complete.")


def create_transactions_gm_sql():
    print(" -- Creating Transactions (GM) table...")
    stmt = "DROP TABLE IF EXISTS transactions_gm;"
    cursor.execute(stmt)
    stmt1 = "CREATE TABLE transactions_gm AS " \
            "SELECT * FROM receivingdata " \
            "WHERE dept_nbr IN ('7','9','14','17','20','22','71','72','74','87') " \
            "ORDER BY receivingdata.Event_Timestamp DESC;"
    cursor.execute(stmt1)
    print(" -- Transactions table creation complete.")


def create_receiving_gm_sql():
    print(" -- Creating Receiving (GM) table...")
    stmt = "DROP TABLE IF EXISTS receiving_gm;"
    cursor.execute(stmt)
    stmt1 = "CREATE TABLE receiving_gm AS " \
            "SELECT * FROM transactions_gm " \
            "WHERE Transaction_Type = \"Receiving\" " \
            "ORDER BY Event_Timestamp DESC;"
    cursor.execute(stmt1)
    print(" -- Receiving (GM) table creation complete.")


def create_upc_no_check_sql():
    cursor.execute("DROP TABLE IF EXISTS UPC_No_Check_Digit;")
    stmt = "CREATE TABLE UPC_No_Check_Digit AS " \
           "SELECT Receiving_GM.eGTIN, LEFT(Receiving_GM.eGTIN, LENGTH(Receiving_GM.eGTIN) - 1) " \
           "AS UPC_No_Check FROM Receiving_GM;"
    cursor.execute(stmt)


def create_receiving_overview_sql():
    cursor.execute("DROP TABLE IF EXISTS Receiving_Overview;")
    stmt1 = """
    CREATE TABLE Receiving_Overview AS SELECT
    COALESCE(ItemFile.Vendor_NBR, 'Not Found in Item File') AS Vendor_NBR,
    Receiving_GM.eGTIN,
    COALESCE(TotalItemsCount.Total_RFID, 0) AS Total_RFID,
    COALESCE(ItemFile.ei_onhand_qty, 'Not Found in Item File') AS ei_onhand_qty,
    COALESCE(Receiving_GM_Sum.Receiving_Total, 0) AS Receiving_Total,
    COALESCE(Transactions_GM_Sum.Sum_Transactions_Total, 0) AS Sum_Transactions_Total,
    COALESCE(New_EPCs_Count.New_EPC_Total, 0) AS New_EPC_Total,
    CASE
    WHEN COALESCE(Transactions_GM_Sum.Sum_Transactions_Total, 0) = COALESCE(New_EPCs_Count.New_EPC_Total, 0) THEN 'Match'
    WHEN COALESCE(Transactions_GM_Sum.Sum_Transactions_Total, 0) > COALESCE(New_EPCs_Count.New_EPC_Total, 0) THEN 'Under'
    ELSE 'Over'
    END AS Matches,
    COALESCE(Agg_Qty, 0) AS Aggregate_Qty,
    COALESCE(ItemFile.dept_catg_grp_desc, 'Not Found in Item File') AS dept_catg_grp_desc,
    COALESCE(ItemFile.dept_category_desc, 'Not Found in Item File') AS dept_category_desc,
    COALESCE(ItemFile.vendor_name, 'Not Found in Item File') AS Vendor_Name,
    COALESCE(ItemFile.brand_family_name, 'Not Found in Item File') AS Brand_Family_Name,
    COALESCE(ItemFile.dept_nbr, 'Not Found in Item File') AS Dept_NBR,
    COALESCE(ItemFile.repl_group_nbr, 'Not Found in Item File') AS REPL_Group_NBR,
    COALESCE(No_Check, 'Error') AS UPC_No_Check,
    CASE
    WHEN qbmasteritems.UPC = No_Check_Results.No_Check THEN 'Found'
    ELSE 'Not Found in QB'
    END AS Found_In_QB,
    COALESCE(qbmasteritems.Item_Validation_Status, 'Not Found in QB') AS Item_Validation_Status
    FROM
    (SELECT eGTIN FROM Receiving_GM GROUP BY eGTIN) AS Receiving_GM
    LEFT JOIN ItemFile ON Receiving_GM.eGTIN = ItemFile.gtin
    LEFT JOIN (SELECT gtin, COUNT(*) AS Total_RFID FROM TotalItems GROUP BY gtin) AS TotalItemsCount ON TotalItemsCount.gtin = Receiving_GM.eGTIN
    LEFT JOIN (SELECT eGTIN, SUM(Transaction_QTY) AS Receiving_Total FROM receiving_gm GROUP BY eGTIN) AS Receiving_GM_Sum ON Receiving_GM_Sum.eGTIN = Receiving_GM.eGTIN
    LEFT JOIN (SELECT eGTIN, SUM(Transaction_QTY) AS Sum_Transactions_Total FROM transactions_gm GROUP BY eGTIN) AS Transactions_GM_Sum ON Transactions_GM_Sum.eGTIN = Receiving_GM.eGTIN
    LEFT JOIN (SELECT UPC, COUNT(*) AS New_EPC_Total FROM newepcs GROUP BY UPC) AS New_EPCs_Count ON New_EPCs_Count.UPC = Receiving_GM.eGTIN
    LEFT JOIN (SELECT eGTIN, MAX(Aggregate_Qty) AS Agg_Qty FROM Receiving_GM GROUP BY eGTIN) AS Aggregate_Qty_Results ON Aggregate_Qty_Results.eGTIN = Receiving_GM.eGTIN
    LEFT JOIN (SELECT DISTINCT eGTIN AS eGTIN, UPC_No_Check AS No_Check FROM UPC_No_Check_Digit) AS No_Check_Results ON No_Check_Results.eGTIN = receiving_gm.eGTIN
    LEFT JOIN qbmasteritems ON No_Check_Results.No_Check = qbmasteritems.UPC;
    """
    cursor.execute(stmt1)
    conn.commit()


def export_receiving_overview_xlsx():
    print("Gathering Receiving Overview table for export...")
    rec_over = sql.read_sql('SELECT * FROM receivingsystem.receiving_overview', conn)
    conn.close()
    return rec_over


def generate_report():
    try:
        if submit_info():
            print("Successfully submitted. Starting Receiving Report Generation...")
            connect_to_mysql()
            import_receiving_sql()
            import_new_epcs_sql()
            import_active_epcs_sql()
            import_qb_sql()
            import_item_file_sql()
            create_upc_drop_sql()
            create_total_items_sql()
            create_transactions_gm_sql()
            create_receiving_gm_sql()
            create_upc_no_check_sql()
            create_receiving_overview_sql()
            print("Report Generated for Store {}.".format(store_num))
            # print("Quitting Application...")
            # app.quit()
        else:
            print("\n------------------------------------------------------------------------"
                  "\n:: ERROR :: Invalid inputs! Please enter valid inputs before submitting!"
                  "\n------------------------------------------------------------------------")
    except Exception as e:
        print(":: ERROR :: Something went wrong! Please restart the application!")
        print(e)


'''Input Verification'''


def store_number_verification():
    global store_num
    store_num = store_entry.get()
    try:
        int(store_num)
        print("Store Number: {}".format(store_num))
        return True
    except:
        print(":: ERROR :: Store Num is not an int!")
        return False


def date_verification():
    global date
    date = date_entry.get()
    try:
        if date == "":
            return False
        date_list = date.split(".")
        if len(date_list[0]) == 4 and isinstance(int(date_list[0]), int):
            if len(date_list[1]) == 2 and isinstance(int(date_list[1]), int):
                if len(date_list[2]) == 2 and isinstance(int(date_list[2]), int):
                    print("Date: {}".format(date))
                    return True
    except:
        print(":: ERROR :: Date input is not valid!")
        return False


def receiving_path_verification():
    global receiving_data_path
    try:
        if receiving_data_path == "":
            return False
        else:
            return True
    except:
        print(":: ERROR :: Receiving Data file path has not been specified!")
        return False


def new_epcs_path_verification():
    global new_epcs_path
    try:
        if new_epcs_path == "":
            return False
        else:
            return True
    except:
        print(":: ERROR :: New EPCs file path has not been specified!")
        return False


def qb_path_verification():
    global qb_master_path
    try:
        if qb_master_path == "":
            return False
        else:
            return True
    except:
        print(":: ERROR :: QB Master Items file path has not been specified!")
        return False


def item_file_path_verification():
    global item_file_path
    try:
        if item_file_path == "":
            return False
        else:
            return True
    except:
        print(":: ERROR :: Item File path has not been specified!")
        return False


class InterfaceCreation:

    def __init__(self, root, w, h):
        self.root = root
        self.width = w
        self.height = h
        self.store_list = []
        self.store_num = None
        self.date_input = None
        self.folder_created = False

    customtkinter.set_appearance_mode("Dark")
    customtkinter.set_default_color_theme("dark-blue")
    global app
    app = customtkinter.CTk()
    app.title("Receiving Report System")
    app.geometry("800x600")

    '''
    Frame Creation
    '''
    main_frame = customtkinter.CTkFrame(master=app, fg_color="transparent")
    top_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    top_frame.configure(height=40)
    left_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    right_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    middle_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    bottom_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")

    main_frame.pack(fill="both", expand=True)
    top_frame.pack(side=TOP, fill="x")
    bottom_frame.pack(side=BOTTOM, fill="x")
    middle_frame.pack(side=BOTTOM, fill="x")
    left_frame.pack(side=LEFT, fill="both", expand=True)
    right_frame.pack(side=RIGHT, fill="both", expand=True)

    # file_name_text_box = customtkinter.CTkTextbox(master=bottom_frame)

    '''
    Store and Date Entry Creation
    '''
    global store_entry
    store_entry = customtkinter.CTkEntry(master=left_frame, placeholder_text="Store #")

    global date_entry
    date_entry = customtkinter.CTkEntry(master=right_frame, placeholder_text="Date (YYYY.MM.DD)")

    store_entry.pack(padx=30, pady=50)
    date_entry.pack(padx=30, pady=50)

    '''
    Button Creation
    '''
    receiving_data_button = customtkinter.CTkButton(master=left_frame, text="Receiving Data (.csv)",
                                                    command=import_receiving_data)
    new_epc_button = customtkinter.CTkButton(master=left_frame, text="New EPCs (.xlsx)", command=import_new_epcs)
    qb_master_items_button = customtkinter.CTkButton(master=left_frame, text="QB Master Items (.xlsb)",
                                                     command=import_qb_master_items)
    item_file_button = customtkinter.CTkButton(master=right_frame, text="Item File (.csv)", command=import_item_file)
    settings_button = customtkinter.CTkButton(master=top_frame, text="Settings", width=75, height=30,
                                              command=open_settings)
    quit_button = customtkinter.CTkButton(master=top_frame, text="Quit", width=75, height=30, command=quit_app)
    submit_button = customtkinter.CTkButton(master=middle_frame, text="Submit", command=generate_report)

    receiving_data_button.pack(pady=5)
    new_epc_button.pack(pady=5)
    qb_master_items_button.pack(pady=5)
    item_file_button.pack(pady=5)
    settings_button.pack(anchor=NE, padx=15, pady=10)
    quit_button.pack(anchor=NE, padx=15, pady=0)
    submit_button.pack(padx=100, pady=10)

    app.mainloop()
