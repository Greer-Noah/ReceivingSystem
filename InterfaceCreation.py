from tkinter import *
from tkinter import filedialog
import customtkinter
import os
import mysql.connector
import pandas as pd
from pyxlsb import open_workbook as open_xlsb

global store_num
global date

global receiving_data_path
global new_epcs_path
global qb_master_path
global total_items_path
global item_file_path

global conn
global cursor


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


def import_total_items():
    print("Total Items...")
    pop_up_title = "Select Total Items (.xlsx)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
    global total_items_path
    total_items_path = filename
    print(total_items_path)


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
    exit(1)


def submit_info():
    return_value = True
    '''
    if store_number_verification() and date_verification():
        print("Store Number and Date input accepted...")
    else:
        print("Store Number or Date input(s) are invalid!")
        return_value = False

    if receiving_path_verification():
        print("Valid Receiving Data input...")
    else:
        print("Receiving Data input is invalid or unspecified!")
        return_value = False

    if new_epcs_path_verification():
        print("Valid New EPCs file input...")
    else:
        print("New EPCs file input is invalid or unspecified!")
        return_value = False

    if qb_path_verification():
        print("Valid QB Master Items file input...")
    else:
        print("QB Master Items file input is invalid or unspecified!")
        return_value = False

    if total_items_path_verification():
        print("Valid Total Items file input...")
    else:
        print("Total Items file input is invalid or unspecified!")
        return_value = False

    if item_file_path_verification():
        print("Valid Item File input...")
    else:
        print("Item File input is invalid or unspecified!")
        return_value = False
    '''
    for func in [store_number_verification, date_verification, receiving_path_verification, new_epcs_path_verification,
                 qb_path_verification, total_items_path_verification, item_file_path_verification]:
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
        print(":: ERROR:: Could not import Receiving Data!")
        print(e)


def import_new_epcs_sql():
    stmt = "DROP TABLE IF EXISTS NewEPCs"
    cursor.execute(stmt)
    pass


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
                            "Item_Arrival_Status text, Vendor_Number text, Vendor_Name text, Dept_NBR text, UPC text, " \
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


def generate_report():
    try:
        if submit_info():
            print("Successfully submitted. Starting Receiving Report Generation...")
            connect_to_mysql()
            import_receiving_sql()
            import_new_epcs_sql()
            import_qb_sql()
            import_item_file_sql()
            print("Report Generated.")
            conn.close()
            if conn:
                conn.close()
                print("MySQL connection is closed.")
            print("Quitting Application...")
            exit(1)
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


def total_items_path_verification():
    global total_items_path
    try:
        if total_items_path == "":
            return False
        else:
            return True
    except:
        print(":: ERROR :: Total Items file path has not been specified!")
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

    file_name_text_box = customtkinter.CTkTextbox(master=bottom_frame)

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
    total_items_button = customtkinter.CTkButton(master=right_frame, text="Total Items (.xlsx)",
                                                 command=import_total_items)
    item_file_button = customtkinter.CTkButton(master=right_frame, text="Item File (.csv)", command=import_item_file)
    settings_button = customtkinter.CTkButton(master=top_frame, text="Settings", width=75, height=30,
                                              command=open_settings)
    quit_button = customtkinter.CTkButton(master=top_frame, text="Quit", width=75, height=30, command=quit_app)
    submit_button = customtkinter.CTkButton(master=middle_frame, text="Submit", command=generate_report)

    receiving_data_button.pack(pady=5)
    new_epc_button.pack(pady=5)
    qb_master_items_button.pack(pady=5)
    total_items_button.pack(pady=5)
    item_file_button.pack(pady=5)
    settings_button.pack(anchor=NE, padx=15, pady=10)
    quit_button.pack(anchor=NE, padx=15, pady=0)
    submit_button.pack(padx=100, pady=10)

    app.mainloop()
