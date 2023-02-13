import customtkinter
import InterfaceCreation
import os
import pandas as pd
import xlsxwriter

interface = InterfaceCreation.InterfaceCreation(customtkinter.CTk, 800, 650)

try:
    if InterfaceCreation.store_num is not None and InterfaceCreation.date is not None:
        report_file_name = "Store{}Receiving_Report{}.xlsx".format(str(InterfaceCreation.store_num), str(InterfaceCreation.date))
        date = InterfaceCreation.date
        str(date)
        date = date.replace(".", "")
        folder_path = str(os.path.join(os.path.expanduser("~"), "Desktop/Receiving_Reports_{}".format(date)))
        os.mkdir(folder_path)
        path = os.path.join(os.path.expanduser("~"), "Desktop/Receiving_Reports_{}".format(date),
                            report_file_name)
        str(path)

        global writer
        writer = pd.ExcelWriter(path, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_numbers': True}})

        receiving_overview_sheet_name = "Store {} Receiving Overview".format(InterfaceCreation.store_num)
        str(receiving_overview_sheet_name)

        print("Exporting Receiving Overview...")
        InterfaceCreation.export_receiving_overview_xlsx().to_excel(writer, receiving_overview_sheet_name, startrow=0, startcol=0, index=False)
        writer.save()
    InterfaceCreation.conn.close()
    if InterfaceCreation.conn:
        InterfaceCreation.conn.close()
        print("MySQL connection is closed.")
except AttributeError as ae:
    pass
except Exception as e:
    print("Something went wrong...")
    print(":: ERROR :: {}".format(e))

# raise SystemExit(0)
exit(1)