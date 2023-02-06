import customtkinter
import InterfaceCreation
interfaceCreation = InterfaceCreation.InterfaceCreation(customtkinter.CTk, 800, 600)

import os
import pandas as pd
import InterfaceCreation

interface = InterfaceCreation.InterfaceCreation(customtkinter.CTk, 800, 650)
# print(interface.store_list)


report_file_name = "WeeklyReport{0}.xlsx".format(interface.date_input)
path = os.path.join(os.path.expanduser("~"), "Desktop/TrackingReports_{0}".format(interface.date_input),
                    report_file_name)
str(path)

global writer
writer = pd.ExcelWriter(path, engine='xlsxwriter')

combined_matching_sheet_name = "Combined Matching"
str(combined_matching_sheet_name)
interface.store_list[-1].get_combined().to_excel(writer, combined_matching_sheet_name, startrow=0, startcol=0, index=False)

combined_repl_sheet_name = "REPL_NBR_Breakdown Total"
str(combined_repl_sheet_name)
interface.store_list[-1].get_combined_repl().to_excel(writer, combined_repl_sheet_name, startrow=0, startcol=0, index=False)

for store in interface.store_list:
    matching_sheet_name = "Matching {}".format(store.store_num)
    str(matching_sheet_name)
    total_items_sheet_name = "Total Items {}".format(store.store_num)
    str(total_items_sheet_name)
    expected_items_sheet_name = "Expected Items {}".format(store.store_num)
    str(expected_items_sheet_name)
    repl_group_nbr_sheet_name = "REPL Breakdown {}".format(store.store_num)
    str(repl_group_nbr_sheet_name)

    store.get_matching().to_excel(writer, matching_sheet_name, startrow=0, startcol=0, index=False)
    store.get_total_items().to_excel(writer, total_items_sheet_name, startrow=0, startcol=0, index=False)
    store.get_expected().to_excel(writer, expected_items_sheet_name, startrow=0, startcol=0, index=False)
    store.get_repl_nbr().to_excel(writer, repl_group_nbr_sheet_name, startrow=0, startcol=0, index=False)


writer.save()

raise SystemExit(0)