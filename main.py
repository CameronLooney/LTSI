#DISCLAIMER: The code is in no way efficient, it is designed to merely complete the task. It is also not finished as
# some hard coded logic needs to be changed. It is simply the quick and dirty solution to the problem at hand
import streamlit as st
import pandas as pd
import numpy as np



# This will be a global variable to keep track of any problems along the process. This is needed to ensure the file
# will not be downloaded unless it has no errors.
error_count = 0
# Title of the Application
st.set_page_config(page_title='LTSI Open Orders')

st.write("""

# LTSI Tool 
## Instructions\n 
### For the first upload please make sure you have the following:\n
Sheet 1: vlookup \n 
Sheet 2: Previous \n 
Sheet 3: Dropdown Menu \n  
Sheet 4: LTSI tool True \n \n 

### For your second upload please upload your raw download for the day \n \n \n

### Contact me if issues arise:
Slack: @Cameron Looney \n
email: cameron_j_looney@apple.com""")
# Need to uploads to generate Open Orders, one is a helper file which is used for computation and feedback. The master is
# the file downloaded from FrontEnd each day
aux = st.file_uploader("Upload Auxiliary File", type="xlsx")
master = st.file_uploader("Upload Raw File", type="xlsx")
# Button to start the process
if st.button("Generate LTSI File"):
    # These ensure that two files have been uploaded
    if aux is None:
        st.error("ERROR: Please upload a viable auxiliary file to continue.")
    if master is None:
        st.error("ERROR: Please upload your raw LTSI download file to continue.")
    # If both files are uploaded we can begin the computation
    if aux is not None and master is not None:
        error_count = 0
        sheetNumCheck = pd.ExcelFile(aux)
        # check that the helper file has the required sheets
        if len(sheetNumCheck.sheet_names) != 4:
            st.error("ERROR: Missing sheet from helper file.\n"
                     "Please ensure you have 4 sheets:\n"
                     "1. vlookup\n"
                     "2. previous\n "
                     "3. dropdown menu\n"
                     "4. LTSI True list\n"
                     "Please clear your uploads and try again")


        else:

            vlookup = pd.read_excel(aux, sheet_name=0,engine="openpyxl")
            vlookup_col_check = ["MPN","Date"]
            # check contents of the vlookup sheet
            def vlookup_checker(x, to_check):
                if not set(to_check).issubset(set(x.columns)):
                    global error_count
                    error_count +=1
                    return False
                else:
                    return True
            if not vlookup_checker(vlookup,vlookup_col_check):
                st.error(f"{' and '.join(set(vlookup_col_check).difference(vlookup.columns))} column not available in the dataframe\n"
                         f"Please fix and try again")

            # load all the excel sheets
            previous = pd.read_excel(aux, sheet_name=1,engine="openpyxl")
            dropdown = pd.read_excel(aux, sheet_name=2,engine="openpyxl")
            TF = pd.read_excel(aux, sheet_name=3,engine="openpyxl")
            master = pd.read_excel(master, sheet_name = 0,engine = "openpyxl")

            # rename to enable a join
            vlookup.rename(columns={'MPN': 'material_num'}, inplace=True)
            # vlookup date data is imcomplete so need to convert and add a place holder date
            # without it we lose viable rows
            vlookup['Date'] = vlookup['Date'].fillna("01.01.90")
            vlookup['Date'] = pd.to_datetime(vlookup.Date, dayfirst=True)
            vlookup['Date'] = [x.date() for x in vlookup.Date]
            vlookup['Date'] = pd.to_datetime(vlookup.Date)
            # perform vlookup type join with the raw download and vlookup
            master = master.merge(vlookup, on= 'material_num', how='left')
            # LOGIC STEP
            # if order placed before it was LTSI delete
            rows = master[master['Date'] > master['ord_entry_date']].index.to_list()

            master = master.drop(rows).reset_index()
            master['sch_line_blocked_for_delv'] = master['sch_line_blocked_for_delv'].astype(str)
            master['sch_line_blocked_for_delv'] = master['sch_line_blocked_for_delv'].replace("nan", "")

            master['del_blk'] = master['del_blk'].astype(str)
            master['del_blk'] = master['del_blk'].replace("nan", "")
            # LOGIC STEP
            # Drop rows if they are 6 months old and still have a 94 block. This is more of a dodgy fix. Risk of deleting viable
            # rows however the risk is low. Needed as we overshoot open Orders by 500+ without it.
            from datetime import datetime, timedelta
            six_months = datetime.now() - timedelta(188)
            rows_94 = master[
                (master['ord_entry_date'] < six_months) & (master["sch_line_blocked_for_delv"] == 94)].index.to_list()

            # drop these rows
            master = master.drop(rows_94).reset_index(drop=True)
            # LOGIC STEP
            # Again this step is more of my interpretation in order to get as accurate as possible. This removes 1 year+ Orders
            # thinking is if it was going to be locked in, it would of happened by now
            twelve_months = datetime.now() - timedelta(365)
            rows_old = master[(master['ord_entry_date'] < twelve_months)].index.to_list()


            master = master.drop(rows_old).reset_index(drop=True)
            # LOGIC STEP
            # Drop any rows where we have no quantity left.
            master = master.loc[master['remaining_qty'] != 0]
            # This is hard coded logic, There was no discernible reason they were dropped in Macro. However on 3 concurrent test days they
            # weren't present so for now we will hardcode to drop them from our df
            country2021drop = master[(master['ord_entry_date'].dt.year == 2021) & (master['country'].isin(
                ['Germany', 'Spain', "Turkey", "Belgium / Luxembourg", "Switzerland"]))].index.to_list()
            master = master.drop(country2021drop).reset_index(drop=True)



            from datetime import datetime, timedelta

            # LOGIC STEP
            # Drop rows where the customer doesnt want it for 3 months. Again this was an interpretation to match Macro
            today = datetime.today()
            twelve_weeks =  today + timedelta(weeks=12)
            rows = master[master['cust_req_date']> twelve_weeks].index.to_list()
            master = master.drop(rows).reset_index(drop=True)



            # These are the columns that wont be needed in our final df so we can drop
            cols =['sales_org', 'country', 'cust_num', 'customer_name', 'sales_dis', 'rtm',
                   'sales_ord', 'sd_line_item',
                   'order_method', 'del_blk', 'cust_req_date', 'ord_entry_date',
                   'cust_po_num', 'ship_num', 'ship_cust', 'ship_city', 'plant',
                   'material_num', 'brand', 'lob', 'project_code', 'material_desc',
                   'mpn_desc', 'ord_qty', 'shpd_qty', 'delivery_qty', 'remaining_qty',
                   'delivery_priority', 'opt_delivery_qt', 'rem_mod_opt_qt',
                   'sch_line_blocked_for_delv', ]







            # APPLY REDUCTION
            reduced = master[cols]
            # need to convert type as the 95 block was being converted to a date when introducd back into excel
            reduced['del_blk'] = np.where(pd.isnull(reduced['del_blk']), reduced['del_blk'], reduced['del_blk'].astype(str))



            # Merging main df with the True False LTSI df based on sales ord num
            reduced.rename(columns={'sales_ord': 'salesOrderNum'},inplace=True)
            reduced['holder']=reduced.groupby('salesOrderNum').cumcount()
            TF['holder']=TF.groupby('salesOrderNum').cumcount()
            merged = reduced.merge(TF,how='left').drop('holder',1)
            idx = 8
            new_col = merged['salesOrderNum'].astype(str) + merged['sd_line_item'].astype(str)
            # insert merged column to act as unique key
            merged.insert(loc=idx, column='Sales Order and Line Item', value=new_col)
            merged['Sales Order and Line Item'] = merged['Sales Order and Line Item'].astype(int)

            # CREATE AND FILL THE VALID IN LTSI COLUMN
            # THIS IS BETTER THAN OTHER MERGE, PREVENTS MAKING COPIES
            merged.rename(columns={'Unnamed: 1': 'Valid in LTSI Tool'},inplace=True)
            merged["Valid in LTSI Tool"].fillna("FALSE", inplace=True)
            mask = merged.applymap(type) != bool
            d = {True: 'TRUE', False: 'FALSE'}
            merged = merged.where(mask, merged.replace(d))

            # conditions for order Status
            # Logic is as followed
            # Manual SAP/Prority 13/ Valid in LTSI =Ship
            # block in either block column is not shippable
            conditions = [merged['order_method'] == "Manual SAP",
                          merged['delivery_priority'] == 13,
                          merged["Valid in LTSI Tool"] == "TRUE",
                          ~merged["del_blk"].isnull(),
                          ~merged["sch_line_blocked_for_delv"].isnull()]
            outputs = ["Shippable", "Shippable", "Shippable", "Blocked","Blocked"]
            # CONCAT ID AND LINE ORDER AND ADD
            #  IN   DEX IS 8
            res = np.select(conditions, outputs, "Under Review by C-SAM")
            res = pd.Series(res)
            merged['Status (SS)'] = res
            feedback = previous.drop('Status (SS)', 1)
            merged['g'] = merged.groupby('Sales Order and Line Item').cumcount()
            feedback['g'] = feedback.groupby('Sales Order and Line Item').cumcount()
            merged = merged.merge(feedback, how='left').drop('g', 1)

            merged['g'] = merged.groupby('Sales Order and Line Item').cumcount()
            previous['g'] = previous.groupby('Sales Order and Line Item').cumcount()
            merged = merged.merge(previous, how='left').drop('g', 1)


            import io
            # Writing df to Excel Sheet
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Write each dataframe to a different worksheet.
                #data["Date"] = pd.to_datetime(data["Date"])

                #pd.to_datetime('date')
                merged.to_excel(writer, sheet_name='Sheet1',index = False)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                formatdict = {'num_format': 'dd/mm/yyyy'}
                fmt = workbook.add_format(formatdict)
                worksheet.set_column('K:K', None, fmt)
                worksheet.set_column('L:L', None, fmt)
                for column in merged:
                    column_width = max(merged[column].astype(str).map(len).max(), len(column))
                    col_idx = merged.columns.get_loc(column)
                    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
                    worksheet.autofilter(0, 0, merged.shape[0], merged.shape[1])



                writer.save()
                d1 = today.strftime("%d/%m/%Y")
                st.write("Download Completed File:")
                if error_count ==0:
                    st.download_button(
                        label="Download Excel worksheets",
                        data=buffer,
                        file_name="LTSI_file_" + d1 + ".xlsx",
                        mime="application/vnd.ms-excel"
                    )
                else:
                    st.error("ERROR: An error was detected. Please try fix file format and try again.")



