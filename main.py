import streamlit as st

import pandas as pd
import numpy as np




st.set_page_config(page_title='LTSI Open Orders')

st.write("""

# LTSI Tool 

### Contact me if issues arise:
Slack: @Cameron Looney \n
email: cameron_j_looney@apple.com""")

aux = st.file_uploader("Upload Auxillary File", type="xlsx")
master = st.file_uploader("Upload Row File File", type="xlsx")

if aux is not None and master is not None:
    vlookup = pd.read_excel(aux, sheet_name=0,engine="openpyxl")
    previous = pd.read_excel(aux, sheet_name=1,engine="openpyxl")
    dropdown = pd.read_excel(aux, sheet_name=2,engine="openpyxl")
    TF = pd.read_excel(aux, sheet_name=3,engine="openpyxl")
    master = pd.read_excel(master, sheet_name = 0,engine = "openpyxl")


    vlookup.rename(columns={'MPN': 'material_num'}, inplace=True)
    master = master.merge(vlookup, on= 'material_num', how='left')
    print(list(master.columns))

    rows = master[master['Date']>= master['ord_entry_date']].index.to_list()
    master = master.drop(rows).reset_index()





    # TO FAR AHEAD DATES

    from datetime import datetime, timedelta


    today = datetime.today()
    three_weeks =  today + timedelta(weeks=12)
    rows = master[master['cust_req_date']> three_weeks].index.to_list()
    master = master.drop(rows).reset_index()


    # DROP ROWS WHERE REMAINING = 0
    x = list(master['remaining_qty'].unique())

    master = master.loc[master['remaining_qty'] != 0]


    # DROP UNNEEDED COLUMNS
    cols =['sales_org', 'country', 'cust_num', 'customer_name', 'sales_dis', 'rtm',
           'sales_ord', 'sd_line_item',
           'order_method', 'del_blk', 'cust_req_date', 'ord_entry_date',
           'cust_po_num', 'ship_num', 'ship_cust', 'ship_city', 'plant',
           'material_num', 'brand', 'lob', 'project_code', 'material_desc',
           'mpn_desc', 'ord_qty', 'shpd_qty', 'delivery_qty', 'remaining_qty',
           'delivery_priority', 'opt_delivery_qt', 'rem_mod_opt_qt',
           'sch_line_blocked_for_delv', ]





    make = ['Sales Order and Line Item','Status (SS)', 'Valid in LTSI Tool',
           'Action (SDM)', 'Comments (SDM)', 'Estimated DN Date',
           'Action (SDM) 15 March', 'Comments (SDM) 15 March',
           'Estimated DN Date 15 March']

    # APPLY REDUCTION
    reduced = master[cols]

    # CREATE AND FILL THE VALID IN LTSI COL
    # THIS IS BETTER THAN OTHER MERGE, PREVENTS MAKING COPIES
    reduced.rename(columns={'sales_ord': 'salesOrderNum'},inplace=True)
    reduced['g']=reduced.groupby('salesOrderNum').cumcount()
    TF['g']=TF.groupby('salesOrderNum').cumcount()
    merged = reduced.merge(TF,how='left').drop('g',1)
    idx = 8
    new_col = merged['salesOrderNum'].astype(str) + merged['sd_line_item'].astype(str)

    merged.insert(loc=idx, column='Sales Order and Line Item', value=new_col)
    merged['Sales Order and Line Item'] = merged['Sales Order and Line Item'].astype(int)


    merged.rename(columns={'Unnamed: 1': 'Valid in LTSI Tool'},inplace=True)
    merged["Valid in LTSI Tool"].fillna("FALSE", inplace=True)
    mask = merged.applymap(type) != bool
    d = {True: 'TRUE', False: 'FALSE'}
    merged = merged.where(mask, merged.replace(d))
    merged["cust_req_date"] = merged["cust_req_date"].dt.strftime('%d/%m/%Y')
    merged["ord_entry_date"] = merged["ord_entry_date"].dt.strftime('%d/%m/%Y')

    conditions = [merged['order_method'] == "Manual SAP",
                  merged['delivery_priority'] == 13,
                  merged["Valid in LTSI Tool"] == "TRUE",
                  ~merged["del_blk"].isnull(),
                  ~merged["sch_line_blocked_for_delv"].isnull()]
    outputs = ["Shippable", "Shippable", "Shippable", "Blocked","Blocked"]
    # CONCAT ID AND LINE ORDER AND ADD
    #  INDEX IS 8

    res = np.select(conditions, outputs, 'Under Review by C-SAM')
    res = pd.Series(res)
    merged['Status (SS)'] = res
    feedback = previous.drop('Status (SS)', 1)
    merged['g'] = merged.groupby('Sales Order and Line Item').cumcount()
    feedback['g'] = feedback.groupby('Sales Order and Line Item').cumcount()
    merged = merged.merge(feedback, how='left').drop('g', 1)


    import io

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
        st.download_button(
            label="Download Excel worksheets",
            data=buffer,
            file_name="LTSI_file_" + d1 + ".xlsx",
            mime="application/vnd.ms-excel"
        )


