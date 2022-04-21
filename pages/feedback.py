# Need options to upload one to 3 files so i can do if none options
# merge datasets on sch line item doing whiochever merge need it to probabily be efficent
# highlight yellow if status is under review with CSAM or action sdm is Not order created on tool Not processed on LTSI tool
import pandas as pd
from datetime import datetime, timedelta
import io
import streamlit as st


# ideas
# join the all the columns then merge ?
def app():
    # st.set_page_config(page_title='LTSI Feedback Form')

    st.write("""

        # LTSI Feedback 
        ### Instructions: \n
        - Used when SDM feedback is in seperate files 
        - Upload excel feedback files (up to 3)
        - Upload Open Orders File """)
    st.write("## Upload 1 to 3 Feedback Files")
    feedback1 = st.file_uploader("Upload Feedback File 1", type="xlsx")
    feedback2 = st.file_uploader("Upload Feedback File 2", type="xlsx")
    feedback3 = st.file_uploader("Upload Feedback File 3", type="xlsx")
    st.write("## Upload Open Orders File")
    open_orders = st.file_uploader("Upload Open Order File if feedback does not contain all open order rows",
                                   type="xlsx")
    if st.button("Create Feedback"):
        def download_file(merged):
            # Writing df to Excel Sheet
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:

                # Write each dataframe to a different worksheet.
                # data["Date"] = pd.to_datetime(data["Date"])

                # pd.to_datetime('date')
                merged.to_excel(writer, sheet_name='Sheet1', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                formatdict = {'num_format': 'dd/mm/yyyy'}
                fmt = workbook.add_format(formatdict)
                worksheet.set_column('K:K', None, fmt)
                worksheet.set_column('L:L', None, fmt)
                # Light yellow fill with dark yellow text.
                number_rows = len(merged.index) + 1
                yellow_format = workbook.add_format({'bg_color': '#FFEB9C'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AG2="Under Review with CSAM"',
                                              'format': yellow_format})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AG2="Under Review with C-SAM"',
                                              'format': yellow_format})
                red_format = workbook.add_format({'bg_color': '#ffc7ce'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AG2="Blocked"',
                                              'format': red_format})

                green_format = workbook.add_format({'bg_color': '#c6efce'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AG2="Shippable"',
                                              'format': green_format})
                # COL MIGHT BE AH
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Under Review with CSAM"',
                                              'format': yellow_format})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Under Review with C-SAM"',
                                              'format': yellow_format})
                red_format = workbook.add_format({'bg_color': '#ffc7ce'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Blocked"',
                                              'format': red_format})

                green_format = workbook.add_format({'bg_color': '#c6efce'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Shippable"',
                                              'format': green_format})
                for column in merged:
                    column_width = max(merged[column].astype(str).map(len).max(), len(column))
                    col_idx = merged.columns.get_loc(column)
                    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
                    worksheet.autofilter(0, 0, merged.shape[0], merged.shape[1])
                worksheet.set_column(11, 12, 20)
                worksheet.set_column(12, 13, 20)
                worksheet.set_column(13, 14, 20)
                header_format = workbook.add_format({'bold': True,
                                                     'bottom': 2,
                                                     'bg_color': '#0AB2F7'})

                # Write the column headers with the defined format.
                for col_num, value in enumerate(merged.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                my_format = workbook.add_format()
                my_format.set_align('left')

                worksheet.set_column('N:N', None, my_format)
                writer.save()
                today = datetime.today()
                d1 = today.strftime("%d/%m/%Y")
                st.write("Download Completed File:")
                st.download_button(
                    label="Download Excel worksheets",
                    data=buffer,
                    file_name="LTSI_file_" + d1 + ".xlsx",
                    mime="application/vnd.ms-excel"
                )



        def columns_to_keep():
            cols = ['sales_org', 'country', 'cust_num', 'customer_name', 'sales_dis', 'rtm', 'sd_line_item',
                    'order_method', 'del_blk', 'cust_req_date', 'ord_entry_date',
                    'cust_po_num', 'ship_num', 'ship_cust', 'ship_city', 'plant',
                    'material_num', 'brand', 'lob', 'project_code', 'material_desc',
                    'mpn_desc', 'ord_qty', 'shpd_qty', 'delivery_qty', 'remaining_qty',
                    'delivery_priority', 'opt_delivery_qt', 'rem_mod_opt_qt',
                    'sch_line_blocked_for_delv']
            return cols


        def old_feedback_getter(df):
            cols = [8]
            col_count = 37
            if df.shape[1] >= 39:
                while col_count < df.shape[1]:
                    cols.append(col_count)
                    col_count += 1

            return df.iloc[:, cols]

        def new_feedback_getter(df):
            return df.iloc[:, [8, 34, 35, 36]]

        def open_new_feedback_merge(open, new_feedback):
            return open.merge(new_feedback, how="left", on="Sales Order and Line Item")

        def case2(feedback, open_orders):
            feed1 = pd.read_excel(feedback, sheet_name=0, engine="openpyxl")
            openOrders = pd.read_excel(open_orders, sheet_name=0, engine="openpyxl")
            old_feedback = old_feedback_getter(feed1)
            new_feedback = new_feedback_getter(feed1)
            open = openOrders.iloc[:, :34]
            combined_feedback = open.merge(new_feedback, how="left", on="Sales Order and Line Item")
            final = combined_feedback.merge(old_feedback, how="left", on="Sales Order and Line Item")
            cols = columns_to_keep()
            final.drop_duplicates(subset=cols, keep='first', inplace=True)
            download_file(final)

        def case3(feedback1, feedback2, open_orders):
            feed1 = pd.read_excel(feedback1, sheet_name=0, engine="openpyxl")
            feed2 = pd.read_excel(feedback2, sheet_name=0, engine="openpyxl")
            openOrders = pd.read_excel(open_orders, sheet_name=0, engine="openpyxl")
            old_feedback1 = old_feedback_getter(feed1)
            new_feedback1 = new_feedback_getter(feed1)
            old_feedback2 = old_feedback_getter(feed2)
            new_feedback2 = new_feedback_getter(feed2)
            open = openOrders.iloc[:, :34]
            joined_new_feedback = pd.concat([new_feedback1, new_feedback2], ignore_index=True)
            joined_old_feedback = pd.concat([old_feedback1, old_feedback2], ignore_index=True)
            combined_feedback = open.merge(joined_new_feedback, how="left", on="Sales Order and Line Item")
            final = combined_feedback.merge(joined_old_feedback, how="left", on="Sales Order and Line Item")
            cols = columns_to_keep()
            final.drop_duplicates(subset=cols, keep='first', inplace=True)
            download_file(final)

        def case4(feedback1, feedback2, feedback3, open_orders):
            feed1 = pd.read_excel(feedback1, sheet_name=0, engine="openpyxl")
            feed2 = pd.read_excel(feedback2, sheet_name=0, engine="openpyxl")
            feed3 = pd.read_excel(feedback3, sheet_name=0, engine="openpyxl")
            openOrders = pd.read_excel(open_orders, sheet_name=0, engine="openpyxl")
            old_feedback1 = old_feedback_getter(feed1)
            new_feedback1 = new_feedback_getter(feed1)
            old_feedback2 = old_feedback_getter(feed2)
            new_feedback2 = new_feedback_getter(feed2)
            old_feedback3 = old_feedback_getter(feed3)
            new_feedback3 = new_feedback_getter(feed3)
            open = openOrders.iloc[:, :34]
            joined_new_feedback = pd.concat([new_feedback1, new_feedback2, new_feedback3], ignore_index=True)
            joined_old_feedback = pd.concat([old_feedback1, old_feedback2, old_feedback3], ignore_index=True)
            combined_feedback = open.merge(joined_new_feedback, how="left", on="Sales Order and Line Item")
            final = combined_feedback.merge(joined_old_feedback, how="left", on="Sales Order and Line Item")
            cols = columns_to_keep()
            final.drop_duplicates(subset=cols, keep='first', inplace=True)
            download_file(final)

        def case5(feedback1, feedback2):
            feed1 = pd.read_excel(feedback1, sheet_name=0, engine="openpyxl")
            feed2 = pd.read_excel(feedback2, sheet_name=0, engine="openpyxl")
            open = feed1.iloc[:, :34]
            old_feedback = old_feedback_getter(feed1)
            # drop na
            new_feedback1 = new_feedback_getter(feed1)
            new_feedback2 = new_feedback_getter(feed2)
            new_feedback1 = new_feedback1[new_feedback1.iloc[:, 1].notna()]
            new_feedback2 = new_feedback2[new_feedback2.iloc[:, 1].notna()]
            joined_new_feedback = pd.concat([new_feedback1, new_feedback2], ignore_index=True)
            combined_feedback = open.merge(joined_new_feedback, how="left", on="Sales Order and Line Item")
            final = combined_feedback.merge(old_feedback, how="left", on="Sales Order and Line Item")

            cols = columns_to_keep()

            final.drop_duplicates(subset=cols, keep='first', inplace=True)
            download_file(final)

        def case6(feedback1, feedback2, feedback3):
            feed1 = pd.read_excel(feedback1, sheet_name=0, engine="openpyxl")
            feed2 = pd.read_excel(feedback2, sheet_name=0, engine="openpyxl")
            feed3 = pd.read_excel(feedback3, sheet_name=0, engine="openpyxl")
            open = feed1.iloc[:, :34]
            old_feedback = old_feedback_getter(feed1)
            # drop na
            new_feedback1 = new_feedback_getter(feed1)
            new_feedback1 = new_feedback1[new_feedback1.iloc[:, 1].notna()]
            new_feedback2 = new_feedback_getter(feed2)
            new_feedback2 = new_feedback2[new_feedback2.iloc[:, 1].notna()]
            new_feedback3 = new_feedback_getter(feed3)
            new_feedback3 = new_feedback3[new_feedback3.iloc[:, 1].notna()]
            joined_new_feedback = pd.concat([new_feedback1, new_feedback2, new_feedback3], ignore_index=True)
            combined_feedback = open.merge(joined_new_feedback, how="left", on="Sales Order and Line Item")
            final = combined_feedback.merge(old_feedback, how="left", on="Sales Order and Line Item")
            cols = columns_to_keep()
            final.drop_duplicates(subset=cols, keep='first', inplace=True)
            download_file(final)

        # Case 1 feedback + no open (has all rows)
        if feedback1 is not None and feedback2 is None and feedback3 is None and open_orders is None:
            st.error("File already complete, no need to upload")
        if feedback1 is None and feedback2 is not None and feedback3 is None and open_orders is None:
            st.error("File already complete, no need to upload")
        if feedback1 is None and feedback2 is None and feedback3 is not None and open_orders is None:
            st.error("File already complete, no need to upload")

        # Case 2 one feedback + open -> combine
        if feedback1 is not None and feedback2 is None and feedback3 is None and open_orders is not None:
            case2(feedback1, open_orders)
        if feedback1 is None and feedback2 is not None and feedback3 is None and open_orders is not None:
            case2(feedback2, open_orders)
        if feedback1 is None and feedback2 is None and feedback3 is not None and open_orders is not None:
            case2(feedback3, open_orders)

        # Case 3 two feedback + open -> combine
        if feedback1 is not None and feedback2 is not None and feedback3 is None and open_orders is not None:
            case3(feedback1, feedback2, open_orders)
        if feedback1 is not None and feedback2 is None and feedback3 is not None and open_orders is not None:
            case3(feedback1, feedback3, open_orders)
        if feedback1 is None and feedback2 is not None and feedback3 is not None and open_orders is not None:
            case3(feedback2, feedback3, open_orders)

        # Case 4 3 feedback + open orders
        if feedback1 is not None and feedback2 is not None and feedback3 is not None and open_orders is not None:
            case4(feedback1, feedback2, feedback3, open_orders)

        # Case 5 2 feedbacks and no open orders
        if feedback1 is not None and feedback2 is not None and feedback3 is None and open_orders is None:
            case5(feedback1, feedback2)
        if feedback1 is not None and feedback2 is None and feedback3 is not None and open_orders is None:
            case5(feedback1, feedback3)
        if feedback1 is None and feedback2 is not None and feedback3 is not None and open_orders is None:
            case5(feedback2, feedback3)

        # Case 6 3 feedbacks and no Open
        if feedback1 is not None and feedback2 is not None and feedback3 is not None and open_orders is None:
            case6(feedback1, feedback2, feedback3)

        # Case 7 0 feedback and Open
        if feedback1 is None and feedback2 is None and feedback3 is None and open_orders is not None:
            st.error("Error: No Feedback uploaded. \n\n"
                     "Open Order file up to date")

        # Case 8 no files uploaded
        if feedback1 is None and feedback2 is None and feedback3 is None and open_orders is None:
            st.error("Error: No Feedback/Files uploaded.")
