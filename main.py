import utils.utils as ut
import utils.clean_data as cd
import utils.format_excel as fe
import time


def main():
    ut.set_logger()

    # Convert the input string to a list
    invoice_filenames = ut.find_workbook_list()
    mismatch_df = None  # Initialize workbook variable

    for invoice_filename in invoice_filenames:
        # Extract invoice sheet names
        invoice_sheet_names = ut.list_visible_sheets_in_workbook(invoice_filename)

        # Copy template and rename the workbook and sheets
        workbook_filename = ut.copy_and_rename_excel(invoice_filename, invoice_sheet_names)

        for invoice_sheet_name in invoice_sheet_names:

            # Import data from excel workbook
            import_df = ut.read_excel_data(invoice_filename, invoice_sheet_name)

            # Intialize class for cleaning data
            clean_df = cd.CleanData(import_df)

            # Calculate T/S
            clean_df.calculate_ts()

            # find empty comment
            empty_comments_df = clean_df.find_empty_comments()

            # Extract times from comment
            format_issue_df = clean_df.extract_times()

            # Check for military time format
            military_time_issue_df = clean_df.check_military_time_format()

            # Calculate time worked
            differing_time_df, acceptable_df = clean_df.calculate_time_worked()

            # Clean and calcualte subtotals for summar sheet
            final_df = clean_df.clean_summary_df()

            # Extract final dataframe
            final_df = clean_df.df

            # Organize dataframes
            dataframes = {
                "Empty": empty_comments_df,
                "Format": format_issue_df,
                "Military": military_time_issue_df,
                "Mismatch": mismatch_df,
                "ConflictingTime": differing_time_df,
                "Acceptable": acceptable_df,
                "Summary": final_df
            }

            # Calculate total billed vs total commented hours
            mismatch_df = cd.calc_total_billed_n_comment(invoice_sheet_name, invoice_sheet_names, dataframes, workbook_filename)

            # Paste data from dataframes into worksheets
            ut.paste_all_to_excel(dataframes, workbook_filename, invoice_sheet_name)

            # Format worksheets
            fe.format_all_code_sheets(workbook_filename, dataframes, invoice_sheet_name)


if __name__ == '__main__':
    main()
    print("Program Completed!\nGoodbye!")
    time.sleep(1)  # Gives the end user time to read the message above
