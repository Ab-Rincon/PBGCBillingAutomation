import re
import pandas as pd
import logging
from datetime import datetime, date, timedelta
import utils.utils as ut
import utils.format_excel as fe
import openpyxl
from config.settings import overbilled_loc


def calculate_total_time_difference(in_times, out_times):
    # Initialize a variable to hold the sum of the differences
    total_hours = 0

    # Iterate over the paired in and out times
    for in_time, out_time in zip(in_times, out_times):

        # Parse the time strings to datetime.time objects
        in_time = pd.to_datetime(in_time, format='%H:%M').time()
        out_time = pd.to_datetime(out_time, format='%H:%M').time()

        # Calculate the difference in hours
        difference = (datetime.combine(date.min, out_time) - datetime.combine(date.min, in_time)).total_seconds() / 3600

        # Add the difference to the total
        total_hours += difference

    # Round to the nearest 0.25
    total_hours = round(total_hours * 4) / 4
    return total_hours


class CleanData():
    def __init__(self, df: pd.DataFrame) -> None:
        self.df = df

    def calculate_ts(self):
        def get_week_start_end_dates(date_str: str, date_format: str = "%m/%d/%Y") -> str:
            # Parse the date string into a datetime object
            given_date = datetime.strptime(date_str, date_format)

            # Adjust the start of the week to Sunday
            start_of_week = given_date - timedelta(days=(given_date.weekday() + 1) % 7)

            # Calculate the end of the week as Saturday
            end_of_week = start_of_week + timedelta(days=6)

            # Return the start and end of the week as a formatted string range
            ts_range = f'{start_of_week.strftime(date_format)} - {end_of_week.strftime(date_format)}'
            return ts_range

        # Apply the function to the 'Date' column of the dataframe
        self.df['T/S'] = self.df['Date'].apply(get_week_start_end_dates)
        logging.info('Calculated T/S periods')
        logging.debug(self.df['T/S'])

    def find_empty_comments(self) -> pd.DataFrame:
        # Find empty comments
        self.df["Empty Comment"] = self.df['Time Period'].isna() | (self.df['Time Period'] == "")

        # Create empty comment dataframe
        empty_comment_df = self.df.loc[self.df['Empty Comment'], ['Name', 'T/S', 'Date', 'Empty Comment', 'Time Period']]
        logging.info(f'There are {len(empty_comment_df.index)} empty comments')

        # Extract empty comments going forward
        self.df = self.df.loc[~self.df['Empty Comment']]
        self.df = self.df.drop('Empty Comment', axis=1)
        return empty_comment_df

    def extract_times(self) -> pd.DataFrame:
        if self.df.empty:
            return self.df

        def process_time_entries(row):
            """
            Takes in a row of time comments and extracts in times and outimes and issues
            """

            def extract_time_pairs(text: str):
                """
                Find time in and time out from comment
                """
                # Replace "24:00" with "23:59" for consistency
                logging.debug(f'The current comment is:\n{text}')

                # This pattern looks for pairs of times
                pattern = r"(?:time in:?\s*(\d{1,2}:\d{2}))\s*[-,\u2013]?\s*(?:time out:?\s*(\d{1,2}:\d{2}))"
                time_pairs = re.findall(pattern, text, re.I)

                # Extract lists of in times and out times
                in_times = [in_time for in_time, _ in time_pairs]
                out_times = [out_time for _, out_time in time_pairs]

                # Special handling for midnight formats
                out_times = ["23:59" if time == "00:00" or time == "24:00" or time == "0:00" else time for time in out_times]

                # Match "00:00" or "0:00" as pairs explicitly
                zero_time_pattern = r"\b\d{1,2}:\d{2}\b"
                zero_time_matches = re.findall(zero_time_pattern, text)

                # Assume that "00:00" or "0:00" must occur in pairs for in and out times
                zero_time_pair_count = len(zero_time_matches)
                logging.debug(f'There are currently {len(in_times)} in times and {zero_time_pair_count} pair(s) of zero digit times')

                # Check for format issues
                format_issue = len(in_times) != len(out_times) or len(time_pairs) * 2 != len(zero_time_matches)
                return in_times, out_times, format_issue

            # Convert to string to handle NaN or other non-string types gracefully
            time_text = str(row['Time Period'])
            intimes, outtimes, format_issue = extract_time_pairs(time_text)

            # Check if any time entry doesn't match the 15-minute intervals
            valid_endings = ["00", "15", "30", "45", "59"]
            for time_entry in intimes + outtimes:
                if time_entry[-2:] not in valid_endings:
                    format_issue = True
                    break

            # Check if there are any 'in times' without corresponding 'out times' and vice versa
            if not format_issue and len(intimes) != len(outtimes):
                format_issue = True

            # Check if the lists are empty which also indicates a format issue
            elif not format_issue and not intimes and not outtimes:
                format_issue = True

            return pd.Series([intimes, outtimes, format_issue], index=['In Times', 'Out Times', 'Format Issue'])

        # Apply proces time entry function to dataframe
        self.df[['In Times', 'Out Times', 'Format Issue']] = self.df.apply(process_time_entries, axis=1)

        # Create format issue dataframe
        format_issue_df = self.df.loc[self.df['Format Issue'], ['Name', 'T/S', 'Date', 'Format Issue', 'Time Period']]
        logging.info(f'There are {len(format_issue_df.index)} format issues')

        # Extract format issue comments going forward
        self.df = self.df.loc[~self.df['Format Issue']]
        self.df.drop('Format Issue', axis=1, inplace=True)
        return format_issue_df

    def check_military_time_format(self):
        if self.df.empty:
            return self.df

        def check_military_time(in_times, out_times) -> bool:
            """
            Check if the time in time out order makes sense
            """
            military_time_issue = False

            for in_time, out_time in zip(in_times, out_times):
                try:
                    # Parse the time strings to datetime.time objects
                    in_time_obj = pd.to_datetime(in_time, format='%H:%M').time()
                    out_time_obj = pd.to_datetime(out_time, format='%H:%M').time()

                    # Check if out time is before in time
                    if out_time_obj <= in_time_obj:
                        military_time_issue = True
                        logging.debug(f"in time: {in_time_obj}, out time: {out_time_obj}, military issue: {military_time_issue}")

                except Exception as e:
                    military_time_issue = True
                    logging.debug(f'There is a military time issue. Exception: {e}')
                    return military_time_issue

            return military_time_issue

        # apply check military time function to dataframe
        self.df['Military Time Issue'] = self.df.apply(lambda row: check_military_time(row['In Times'], row['Out Times']), axis=1)

        # Find military time issues
        military_time_issue_df = self.df.loc[self.df['Military Time Issue'], ['Name', 'T/S', 'Date', 'Military Time Issue', 'Time Period']]
        logging.info(f'There are {len(military_time_issue_df.index)} military time issues')

        # Extract empty comments going forward
        self.df = self.df.loc[~self.df['Military Time Issue']]
        self.df.drop('Military Time Issue', axis=1, inplace=True)
        return military_time_issue_df

    def calculate_time_worked(self) -> pd.DataFrame:
        if self.df.empty:
            return self.df, self.df

        def format_time_pair(in_times, out_times):
            formatted_times = []

            for index, (in_time, out_time) in enumerate(zip(in_times, out_times)):
                # Parse the time strings to datetime objects
                in_datetime = pd.to_datetime(in_time, format='%H:%M')
                out_datetime = pd.to_datetime(out_time, format='%H:%M')

                # Format the datetime objects to 'HH:MM' format
                formatted_in_time = in_datetime.strftime('%H:%M')
                formatted_out_time = out_datetime.strftime('%H:%M')

                # Check if the out time is 23:59 and replace it with '00:00' of the next day, generally considered as '24:00'
                if formatted_out_time == '23:59':
                    formatted_out_time = '00:00'

                # Append the formatted time pair to the list
                formatted_pair = f"Time in: {formatted_in_time} - Time out: {formatted_out_time}"
                if (index + 1) % 2 == 0 and index < len(in_times) - 1:  # Add newline after each pair except the last
                    formatted_pair += "\n"
                elif index < len(in_times) - 1:  # Add semicolon only if it's not the last pair
                    formatted_pair += "; "

                formatted_times.append(formatted_pair)

            # Join the formatted time pairs without any additional character, as the semicolon and newline are already added
            return "".join(formatted_times)

        # Apply the function to each row and assign the result to a new column 'Commented Time Worked'
        self.df['Commented Time Worked'] = self.df.apply(lambda row: calculate_total_time_difference(row['In Times'], row['Out Times']), axis=1)
        self.df['Formatted Time Comments'] = self.df.apply(lambda row: format_time_pair(row['In Times'], row['Out Times']), axis=1)

        # Calculate time charged
        self.df['Total Hours Worked'] = self.df.groupby(['Name', 'Date'])['Hours Worked'].transform('sum')

        # Calcaulate where
        # Convert 'In Times' and 'Out Times' to string for comparison purposes
        self.df['In Times Str'] = self.df['In Times'].apply(lambda x: ';'.join(map(str, x)))
        self.df['Out Times Str'] = self.df['Out Times'].apply(lambda x: ';'.join(map(str, x)))

        # Calculate unique 'In Times' and 'Out Times' for each group
        self.df['Unique In Times'] = self.df.groupby(['Name', 'Date'])['In Times Str'].transform('nunique')
        self.df['Unique Out Times'] = self.df.groupby(['Name', 'Date'])['Out Times Str'].transform('nunique')

        # Determine if there's is a conflicting time in and time out for the same date
        self.df['Conflicting Comment'] = ""
        self.df['Conflicting Comment'] = (self.df['Unique In Times'] > 1) | (self.df['Unique Out Times'] > 1)
        conflicting_comment_df = self.df.loc[self.df['Conflicting Comment'], ['Name', 'T/S', 'Date', 'Conflicting Comment', 'Time Period']]
        logging.info(f'There are {len(conflicting_comment_df.index)} conflicting time comments')
        self.df = self.df.loc[~self.df['Conflicting Comment']]
        self.df.drop('Conflicting Comment', axis=1, inplace=True)

        # Clean up temporary columns
        self.df.drop(['In Times Str', 'Out Times Str', 'Unique In Times', 'Unique Out Times'], axis=1, inplace=True)

        # Create acceptable comment df
        acceptable_df = self.df[['Name', 'T/S', 'Date', 'Time Period', 'Formatted Time Comments']]
        logging.info(f'There are {len(acceptable_df.index)} acceptable time comments')
        return conflicting_comment_df, acceptable_df


def calc_total_billed_n_comment(invoice_sheet_name: str, invoice_sheet_names: list, dataframes: dict, workbook_filename: str) -> pd.DataFrame:
    summary_df = dataframes['Summary']
    workbook_df = dataframes['Mismatch']

    # Concat charge code sheet data
    if invoice_sheet_name == invoice_sheet_names[0]:
        workbook_df = summary_df
        return workbook_df

    if invoice_sheet_name != invoice_sheet_names[-1]:
        workbook_df = pd.concat([workbook_df, summary_df], ignore_index=True)
        return workbook_df

    # Add the summed hours back to the original DataFrame
    workbook_df['Daily Total Hours Worked'] = workbook_df.groupby(['Name', 'Date'])['Hours Worked'].transform('sum')
    workbook_df['Daily Commented Time Worked'] = workbook_df.groupby(['Name', 'Date'])['Commented Time Worked'].transform('sum')

    # Add a column to flag conflicts in reported time
    workbook_df['Conflicting Time Worked'] = workbook_df['Daily Total Hours Worked'] > workbook_df['Daily Commented Time Worked']

    # Save the subset with conflicts to another CSV
    subset_columns = ['Name', 'T/S', 'Date', 'Charge Code', 'Conflicting Time Worked', 'Time Period', 'Total Hours Worked',
                      'Commented Time Worked', 'Daily Total Hours Worked', 'Daily Commented Time Worked']
    overbilled_df = workbook_df.loc[workbook_df['Conflicting Time Worked'], subset_columns]
    overbilled_df.sort_values(['Name', 'Date'], ascending=[True, False], inplace=True)
    logging.info(f'There are {len(overbilled_df.index)} mismatched time comments')

    # Create overbilled sheet
    workbook = openpyxl.load_workbook(workbook_filename)
    worksheet = workbook['Mismatch']
    template_sheet = 'Mismatch'
    locations = overbilled_loc

    # Paste values into workbook
    empty_df = ut.create_overbilled_sheet(overbilled_df, worksheet)

    # if there are values format the workbook then move it first
    if not empty_df:
        fe.format_sheet(workbook, workbook_filename, worksheet, template_sheet, locations, dataframes)
        workbook.move_sheet(worksheet, -len(workbook.sheetnames))
        workbook.save(workbook_filename)

    # Save and close workbook
    workbook.save(workbook_filename)
    workbook.close()
    return overbilled_df
