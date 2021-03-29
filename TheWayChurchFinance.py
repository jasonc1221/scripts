import pandas as pd
import argparse
import datetime
import os
from dateutil.relativedelta import *
import shutil

class TheWayChurchFinance:
    def __init__(self):
        self.finance_file = 'TheWayChurchFinance.xlsx'
        self.start_date = '01/2021'
        self.end_date = ''
        self.account_codes_file = 'AccountCodes.xlsx'
        self.journal_file = 'journal.xlsx'
        self.account_history_file = 'AccountHistory.csv'
        self.finance_df = {}

        self.get_args()
        self.main()
        print('**********PROGRAM RAN SUCCESSFULLY**********')
        # close = input('Press any key to close')
    
    def get_args(self):
        parser = argparse.ArgumentParser("Get files and date for TheWayChurchFinance.xlsx")
        parser.add_argument('--account-codes-file', help='Account codes for every department', default='AccountCodes.xlsx')
        parser.add_argument('--journal-file', help='Journal file of checks written', default='journal.xlsx')
        parser.add_argument('--account-history-file', help='Hanmi Bank Account History File', default='AccountHistory.csv')
        parser.add_argument('--start-date', default='')
        parser.add_argument('--end-date', default='')
        args = parser.parse_args()

        self.account_codes_file = args.account_codes_file
        self.journal_file = args.journal_file
        self.account_history_file = args.account_history_file
        self.start_date = args.start_date
        self.end_date = ''

    def main(self):
        self.create_copy_of_old_finance_sheet()
        
        # Creating pd.Dataframe of files
        self.account_codes = self.get_dataframe_of_file(self.account_codes_file)
        self.journal = self.get_dataframe_of_file(self.journal_file)
        self.account_history = self.get_dataframe_of_file(self.account_history_file)
    
        # Creating the start and end datetime variables
        self.start_datetime = datetime.datetime.strptime(self.start_date, '%m/%Y') if self.start_date else datetime.datetime.strptime('01/2021', '%m/%Y')
        self.end_datetime = datetime.datetime.strptime(self.end_date, '%m/%Y') if self.end_date else datetime.datetime.now()

        # Extracting data from files
        self.account_codes_extracted = self.extract_account_codes()
        self.journal_checks = self.extract_journal_checks()
        self.account_history_checks = self.extract_account_history_checks()

        # Creating Excel with data extracted
        self.create_finance_sheet_AccountCodeBalance()
        self.create_finance_sheet_MatchedChecks()
        self.write_finance_sheet()
    
    def raise_exception(self, file_name, error_msg, index, row,  index_offset=2):
        print(f'____________FIX ERROR in {file_name}____________')
        print(f'{error_msg} in {file_name}')
        print(f'ROW #{index + index_offset}')
        print(row)
        raise Exception(f'{error_msg} in {file_name}')

    def create_copy_of_old_finance_sheet(self):
        copy_folder = os.path.join(os.getcwd(), 'copy')
        if not os.path.isdir(copy_folder):
            os.mkdir(copy_folder) 

        if self.finance_file in os.listdir():
            print('Creating copy of old finance sheet')
            now = datetime.datetime.now()
            month_day_year_hour_min = now.strftime('%m_%d_%Y_%H_%M')
            old_file = os.path.join(os.getcwd(), self.finance_file)
            old_copy_file = os.path.join(os.getcwd(), f'copy/TheWayChurchFinance_{month_day_year_hour_min}.xlsx')
            shutil.copy(old_file, old_copy_file)

    def get_dataframe_of_file(self, file_name):
        # Checking if file is correct file type
        if '.csv' not in file_name and '.xlsx' not in file_name:
            raise Exception('{} needs to be csv or xlsx'.format(file_name))

        # Return pandas.Dataframe
        if '.csv' in file_name:
            return pd.read_csv(file_name)
        if '.xlsx' in file_name:
            return pd.read_excel(file_name)

    def write_finance_sheet(self):
        # finance_df is a dict of pandas.DataFrame
        # Create or overwrite sheet
        with pd.ExcelWriter(self.finance_file, engine='xlsxwriter') as writer:
            workbook  = writer.book

            # Get all month_years to date
            month_year = []
            date = self.start_datetime
            while date < self.end_datetime:
                month_year_text = date.strftime('%h%Y')
                month_year.append(month_year_text)
                # Increment date month by 1
                date = date + relativedelta(months=+1)
            
            # Putting All Sheets into xlsx
            for sheet in self.finance_df:
                sheet_name = sheet

                startcol = 0 
                # # Post DF starts +1 column away from Signed DF
                if ' ' in sheet:
                    if 'Post' in sheet:
                        # startcol = len(self.finance_df[sheet].columns) + 1
                        pass
                    temp_sheet_name = sheet_name.split()[0]
                
                self.finance_df[sheet].to_excel(writer, sheet_name=sheet_name, index=False, startcol=startcol)
                worksheet = writer.sheets[sheet_name]

                # Get the dimensions of the dataframe.
                (max_row, max_col) = self.finance_df[sheet].shape

                # Set the autofilter.
                worksheet.autofilter(0, 0, max_row, max_col - 1)

                # Make the columns wider for clarity
                worksheet.set_column(0,  max_col - 1, 15)
                money_format = workbook.add_format({'num_format': '$#,##0.00'})
                # Formatting Money Columns
                if sheet_name == 'AccountCodeBalance':
                    worksheet.set_column(5, 5+len(month_year), None, money_format)
                elif 'Post' in sheet_name or 'Signed' in sheet_name:
                    worksheet.set_column(2, 3, None, money_format)

    def extract_account_codes(self):
        print('Extracting Account Codes')
        account_codes_extracted = {}
        for index, row in self.account_codes.iterrows():
            row_data = {
                'Account Group Name': row['Account Group Name'] if not pd.isna(row['Account Group Name']) else '',
                'Account Group': int(row['Account Group']) if not pd.isna(row['Account Group']) else 0,
                'Account Name': row['Account Name'] if not pd.isna(row['Account Name']) else '',
                'Account': int(row['Account']) if not pd.isna(row['Account']) else 0
            }
            if pd.isna(row['Account Group']):
                continue
            if row_data['Account'] in account_codes_extracted:
                self.raise_exception(self.account_codes_file, 'Duplicate Account', index, row)
            else:
                account_codes_extracted[row_data['Account']] = row_data
        return account_codes_extracted

    def extract_journal_checks(self):
        print('Extracting Journal Checks')
        journal_checks = {}
        for index, row in self.journal.iterrows():
            # Check if timestamp of check is between start and end datetime
            date = row['Date']
            if self.start_datetime < date < self.end_datetime:
                # Ignore deposits and stop on -split- rows
                if row['Account'] == '-split-':
                    # Ignore deposit rows
                    deposit = row['Deposit'] if not pd.isna(row['Deposit']) else 0
                    if deposit:
                        continue
                        # old code for adding up deposits
                        # self.account_codes_extracted[0][month_year_text] = self.account_codes_extracted[0].get(month_year_text, 0) + row_data['Deposit']
                    else:
                        account_code = row['Account']
                        self.raise_exception(self.journal_file, f'Invalid Account Code {account_code}', index, row)
                try:
                    row_data = {
                        'Account': int(row['Account'].split()[0]) if not str(row['Account']).isdigit() else int(row['Account']),
                        'Payment': row['Payment'] if not pd.isna(row['Payment']) else 0,
                        'Deposit': row['Deposit'] if not pd.isna(row['Deposit']) else 0,
                        'Date': row['Date'].strftime('%m/%d/%Y')
                    }
                except:
                    self.raise_exception(self.journal_file, f'Bad Row', index, row)
                # Get month year of the journal check e.g Jan 2021, Feb 2021, Mar 2021...
                month_year_text = date.strftime('%h %Y')
                # Add to Account Code month sum
                if not pd.isna(row['Number']) and str(row['Number']).isdigit(): # Check Number
                    check_num = int(row['Number'])
                    if check_num in journal_checks:
                        self.raise_exception(self.journal_file, f'Duplicate Check Number {check_num}', index, row)
                    else:
                        journal_checks[int(row['Number'])] = row_data
                if self.account_codes_extracted.get(row_data['Account']):
                    self.account_codes_extracted[row_data['Account']][month_year_text] = self.account_codes_extracted[row_data['Account']].get(month_year_text, 0) + row_data['Payment']
                    self.account_codes_extracted[row_data['Account']][month_year_text] = round(self.account_codes_extracted[row_data['Account']][month_year_text], 2)
                else:
                    account_code = row['Account']
                    self.raise_exception(self.journal_file, f'Invalid Account Code {account_code}', index, row)

        return journal_checks
                    
    def extract_account_history_checks(self):
        print('Extracting AccountHistory Checks')
        account_history_checks = {}
        for index, row in self.account_history.iterrows():
            date = datetime.datetime.strptime(row['Post Date'], '%m/%d/%Y')
            if self.start_datetime < date < self.end_datetime:
                row_data = {
                    'Post Date': date.strftime('%m/%d/%Y'),
                    'Debit': row['Debit'] if not pd.isna(row['Debit']) else 0,
                    'Credit': row['Credit'] if not pd.isna(row['Credit']) else 0,
                }
                if not pd.isna(row['Check']): # Check Number
                    account_history_checks[int(row['Check'])] = row_data
                elif row['Description'] and 'CHECK' in row['Description']:
                    check_num = int(row['Description'].split()[-1])
                    account_history_checks[check_num] = row_data
        return account_history_checks

    def create_finance_sheet_AccountCodeBalance(self):
        print('Creating Finance Sheet AccountCodeBalance')
        sheet_name = 'AccountCodeBalance'
        account_codes_balance_df = self.account_codes.copy(deep=True)
        # Create columns with month_year_sum from account_codes_extracted
        date = self.start_datetime
        while date < self.end_datetime:
            # Create new month year column if does not exist
            month_year_text = date.strftime('%h %Y')
            if month_year_text not in account_codes_balance_df.columns:
                account_codes_balance_df[month_year_text] = '0'
            
            # Change value of month year sum if sum exists
            for index, row in account_codes_balance_df.iterrows():
                account = int(row['Account'])
                if self.account_codes_extracted.get(account):
                    month_year_sum = self.account_codes_extracted[account].get(month_year_text, 0)
                    account_codes_balance_df.at[index, month_year_text] = month_year_sum
            
            # Increment date month by 1
            date = date + relativedelta(months=+1)
        
        # Create AccountCodeBalance sheet
        self.finance_df[sheet_name] = account_codes_balance_df

    def create_finance_sheet_MatchedChecks(self):
        print('Creating Finance Sheet MatchedChecks')
        sheet_name = 'MatchedChecks'

        # Check what Journal checks match in AccountHistory Checks
        j_checks = set(self.journal_checks.keys())
        ah_checks = set(self.account_history_checks.keys())
        checks_matching = j_checks & ah_checks
        # print('Total Journal Checks:', len(j_checks))
        # print('Total AccountHistory Checks:', len(ah_checks))
        # print('Journal Checks in AccountHistory Checks:', len(checks_matching))

        # checks_matching data
        matched_checks_dict = {}
        matched_sum = 0
        for check in checks_matching:
            matched_checks_dict[check] = self.journal_checks[check]
            matched_checks_dict[check]['Post Date'] = self.account_history_checks[check]['Post Date']
            matched_sum += self.journal_checks[check]['Payment']
            matched_sum = round(matched_sum, 2)
        start_month_day_year = self.start_datetime.strftime('%h %d %Y')
        end_month_day_year = self.end_datetime.strftime('%h %d %Y')
        # print(f'Sum of Journal Checks in AccountHistory between {start_month_day_year} - {end_month_day_year}: ${matched_sum}')
        
        # Create matched_checks_df for all months
        matched_checks_dfs = {}
        latest_month = ''
        date = self.start_datetime
        while date < self.end_datetime:
            month_year_text = date.strftime('%h%Y')
            latest_month = month_year_text
            month_year_signed_text = f'{month_year_text} Signed'
            month_year_posted_text = f'{month_year_text} Post'
            matched_checks_dfs[month_year_signed_text] = pd.DataFrame(columns=['Check #', 'Account', 'Paid', 'Pending', 'Signed Date', 'Post Date'])
            matched_checks_dfs[month_year_posted_text] = pd.DataFrame(columns=['Check #', 'Account', 'Paid', 'Pending', 'Signed Date', 'Post Date'])
            # Increment date month by 1
            date = date + relativedelta(months=+1)

        unmatched_checks = []
        # Put checks in correct months
        for check in self.journal_checks:
            new_row_data = {
                'Check #': check, 
                'Account': self.journal_checks[check]['Account'], 
                'Paid': 0, 
                'Pending': 0,
                'Signed Date': self.journal_checks[check]['Date'],
                'Post Date': ''
            }
            payment = self.journal_checks[check]['Payment']
            if check in matched_checks_dict:
                new_row_data['Paid'] = payment
                new_row_data['Post Date'] = matched_checks_dict[check]['Post Date']
                # Adding to appropriate Month Year Post Sheet
                post_month_year_text = datetime.datetime.strptime(new_row_data['Post Date'], '%m/%d/%Y').strftime('%h%Y')
                month_year_posted_text = f'{post_month_year_text} Post'
                matched_checks_dfs[month_year_posted_text] = matched_checks_dfs[month_year_posted_text].append(new_row_data, ignore_index=True)
            else:
                new_row_data['Pending'] = payment
                unmatched_checks.append(new_row_data)
            
            # Adding to appropriate Month Year Signed Sheet
            sign_month_year_text = datetime.datetime.strptime(new_row_data['Signed Date'], '%m/%d/%Y').strftime('%h%Y')
            month_year_signed_text = f'{sign_month_year_text} Signed'
            matched_checks_dfs[month_year_signed_text] = matched_checks_dfs[month_year_signed_text].append(new_row_data, ignore_index=True)
        
        # Adding all unmatched checks to latest Month Year Post Sheet
        month_year_posted_text = f'{latest_month} Post'
        for unmatched_check in unmatched_checks:
            matched_checks_dfs[month_year_posted_text] = matched_checks_dfs[month_year_posted_text].append(unmatched_check, ignore_index=True)

        # Adding all matched_checks_dfs into finance_df
        self.finance_df.update(matched_checks_dfs)


if __name__ == '__main__':
    the_way_church_finance = TheWayChurchFinance()