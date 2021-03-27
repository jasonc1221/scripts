import pandas as pd
import argparse
import datetime
import os
from dateutil.relativedelta import *
import shutil

finance_file = 'TheWayChurchFinance.xlsx'
start_datetime = None
end_datetime = None

def create_copy_of_old_finance_sheet():
    copy_folder = os.path.join(os.getcwd(), 'copy')
    if not os.path.isdir(copy_folder):
        os.mkdir(copy_folder) 

    if finance_file in os.listdir():
        print('Creating copy of old finance sheet')
        now = datetime.datetime.now()
        month_day_year_hour_min = now.strftime('%m_%d_%Y_%H_%M')
        old_file = os.path.join(os.getcwd(), finance_file)
        old_copy_file = os.path.join(os.getcwd(), f'copy/TheWayChurchFinance_{month_day_year_hour_min}.xlsx')
        shutil.copy(old_file, old_copy_file)

def get_dataframe_of_file(file_name):
    # Checking if file is correct file type
    if '.csv' not in file_name and '.xlsx' not in file_name:
        raise Exception('{} needs to be csv or xlsx'.format(file_name))

    # Return pandas.Dataframe
    if '.csv' in file_name:
        return pd.read_csv(file_name)
    if '.xlsx' in file_name:
        return pd.read_excel(file_name)

def create_update_finance_sheet(sheet_name, default_df=None):
    if finance_file in os.listdir():
        print('Using existing file')
        finance_df = pd.read_excel(finance_file, sheet_name=None)
        df = finance_df.get(sheet_name)
    else:
        print('Creating new sheet')
        finance_df = {}
        df = default_df.copy(deep=True) if type(default_df) != type(None) else None
    return finance_df, df

def write_finance_sheet(finance_df):
    # finance_df is a dict of pandas.DataFrame
    # Create or overwrite sheet
    with pd.ExcelWriter(finance_file) as writer:
        for sheet in finance_df:
            finance_df[sheet].to_excel(writer, sheet_name=sheet, index=False)

def extract_account_codes(account_codes):
    print('Extacting Account Codes')
    account_codes_extracted = {}
    for index, row in account_codes.iterrows():
        row_data = {
            'Account Group Name': row['Account Group Name'] if not pd.isna(row['Account Group Name']) else '',
            'Account Group': int(row['Account Group']) if not pd.isna(row['Account Group']) else 0,
            'Account Name': row['Account Name'] if not pd.isna(row['Account Name']) else '',
            'Account': int(row['Account']) if not pd.isna(row['Account']) else 0
        }
        if pd.isna(row['Account Group']):
            continue
        else:
            account_codes_extracted[row_data['Account']] = row_data
    return account_codes_extracted

def extract_journal_checks(journal, account_codes_extracted):
    print('Extacting Journal Checks')
    journal_checks = {}
    for index, row in journal.iterrows():
        # Check if timestamp of check is between start and end datetime
        date = row['Date']
        if start_datetime < date and date < end_datetime:
            row_data = {
                'Account': int(row['Account'].split()[0]) if row['Account'] != '-split-' else '',
                'Payment': row['Payment'] if not pd.isna(row['Payment']) else 0,
                'Deposit': row['Deposit'] if not pd.isna(row['Deposit']) else 0,
                'Date': row['Date'].strftime('%m/%d/%Y')
            }
            # Get month year of the journal check e.g Jan 2021, Feb 2021, Mar 2021...
            month_year_text = date.strftime('%h %Y')
            # Handle deposits and alert of -split- rows
            if row['Account'] == '-split-':
                if row_data['Deposit']:
                    account_codes_extracted[0][month_year_text] = account_codes_extracted[0].get(month_year_text, 0) + row_data['Deposit']
                else:
                    print(index)
                    print(row)
                    raise Exception('Invalid Account Code in journal.xlsx -split-')
                continue
            
            # Add to Account Code month sum
            if not pd.isna(row['Number']) and type(row['Number']) == int: # Check Number
                journal_checks[int(row['Number'])] = row_data
            if account_codes_extracted.get(row_data['Account']):
                account_codes_extracted[row_data['Account']][month_year_text] = account_codes_extracted[row_data['Account']].get(month_year_text, 0) + row_data['Payment']
                account_codes_extracted[row_data['Account']][month_year_text] = round(account_codes_extracted[row_data['Account']][month_year_text], 2)
            else:
                print(index)
                print(row)
                raise Exception('Invalid Account Code in journal.xlsx')
    return journal_checks
                
def extract_account_history_checks(account_history):
    print('Extacting AccountHistory Checks')
    account_history_checks = {}
    for index, row in account_history.iterrows():
        date = datetime.datetime.strptime(row['Post Date'], '%m/%d/%Y')
        if start_datetime < date and date < end_datetime:
            row_data = {
                # 'Post Date': datetime.datetime.strptime(row['Post Date'], '%m/%d/%Y'),
                'Post Date': row['Post Date'],
                'Debit': row['Debit'] if not pd.isna(row['Debit']) else 0,
                'Credit': row['Credit'] if not pd.isna(row['Credit']) else 0,
            }
            if not pd.isna(row['Check']): # Check Number
                account_history_checks[int(row['Check'])] = row_data
            elif row['Description'] and 'CHECK' in row['Description']:
                check_num = int(row['Description'].split()[-1])
                account_history_checks[check_num] = row_data
    return account_history_checks

def create_update_finance_sheet_AccountCodeBalance(account_codes, account_codes_extracted):
    print('Creating/Updating Finance Sheet AccountCodeBalance')
    sheet_name = 'AccountCodeBalance'
    finance_df, account_codes_balance_df = create_update_finance_sheet(sheet_name, account_codes)
    # Update or Create columns with month_year_sum from account_codes_extracted
    date = start_datetime
    while date < end_datetime:
        # Create new month year column if does not exist
        month_year_text = date.strftime('%h %Y')
        if month_year_text not in account_codes_balance_df.columns:
            account_codes_balance_df[month_year_text] = '0'
        
        # Change value of month year sum if sum exists
        for index, row in account_codes_balance_df.iterrows():
            account = int(row['Account'])
            if account_codes_extracted.get(account):
                month_year_sum = str(account_codes_extracted[account].get(month_year_text, '0'))
                account_codes_balance_df.at[index, month_year_text] = month_year_sum
        
        # Increment date month by 1
        date = date + relativedelta(months=+1)
    
    # Overwrite existing AccountCodeBalance sheet
    finance_df[sheet_name] = account_codes_balance_df
    write_finance_sheet(finance_df)

def create_update_finance_sheet_MatchedChecks(journal_checks, account_history_checks):
    # TODO Add Date when check was written
    print('Creating/Updating Finance Sheet MatchedChecks')
    sheet_name = 'MatchedChecks'
    finance_df, matched_checks_df = create_update_finance_sheet(sheet_name)

    # Check what Journal checks match in AccountHistory Checks
    j_checks = set(journal_checks.keys())
    ah_checks = set(account_history_checks.keys())
    checks_matching = j_checks & ah_checks
    print('Total Journal Checks:', len(j_checks))
    print('Total AccountHistory Checks:', len(ah_checks))
    print('Journal Checks in AccountHistory Checks:', len(checks_matching))

    # checks_matching data
    matched_checks_dict = {}
    matched_sum = 0
    for check in checks_matching:
        matched_checks_dict[check] = journal_checks[check]
        matched_checks_dict[check]['Post Date'] = account_history_checks[check]['Post Date']
        matched_sum += journal_checks[check]['Payment']
        matched_sum = round(matched_sum, 2)
    start_month_day_year = start_datetime.strftime('%h %d %Y')
    end_month_day_year = end_datetime.strftime('%h %d %Y')
    print(f'Sum of Journal Checks in AccountHistory between {start_month_day_year} - {end_month_day_year}: ${matched_sum}')
    
    if type(matched_checks_df) != type(None):
        print('Updating Pending Checks')
        for index, row in matched_checks_df.iterrows():
            # Go to next row if check is processed
            if float(row['Paid']):
                continue
            # Verify if check has been processed and change Paid equal to Pending and Pending to 0
            elif float(row['Pending']):
                if int(row['Check #']) in matched_checks_dict:
                    matched_checks_df.at[index, 'Paid'] = row['Pending']
                    matched_checks_df.at[index, 'Pending'] = 0
                    matched_checks_df.at[index, 'Post Date'] = matched_checks_dict[row['Check #']]['Post Date']
    else:
        print(f'Creating new {sheet_name}')
        matched_checks_df = pd.DataFrame(columns=['Check #', 'Account', 'Paid', 'Pending', 'Signed Date', 'Post Date'])
    
    print(f'Adding new Journal Checks to {sheet_name}')
    for check in journal_checks:
        if check not in matched_checks_df['Check #'].to_list():
            new_row_data = {
                'Check #': check, 
                'Account': journal_checks[check]['Account'], 
                'Paid': 0, 
                'Pending': 0,
                'Signed Date': journal_checks[check]['Date'],
                'Post Date': ''
            }
            payment = journal_checks[check]['Payment']
            if check in matched_checks_dict:
                new_row_data['Paid'] = payment
                new_row_data['Post Date'] = matched_checks_dict[check]['Post Date']
            else:
                new_row_data['Pending'] = payment
            matched_checks_df = matched_checks_df.append(new_row_data, ignore_index=True)

    finance_df[sheet_name] = matched_checks_df
    write_finance_sheet(finance_df)

def main(args):
    create_copy_of_old_finance_sheet()
    
    # Creating pd.Dataframe of files
    account_codes = get_dataframe_of_file(args.account_codes_file)
    journal = get_dataframe_of_file(args.journal_file)
    account_history = get_dataframe_of_file(args.account_history_file)
    
    # Creating the start and end datetime variables
    global start_datetime
    start_datetime = datetime.datetime.strptime(args.start_date, '%m/%Y') if args.start_date else datetime.datetime.strptime('01/2021', '%m/%Y')
    global end_datetime
    end_datetime = datetime.datetime.strptime(args.end_date, '%m/%Y') if args.end_date else datetime.datetime.now()

    # Extracting data from files
    account_codes_extracted = extract_account_codes(account_codes)
    journal_checks = extract_journal_checks(journal, account_codes_extracted)
    account_history_checks = extract_account_history_checks(account_history)

    # Creating Excel with data extracted
    create_update_finance_sheet_AccountCodeBalance(account_codes, account_codes_extracted)
    create_update_finance_sheet_MatchedChecks(journal_checks, account_history_checks)

if __name__ == '__main__':
    parser = argparse.ArgumentParser("Get files and date for TheWayChurchFinance.xlsx")
    parser.add_argument('--account-codes-file', help='Account codes for every department', default='AccountCodes.xlsx')
    parser.add_argument('--journal-file', help='Journal file of checks written', default='journal.xlsx')
    parser.add_argument('--account-history-file', help='Hanmi Bank Account History File', default='AccountHistory.csv')
    parser.add_argument('--start-date', default='')
    parser.add_argument('--end-date', default='')
    args = parser.parse_args()
    main(args)
    print('**********PROGRAM RAN SUCCESSFULLY**********')
    # close = input('Press any key to close')