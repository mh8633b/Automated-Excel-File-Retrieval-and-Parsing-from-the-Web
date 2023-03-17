import pandas
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import zipfile
import os
import sys
import pdb


def initializing_driver(chrome_driver_path, download_folder=None):
    try:
        prefs = {}
        chrome_options = Options()

        if download_folder:
            if ':' not in download_folder:
                download_folder = os.path.join(os.getcwd(), download_folder)
            download_folder = download_folder.replace('/', '\\')
            os.makedirs(download_folder, exist_ok=True)
            prefs["download.default_directory"] = download_folder

        prefs["profile.default_content_setting_values.automatic_downloads"] = 1
        prefs['download.prompt_for_download'] = False
        chrome_options.add_experimental_option("prefs", prefs)
        driver_temp = webdriver.Chrome(executable_path=chrome_driver_path, options=chrome_options)
        driver_temp.maximize_window()

        return driver_temp, download_folder
    except:
        print('Error initializing the chrome.')
        os.system('pause')
        sys.exit()


def downloading_wait(download_folder_path, max_wait, file_type='crdownload'):
    try:
        count = 0
        completion_flag = True
        while count < max_wait and completion_flag:
            time.sleep(1)
            completion_flag = False

            for file in os.listdir(download_folder_path):
                if file_type in file:
                    completion_flag = True

            count += 1
            time.sleep(1)

        if count == max_wait:
            return False
        else:
            return True
    except:
        print(sys.exc_info()[1])


def newest(path, required_file='.xlsx'):
    latest_file = None
    for i in range(20):
        try:
            files = os.listdir(path)
            paths = [os.path.join(path, basename) for basename in files]

            if len(paths):
                latest_file = max(paths, key=os.path.getctime)
                if required_file == latest_file[len(required_file)*-1:]:
                    return latest_file
                else:
                    time.sleep(1)
                    continue
        except:
            print(sys.exc_info()[1])

    return latest_file


if __name__ == '__main__':
    # Starting From Here

    # DOWNLOADING FILE
    download_path = 'Files'
    driver, download_path = initializing_driver(chrome_driver_path='chromedriver.exe', download_folder=download_path)

    driver.get('https://www.webhere.com.pk/downloads')

    for count in range(60):
        try:
            # if exists then click and break loop
            driver.find_element_by_xpath('//*[@id="downloads"]/div[2]/div[2]/div[1]/ul/li[1]/a').click()
            break
        except:
            # if does not exist than wait
            time.sleep(1)

    # waiting when file is download
    downloading_wait(download_folder_path=download_path, max_wait=180)

    file_downloaded = newest(path=download_path, required_file='.z')
    print('File Downloaded: \t %s' % file_downloaded)

    driver.quit()

    # EXTRACTING FILE
    with zipfile.ZipFile(file_downloaded, 'r') as zip_ref:
        zip_ref.extractall(download_path)

    lis_file = newest(path=download_path, required_file='.lis')
    print('File Extracted: \t %s' % lis_file)

    # READING PREVIOUS DATA IF ANY
    file_name = os.path.join(os.getcwd(), 'All Data File.xlsx')

    if os.path.exists(file_name):
        previous_data = pandas.read_excel(file_name)
    else:
        previous_data = pandas.DataFrame()

    # creating data frame to store and apply operations on data
    today_data = pandas.DataFrame()

    # PARSING FILE DATA
    with open(lis_file) as f:
        for line in f:
            data_line = dict()

            # Getting Data
            data_line['Date'] = line.split('|')[0]
            data_line['Company Code'] = line.split('|')[1]
            data_line['Company Name'] = line.split('|')[3]
            data_line['Turnover'] = line.split('|')[8]
            data_line['Prv. Rate'] = line.split('|')[9]
            data_line['Open Rate'] = line.split('|')[4]
            data_line['Highest'] = line.split('|')[5]
            data_line['Lowest Rate'] = line.split('|')[6]
            data_line['Last Rate'] = line.split('|')[7]

            today_data = today_data.append(data_line, ignore_index=True)

    # re ordering columns
    today_data = today_data[["Date", "Company Code", "Company Name", "Turnover", "Prv. Rate", "Open Rate", "Highest", "Lowest Rate", "Last Rate"]]

    # unifying data types of data
    today_data['Turnover'] = today_data['Turnover'].astype(float)
    today_data['Prv. Rate'] = today_data['Prv. Rate'].astype(float)
    today_data['Open Rate'] = today_data['Open Rate'].astype(float)
    today_data['Highest'] = today_data['Highest'].astype(float)
    today_data['Lowest Rate'] = today_data['Lowest Rate'].astype(float)
    today_data['Last Rate'] = today_data['Last Rate'].astype(float)

    fluctuation_data = pandas.DataFrame()

    # if there is previous data only then need to compare
    if len(previous_data):
        # GETTING FLUCTUATING TURNOVER VALUE
        # iterate over today's data
        for index, today_data_line in today_data.iterrows():
            # line contains today's one data line
            # get company's previous data
            comp_prev_data = previous_data[previous_data['Company Code'] == today_data_line['Company Code']]

            # get latest turnover value from previous data (yesterday value)
            prev_turnover_value = comp_prev_data['Turnover'].tolist()[0]

            # now comparing latest turnover value in previous data with today's turn over value
            # comparing with one day old data
            # if today's value is 1.5 times or greater then of yesterday's value - consider it as fluctuation
            # not considering when turnover is 0
            if (today_data_line['Turnover'] > 0) and (today_data_line['Turnover'] >= (prev_turnover_value * 1.5)):
                print('POSITIVE FLUCTUATION IN \t "%s - %s"' % (today_data_line['Company Code'], today_data_line['Company Name']))

                # saving data line if there is fluctuation
                fluctuation_data = fluctuation_data.append(today_data_line, ignore_index=True)

    if len(fluctuation_data):
        # re ordering columns
        fluctuation_data = fluctuation_data[["Date", "Company Code", "Company Name", "Turnover", "Prv. Rate", "Open Rate", "Highest", "Lowest Rate",
                                             "Last Rate"]]

    # SAVING DATA
    # appending previous data below to the today's data
    all_data = today_data.append(previous_data, ignore_index=True)

    # writing/saving data to excel
    writer = pandas.ExcelWriter(file_name, engine='xlsxwriter')

    all_data.to_excel(writer, sheet_name='Complete Data', index=False)
    fluctuation_data.to_excel(writer, sheet_name='Fluctuated Data', index=False)

    writer.save()
    writer.close()

    print('Output Saved As: \t %s' % file_name)

    # delete extracted file
    os.remove(lis_file)

    os.system("pause")
    sys.exit()

###################################
# REFERENCE LINKS
"""
pythons basic --- https://www.w3schools.com/python/python_intro.asp    OR      https://www.tutorialspoint.com/python/index.htm
Python Download --- https://www.python.org/downloads/
pycharm --- https://www.jetbrains.com/pycharm/download/#section=windows (IDE to write and execute Code)

# to execute code
1. First install python (recommended settings)
        + make sure to check "Add python to PATH" option
        + help - https://realpython.com/installing-python/
    
2. Install libraries (pandas, selenium, zipfile)
        + open cmd - enter command "pip install library_name"

3. To see/run code open it in pycharm (main.py)

4. To re create exe
        + open cmd in folder where py file is
        + run command "pyinstaller --onefile main.py"
        + first need to install pyintaller using pip command
        + new exe will be in dist folder

5. Chrome driver
        + chromedriver.exe is driver to operate chrome
        + it must be in same folder where main.py file or exe is
        
"""






