from math import nan
from math import isnan
import os
import glob
from tracemalloc import start
import xlsxwriter
import re
import tabula
import pandas as pd
import tkinter
import tkinter.filedialog
from pdfminer import high_level
import PySimpleGUI as sg
import traceback
from LICOR import linear_regression
import copy
        
from csv import DictWriter

def process_times(intervals):

    time_regex = r"([\d.]*)(min|sec)"

    times_dict = {}
    times_list = []

    times = intervals.split(',')
    for time in times:
        time = time.strip(' \t\n\r')
        match = re.search(time_regex, time)
        if match:
            times_dict[match[0]] = {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''}
            if match[2] == 'min':
                times_list.append(float(match[1]))
            if match[2] == 'sec':
                times_list.append(float(match[1])*60)
        else:
            raise Exception("Error parsing time intervals, ensure they were entered correctly")
    return times_dict, times_list

def get_PDFs(folder):
    PDFs = glob.glob(folder + "/*.pdf")
    return PDFs


def process_regex(naming_regex):
    regex = naming_regex.replace('s', '\\D+')
    regex = regex.replace('d', '\\d+(?:\\.\\d+)?')
    regex = regex.replace('x', '.+')
    regex = regex.replace('-', '[\\s_-]+')
    return regex, naming_regex.count('(')


def process_date(date):

    date_regex = r"([a-zA-Z]+)[\s-]?(\d+)[,sth]*[\s-]?(\d{2,4})?"

    months = {'jan' : 'January', 'january' : 'January',
            'feb': 'February', 'february': 'February',
            'mar': 'March', 'march': 'March',
            'apr' : 'April', 'april': 'April',
            'may': 'May',
            'jun': 'June', 'june': 'June',
            'jul': 'July', 'july': 'July',
            'aug': 'August', 'august': 'August',
            'sep': 'September', 'sept': 'September', 'september': 'September',
            'oct': 'October', 'october': 'October',
            'nov': 'November', 'november': 'November',
            'dec': 'December', 'december': 'December'}

    match = re.search(date_regex, str(date))

    if match:
        if str(match[1]).lower() in months:
            month = months[match[1].lower()]
            day = match[2]
            year = match[3]
            return month + " " + day + (", " + year if year else ", 2021")
        else:
            return date
    else:
        if isinstance(date, float):
            if isnan(date):
                return 'empty'
        else:
            return str(date)


def process_PDFs(samples, standards, sample_rows, PDFs, window, regex, count):
    counter = 0
    row_index = 1
    longest_file_name = 0
    standard_regex = r"(\d*)\s?ppm"
    sample_regex = r"([SHLC\d]*)-*\s*([\d.]*[\s-]*min)"
    date_ran_regex = r"[a-zA-Z]{3}\d{2}"
    concentration_regex = r"-?\d*\.\d*"
    dates = []

    for file in PDFs:

        order_check_re = r"Methane"
        order_check = high_level.extract_text(file, '', 0)
        order_check_match = re.search(order_check_re, order_check)

        if order_check_match:
            order = {0: 0, 1: 1, 2: 2, 3: 3, 4: 4}
        else:
            order = {0: 0, 1: 2, 2: 3, 3: 1, 4: 4}

        if len(file) > longest_file_name:
            longest_file_name = len(file)

        event, values = window.read(timeout=10)
        window['-PROG-'].update(counter + 1)
        counter += 1
        if file == 'merged.pdf':
            continue

        print("Processing: " + file)

        tables = tabula.read_pdf(file, pages = 'all')
        methane = tables[0]
        methane = methane.reset_index()
        num_samples = len(tables[0].index)

        processed_tables = []
        processed_tables.append(tables[0])
        

        merged = [0]
        for i in range(len(tables)):
            if i not in merged:
                temp_table = tables[i]
                while len(temp_table.index) != num_samples:
                    k = i + 1
                    temp_table = pd.concat([temp_table, tables[k]])
                    merged.append(k)
                processed_tables.append(temp_table)


        gasses = ['Methane', 'Carbon Dioxide', 'Oxygen', 'Nitrogen', 'Nitrous Oxide']

        for i in range(len(processed_tables)):
            processed_tables[i] = processed_tables[i].reset_index()
            print(processed_tables[i])
            for index, row in processed_tables[i].iterrows():
                if re.search(standard_regex, row['Sample Name']):
                    standard = row['Sample Name'].replace(" ", "")
                    #standard = int(re.search(standard_regex, row['Sample Name'])[1])
                    standards[standard][gasses[order[i]]].append(float(re.search(concentration_regex, row['Conc. Unit'])[0]))
                    if gasses[i] == 'Methane':
                        standards[standard]['File'].append(file)
                
                else:
                    date = process_date(row['Sample ID'])

                    if date.lower() != 'blank':
                        if date not in dates:
                            dates.append(date)
                        if date not in samples:
                            samples[date] = {}
                        current_sample = samples[date]
                        sample_name = re.search(regex, row['Sample Name'])
                        if sample_name:
                            for j in range(count):
                                if sample_name[j + 1] in current_sample:
                                    current_sample = current_sample[sample_name[j + 1]]
                                else:
                                    if j + 1 == count:
                                        current_sample[sample_name[j + 1]] = ['', '', '', '', '', '']
                                        current_sample = current_sample[sample_name[j + 1]]
                                    else:
                                        current_sample[sample_name[j + 1]] = {}
                                        current_sample = current_sample[sample_name[j + 1]]
                            
                            current_sample[order[i]] = float(re.search(concentration_regex, row['Conc. Unit'])[0])
                            current_sample[5] = file
    return samples, standards, sample_rows, longest_file_name, dates




def flux_helper(val, count, current):
    if current > count:
        for k, v in val.items():
            flux_helper(v, count, current + 1)
    
    else:
        for k, v in val.items():
            v["RoC (ppm/min)"] = []
            v["R2"] = []
            for i in range(0, 5):
                X = []
                Y = []
                for time in v:
                    if time == "RoC (ppm/min)" or time == "R2":
                        continue
                    
                    X.append(float(time))
                    Y.append(float(v[time][i]))

                try:
                    m, b, R2 = linear_regression(X, Y)
                    v["RoC (ppm/min)"].append(round(m, 3)) 
                    v["R2"].append(round(R2, 3))
                except:
                    v["RoC (ppm/min)"].append('')
                    v["R2"].append('')


def flux(something, count):
    for k, v in something.items():
        flux_helper(v, count - 1, 0)


def flatten(something):
    list_o_list = []
    for k, v in something.items():
        key_list = [k]
        helper(v, key_list, list_o_list)
    return list_o_list


def helper(val, key_list, list_o_list):
    if isinstance(val, dict):
        for k, v in val.items():
            key_list.append(k)
            helper(v, key_list, list_o_list)
            key_list.pop()

    else:
        list_o_list.append(key_list + val)
        
        

def output_data(samples, standards, sample_rows, longest_file_name, columns, dates, count):
    
    column_number_map = {
        1: 'A',
        2: 'B',
        3: 'C',
        4: 'D',
        5: 'E',
        6: 'F',
        7: 'G',
        8: 'H',
        9: 'I',
        10: 'J'
    }

    print(dates)

    out = tkinter.filedialog.asksaveasfilename(defaultextension='.xlsx')
    workbook = xlsxwriter.Workbook(out)

    title_format = workbook.add_format()
    title_format.set_bg_color('#65BDFB')

    data_format = {}
    flux_format = {}

    data_format[1] = workbook.add_format()
    data_format[1].set_bg_color('#A3D8FC')
    flux_format[1] = workbook.add_format({'bold': True})
    flux_format[1].set_bg_color('#A3D8FC')

    data_format[0] = workbook.add_format()
    data_format[0].set_bg_color('#C2E5FD')
    flux_format[0] = workbook.add_format({'bold': True})
    flux_format[0].set_bg_color('#C2E5FD')

    gasses = ['Methane', 'Carbon Dioxide', 'Oxygen', 'Nitrogen', 'Nitrous Oxide']
    worksheet = workbook.add_worksheet("Standards")
    worksheet.freeze_panes(1, 0)
    worksheet.write_row(0, 0, ["Standard", "Methane (ppm)", "Carbon Dioxide (ppm)", "Oxygen (%)", "Nitrogen (%)", "Nitrous Oxide (ppm)", "File"], title_format)
    row = 1
    data_format_num = 1
    for standard in standards:
        #worksheet.write(i, 0, standard, data_format[data_format_num])
        for k in range(len(standards[standard]['Methane'])):
            worksheet.write_row(row , 0, [standard, standards[standard]['Methane'][k], standards[standard]['Carbon Dioxide'][k], standards[standard]['Oxygen'][k], standards[standard]['Nitrogen'][k], standards[standard]['Nitrous Oxide'][k], standards[standard]['File'][k]], data_format[data_format_num])
        #worksheet.write_column(i, len(gasses) + 1, standards[standard]['File'], data_format[data_format_num])
            row += 1
        data_format_num = 1 - data_format_num
    
    worksheet.set_column(0, 0, len("standard"))
    worksheet.set_column(1, 1, len("Methane (ppm)"))
    worksheet.set_column(2, 2, len("Carbon Dioxide (ppm)"))
    worksheet.set_column(3, 3, len("Oxygen (%)"))
    worksheet.set_column(4, 4, len("Nitrogen (%)"))
    worksheet.set_column(5, 5, len("Nitrous Oxide (ppm)"))
    worksheet.set_column(6, 6, longest_file_name)

    for date in dates:
        if str(date).lower() != 'blank':
            worksheet = workbook.add_worksheet(str(date))
            worksheet.set_column(0 + len(columns), 0 + len(columns), len("Methane (ppm)"))
            worksheet.set_column(1 + len(columns), 1 + len(columns), len("Carbon Dioxide (ppm)"))
            worksheet.set_column(2 + len(columns), 2 + len(columns), len("Oxygen (%)"))
            worksheet.set_column(3 + len(columns), 3 + len(columns), len("Nitrogen (%)"))
            worksheet.set_column(4 + len(columns), 4 + len(columns), len("Nitrous Oxide (ppm)"))
            worksheet.set_column(5 + len(columns), 5 + len(columns), longest_file_name)

            row = 1
            data_format_num = 1

            worksheet.freeze_panes(1, 0)
            for k in range(len(columns)):
                worksheet.set_column(k, k, len(columns[k]))
                if k == count - 1:
                    worksheet.set_column(k, k, len("RoC (ppm/min)"))
            worksheet.write_row(0, 0, columns + ["Methane (ppm)", "Carbon Dioxide (ppm)", "Oxygen (%)", "Nitrogen (%)", "Nitrous Oxide (ppm)", "File"], title_format)
            current_sample = samples[0][1]

            for i in range(len(samples)):
                if samples[i][0] == date:
                    if samples[i][1] != current_sample:
                        data_format_num = 1 - data_format_num
                        current_sample = samples[i][1]
                    if samples[i][count] == 'RoC (ppm/min)':
                        worksheet.write_row(row, 0, samples[i][1:] + [''], flux_format[data_format_num])
                    elif samples[i][count] == 'R2':
                        worksheet.write_row(row, 0, samples[i][1:] + [''], flux_format[data_format_num])
                        sample_row = samples[i][1:-6]
                        row += 1
                        worksheet.write_row(row, 0, sample_row + ['Volume: ', 1, 'Temp: ', 1, 'Surface Area:', 1, ''], flux_format[data_format_num])
                        row += 1

                        CH4_flux = '=({}{}*({}{}/(0.0821*{}{}))*(0.044*1440)/({}{})*(12/44)/1000)'.format(
                            column_number_map[count + 1], row - 2,
                            column_number_map[count + 1], row,
                            column_number_map[count + 3], row,
                            column_number_map[count + 5], row)

                        CO2_flux = '=({}{}*({}{}/(0.0821*{}{}))*(0.044*1440)/({}{})*(12/44)/1000)'.format(
                            column_number_map[count + 2], row - 2,
                            column_number_map[count + 1], row,
                            column_number_map[count + 3], row,
                            column_number_map[count + 5], row)

                        worksheet.write_row(row, 0, sample_row + ["Flux", CH4_flux, CO2_flux, '', '', '', ''], flux_format[data_format_num])
                                                                                                    
                    else:   
                        worksheet.write_row(row, 0, samples[i][1:], data_format[data_format_num])
                    row += 1
    workbook.close()
    return out



def GC():
    sample_rows = {}
    samples = {}
    standards = {}

    layout = [[sg.Text('GC Data Processing Tool', font='Any 36', background_color='#0680BF')],
        [sg.Text("", background_color='#0680BF')],
        [sg.Text('PDF files folder:', size=(15, 1), background_color='#0680BF'), sg.Input(key='-FOLDER-'), sg.FolderBrowse()],
        [sg.Text('Naming regex:', size=(15, 1), background_color='#0680BF'), sg.InputText(key='-REGEX-')],
        [sg.Text('Columns', size=(15, 1), background_color='#0680BF'), sg.InputText(key='-COLS-')],
        [sg.Text('Standards', size=(15, 1), background_color='#0680BF'), sg.InputText(key='-STANDARDS-')],
        [sg.Checkbox("Calculate flux?", key="-FLUX-", background_color='#0680BF', enable_events=True)],
        [sg.Text("", background_color='#0680BF')],
        [sg.Submit(), sg.Cancel()]]

    # Create the window
    window = sg.Window("GC", layout, margins=(80, 50), background_color='#0680BF')
    cancelled = False

    # Create an event loop
    while True:
        event, values = window.read()
        # End program if user closes window or
        # presses the OK button
        if event == "Submit":
            break
        
        elif event == "Cancel" or event == sg.WIN_CLOSED:
            cancelled = True
            break

    
    if cancelled == False:

        try:
            print("Reading files")
            folder = values['-FOLDER-']


            PDFs = get_PDFs(folder)

            if len(PDFs) == 0:
                raise Exception("Error: No PDFs found at selected location")

            BAR_MAX = len(PDFs)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error in inputted information")


        try:
            regex, count = process_regex(values['-REGEX-'])
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error in inputted naming regex")

        try:
            stans = values['-STANDARDS-'].split(',')
            for stan in stans:
                stan = stan.strip(' \t\n\r')
                standards[stan] = {'Methane': [], 'Carbon Dioxide': [], 'Oxygen': [], 'Nitrogen': [], 'Nitrous Oxide': [], 'File': []}
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error in inputted standards")

        try:
            columns = values['-COLS-'].split(', ')
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error in inputted columns")


        layout = [[sg.Text('GC Data Processing Tool', font='Any 36', background_color='#0680BF')],
            [sg.Text("Processing... This may take a few minutes", background_color='#0680BF')],
            [sg.Text("", background_color='#0680BF')],
            [sg.ProgressBar(BAR_MAX, orientation='h', size=(40, 15), key='-PROG-', bar_color=('#38E210', '#FFFFFF'))]]

        window.close()
        window = sg.Window("GC", layout, margins = (80, 50), background_color='#0680BF')
        

        try:
            print("Processing PDFs")
            samples, standards, sample_rows, longest_file_name, dates = process_PDFs(samples, standards, sample_rows, PDFs, window, regex, count)
            print(samples)


            if values['-FLUX-']:
                flux(samples, count)

            samples = flatten(samples)
            for row in samples:
                print(row)
            print(samples)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error processing PDFs: ensure MARIA.isr report format is used")

        

        try:
            print("Outputting results")
            out = output_data(samples, standards, sample_rows, longest_file_name, columns, dates, count)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error outputting results: ensure chosen location is valid")

    window.close()
    return 0