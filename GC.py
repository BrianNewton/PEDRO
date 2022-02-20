import os
import glob
import xlsxwriter
import re
import tabula
import pandas as pd
import tkinter
import tkinter.filedialog
import PySimpleGUI as sg
import copy
        

def process_times(intervals):

    time_regex = r"([\d.]*)(min|sec)"

    times_dict = {}
    times_list = []

    times = intervals.split(',')
    for time in times:
        time = time.strip(' \t\n\r')
        match = re.search(time_regex, time)
        if match:
            print(match[0])
            times_dict[match[0]] = {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''}
            if match[2] == 'min':
                times_list.append(float(match[1]))
            if match[2] == 'sec':
                times_list.append(float(match[1])*60)
        else:
            raise Exception("Error parsing time intervals, ensure they were entered correctly")
    print(times_dict)
    return times_dict, times_list

def get_PDFs(folder):
    PDFs = glob.glob(folder + "/*.pdf")
    return PDFs


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
        return date


def process_PDFs(samples, standards, sample_rows, PDFs, window, times_dict):
    counter = 0
    row_index = 1
    standard_regex = r"(\d*)\s?ppm"
    sample_regex = r"([SHLC\d]*)-*\s*([\d.]*[\s-]*min)"
    date_ran_regex = r"[a-zA-Z]{3}\d{2}"
    concentration_regex = r"-?\d*\.\d*"

    print(times_dict)

    for file in PDFs:
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
                    standards[standard][gasses[i]].append(float(re.search(concentration_regex, row['Conc. Unit'])[0]))
                    if gasses[i] == 'Methane':
                        standards[standard]['File'].append(file)
                

                if re.search(sample_regex, row['Sample Name']):
                    sample_name = re.search(sample_regex, row['Sample Name'])[1]
                    if sample_name not in sample_rows:
                        sample_rows[sample_name] = row_index
                        row_index += 3


                    sample_time = re.search(sample_regex, row['Sample Name'])[2].replace(" ", "")

                    date = process_date(row['Sample ID'])
                    if date in samples:
                        if sample_name in samples[date]:
                            samples[date][sample_name][sample_time][gasses[i]] = float(re.search(concentration_regex, row['Conc. Unit'])[0])
                            samples[date][sample_name][sample_time]['File'] = file
                        else:
                            if times_dict:
                                samples[date][sample_name] = copy.deepcopy(times_dict)
                            else:
                                samples[date][sample_name] = {}
                                samples[date][sample_name][sample_time] = {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''}

                            samples[date][sample_name][sample_time][gasses[i]] = float(re.search(concentration_regex, row['Conc. Unit'])[0])
                            samples[date][sample_name][sample_time]['File'] = file
                    else:
                        if times_dict:
                            samples[date] = {sample_name: copy.deepcopy(times_dict)}
                        else:
                            samples[date] = {}
                            samples[date][sample_name] = {}
                            samples[date][sample_name][sample_time] = {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''}
                       
                        samples[date][sample_name][sample_time][gasses[i]] = float(re.search(concentration_regex, row['Conc. Unit'])[0])
                        samples[date][sample_name][sample_time]['File'] = file
    return samples, standards, sample_rows


def output_data(samples, standards, sample_rows, times_dict):
    out = tkinter.filedialog.asksaveasfilename(defaultextension='.xlsx')
    workbook = xlsxwriter.Workbook(out)

    gasses = ['Methane', 'Carbon Dioxide', 'Oxygen', 'Nitrogen', 'Nitrous Oxide']
    worksheet = workbook.add_worksheet("Standards")
    worksheet.freeze_panes(1, 0)
    worksheet.write_row(0, 0, ["Standard", "Methane (ppm)", "Carbon Dioxide (ppm)", "Oxygen (%)", "Nitrogen (%)", "Nitrous Oxide (ppm)", "File"])
    i = 1
    for standard in standards:
        worksheet.write(i, 0, standard)
        for k in range(len(gasses)):
            worksheet.write_column(i, k + 1, standards[standard][gasses[k]])
        worksheet.write_column(i, len(gasses) + 1, standards[standard]['File'])
        i += len(standards[standard][gasses[0]])
    worksheet.set_column(0, 0, len("standard"))
    worksheet.set_column(1, 1, len("Methane (ppm)"))
    worksheet.set_column(2, 2, len("Carbon Dioxide (ppm)"))
    worksheet.set_column(3, 3, len("Oxygen (%)"))
    worksheet.set_column(4, 4, len("Nitrogen (%)"))
    worksheet.set_column(5, 5, len("Nitrous Oxide (ppm)"))

    for date in samples:
        worksheet = workbook.add_worksheet(str(date))
        worksheet.freeze_panes(1, 0)
        worksheet.set_column(0, 0, len("Sample Name"))
        worksheet.set_column(1, 1, len("Sample Time"))
        worksheet.set_column(2, 2, len("Methane (ppm)"))
        worksheet.set_column(3, 3, len("Carbon Dioxide (ppm)"))
        worksheet.set_column(4, 4, len("Oxygen (%)"))
        worksheet.set_column(5, 5, len("Nitrogen (%)"))
        worksheet.set_column(6, 6, len("Nitrous Oxide (ppm)"))
        worksheet.write_row(0, 0, ["Sample Name", "Sample Time", "Methane (ppm)", "Carbon Dioxide (ppm)", "Oxygen (%)", "Nitrogen (%)", "Nitrous Oxide (ppm)", "File"])
        for sample_name in samples[date]:
            sample = samples[date][sample_name]
            if times_dict:
                i = 0
                for time in times_dict:
                    worksheet.write_row(sample_rows[sample_name] + i, 0, [sample_name, time, sample[time]['Methane'], sample[time]['Carbon Dioxide'], sample[time]['Oxygen'], sample[time]['Nitrogen'], sample[time]['Nitrous Oxide'], sample[time]['File']])
                    i += 1
            else:
                for time in sample:
                    worksheet.write_row(sample_rows[sample_name] + i, 0, [sample_name, time, sample[time]['Methane'], sample[time]['Carbon Dioxide'], sample[time]['Oxygen'], sample[time]['Nitrogen'], sample[time]['Nitrous Oxide'], sample[time]['File']])
                    i += 1

    workbook.close()
    return out



def GC():
    sample_rows = {}
    samples = {}
    standards = {'1ppm': {'Methane': [], 'Carbon Dioxide': [], 'Oxygen': [], 'Nitrogen': [], 'Nitrous Oxide': [], 'File': []},
                    '5ppm': {'Methane': [], 'Carbon Dioxide': [], 'Oxygen': [], 'Nitrogen': [], 'Nitrous Oxide': [], 'File': []}, 
                    '50ppm': {'Methane': [], 'Carbon Dioxide': [], 'Oxygen': [], 'Nitrogen': [], 'Nitrous Oxide': [], 'File': []}}
    
    layout_flux = [[sg.Text("Enter time intervals:", size=(15,1), background_color='#0680BF'), sg.InputText(key='-TIMES-')],
        [sg.Text("Eg. \"0min, 3.5min, 7min\"", background_color='#0680BF')]]

    layout = [[sg.Text('GC Data Processing Tool', font='Any 36', background_color='#0680BF')],
        [sg.Text("", background_color='#0680BF')],
        [sg.Text('PDF files folder:', size=(15, 1), background_color='#0680BF'), sg.Input(key='-FOLDER-'), sg.FolderBrowse()],
        [sg.Checkbox("Calculate flux?", key="-FLUX-", background_color='#0680BF', enable_events=True)],
        [sg.Column(layout_flux,  key='-COL1-', visible=False, background_color='#0680BF'), sg.Column([[]], key='-COL2-', visible=True, background_color='#0680BF' )],
        [sg.Text("", background_color='#0680BF')],
        [sg.Submit(), sg.Cancel()]]

    #layout = [[sg.Column(layout, key='-COL1-'), sg.Column(layout2, visible=False, key='-COL2-'), sg.Column(layout3, visible=False, key='-COL3-')],
     #     [sg.Button('Cycle Layout'), sg.Button('1'), sg.Button('2'), sg.Button('3'), sg.Button('Exit')]]


    # Create the window
    window = sg.Window("GC", layout, margins=(80, 50), background_color='#0680BF')
    cancelled = False

    # Create an event loop
    while True:
        event, values = window.read()
        # End program if user closes window or
        # presses the OK button
        print(event, values)
        if event == "Submit":
            break
        elif event == '-FLUX-':
            if values['-FLUX-']:
                window[f'-COL2-'].update(visible=False)
                window[f'-COL1-'].update(visible=True)
            else:
                window[f'-COL2-'].update(visible=True)
                window[f'-COL1-'].update(visible=False)
        elif event == "Cancel" or event == sg.WIN_CLOSED:
            cancelled = True
            break

    
    if cancelled == False:

        try:
            folder = values['-FOLDER-']

            if values['-FLUX-']:
                times_dict, times_list = process_times(values['-TIMES-'])
            else:
                times_dict = ''

            PDFs = get_PDFs(folder)

            if len(PDFs) == 0:
                raise Exception("Error: No PDFs found at selected location")

            BAR_MAX = len(PDFs)
        except Exception as e:
            window.close()
            raise Exception("Error in inputted information")

        layout = [[sg.Text('GC Data Processing Tool', font='Any 36', background_color='#0680BF')],
            [sg.Text("Processing... This may take a few minutes", background_color='#0680BF')],
            [sg.Text("", background_color='#0680BF')],
            [sg.ProgressBar(BAR_MAX, orientation='h', size=(40, 15), key='-PROG-', bar_color=('#38E210', '#FFFFFF'))]]

        window.close()
        window = sg.Window("GC", layout, margins = (80, 50), background_color='#0680BF')
        
        try:
            samples, standards, sample_rows = process_PDFs(samples, standards, sample_rows, PDFs, window, times_dict)
        except Exception as e:
            window.close()
            print(e)
            raise Exception("Error processing PDFs: ensure MARIA.isr report format is used")

        try:
            out = output_data(samples, standards, sample_rows, times_dict)
        except Exception as e:
            window.close()
            print(e)
            raise Exception("Error outputting results: ensure chosen location is valid")

    window.close()
    return 0