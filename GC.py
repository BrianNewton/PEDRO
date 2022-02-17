import os
import glob
import xlsxwriter
import re
import tabula
import pandas as pd

PDFs = glob.glob("*.pdf")


        

sample_rows = {}
samples = {}
#methane, CO2, O2, N2, NO2
standards = {'1ppm': {'Methane': [], 'Carbon Dioxide': [], 'Oxygen': [], 'Nitrogen': [], 'Nitrous Oxide': [], 'File': []},
                '5ppm': {'Methane': [], 'Carbon Dioxide': [], 'Oxygen': [], 'Nitrogen': [], 'Nitrous Oxide': [], 'File': []}, 
                '50ppm': {'Methane': [], 'Carbon Dioxide': [], 'Oxygen': [], 'Nitrogen': [], 'Nitrous Oxide': [], 'File': []}}
row_index = 1

standard_regex = r"(\d*)\s?ppm"
sample_regex = r"([SHLC\d]*)-*\s*([\d.]*\s?min)"
date_ran_regex = r"[a-zA-Z]{3}\d{2}"
concentration_regex = r"-?\d*\.\d*"

#PDFs = ['Jan20_1_2022_20220127110303.pdf', 'Jan20_2_2022_20220127110215.pdf']

#PDFs = ['CC,SH June 10 2021_20220127110950.pdf']

for file in PDFs:

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
            

            if re.search (sample_regex, row['Sample Name']):
                sample_name = re.search(sample_regex, row['Sample Name'])[1]
                if sample_name not in sample_rows:
                    sample_rows[sample_name] = row_index
                    row_index += 3


                sample_time = re.search(sample_regex, row['Sample Name'])[2].replace(" ", "")
                date = row['Sample ID']

                if date in samples:
                    if sample_name in samples[date]:
                        samples[date][sample_name][sample_time][gasses[i]] = float(re.search(concentration_regex, row['Conc. Unit'])[0])
                        samples[date][sample_name][sample_time]['File'] = file
                    else:
                        samples[date][sample_name] = {'0min' : {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''},
                                                        '3.5min': {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''},
                                                        '7min': {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''}}
                        samples[date][sample_name][sample_time][gasses[i]] = float(re.search(concentration_regex, row['Conc. Unit'])[0])
                        samples[date][sample_name][sample_time]['File'] = file
                else:
                    samples[date] = {sample_name: {'0min' : {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''},
                                                        '3.5min': {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''},
                                                        '7min': {'Methane': '', 'Carbon Dioxide': '', 'Oxygen': '', 'Nitrogen': '', 'Nitrous Oxide': '', 'File': ''}}}
                    samples[date][sample_name][sample_time][gasses[i]] = float(re.search(concentration_regex, row['Conc. Unit'])[0])
                    samples[date][sample_name][sample_time]['File'] = file



workbook = xlsxwriter.Workbook('results.xlsx')
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
        worksheet.write_row(sample_rows[sample_name], 0, [sample_name, "0min", sample['0min']['Methane'], sample['0min']['Carbon Dioxide'], sample['0min']['Oxygen'], sample['0min']['Nitrogen'], sample['0min']['Nitrous Oxide'], sample['0min']['File']])
        worksheet.write_row(sample_rows[sample_name] + 1, 1, ["3.5min", sample["3.5min"]['Methane'], sample["3.5min"]['Carbon Dioxide'], sample["3.5min"]['Oxygen'], sample["3.5min"]['Nitrogen'], sample["3.5min"]['Nitrous Oxide'], sample['3.5min']['File']])
        worksheet.write_row(sample_rows[sample_name] + 2, 1, ["7min", sample["7min"]['Methane'], sample["7min"]['Carbon Dioxide'], sample["7min"]['Oxygen'], sample["7min"]['Nitrogen'], sample["7min"]['Nitrous Oxide'], sample['7min']['File']])

workbook.close()