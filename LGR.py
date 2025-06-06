from math import floor
import sys
import re
import matplotlib.pyplot as plt
from matplotlib.offsetbox import AnchoredText
import os
import tkinter
import tkinter.filedialog
import xlsxwriter
import PySimpleGUI as sg
import traceback

from utils import draggable_lines, linear_regression
    

# Flux object
class Flux:
    def __init__(self, name, light_or_dark, start_time, end_time, chamber_height, surface_area):

        light_regex = r"L|l|Light|light"
        dark_regex = r"D|d|Dark|dark"

        self.surface_area = surface_area

        if re.search(light_regex, light_or_dark):
            self.name = name + " " + "light"    # flux light name
        elif re.search(dark_regex, light_or_dark):
            self.name = name + " " + "dark"     # flux dark name
        else:
            self.name = name    # neither light nor dark name
            
        self.start_time = start_time    # LGR flux start time
        self.end_time = end_time        # LGR flux end time
        self.chamber_height = chamber_height

        self.original_length = 0    # original length of flux data
        self.data_loss = 0          # total percent of data set pruned

        self.times = []             # LGR times
        self.pruned_times = []      # pruned LGR times
        self.time_offsets = []      # offsets for each LGR time index

        self.CH4 = []               # LGR gas concentrations
        self.pruned_CH4 = []        # pruned LGR concentrations
        self.CH4_offsets = []       # offsets for each LGR concentration index

        self.H2O = []
        self.pruned_H2O = []
        self.H2O_offsets = []

        self.temp = 0               # Average temperature from LGR

        self.cuts = []      # every user data cut
            
        self.RSQ = 0        # R^2 for final rate calculation
        self.RoC = 0        # rate of change (concentration/minute)
        self.flux = 0       # final calculated flux


# parses raw LGR data, as well as field data into flux objects
def input_data(field_data, LGR_data, CO2_or_CH4):
    fluxes = [] # set of flux objects

    # find relevant data indices in field data file
    collar_regex = r"collar$|name"
    l_or_d_regex = r"light[ -_]or[ -_]dark"
    start_time_regex = r"start[ -_]time"
    end_time_regex = r"end[ -_]time"
    start_temp_regex = r"start[ -_]temp"
    end_temp_regex = r"end[ -_]temp"
    chamber_height_regex = r"chamber[ -_]height[ ]?(m)"
    surface_area_regex = r"surface[ -_]area[ ]?(m^2)"

    LGR_time_regex = r"Time"
    LGR_CH4_regex = r"\[CH4\]_ppm$"
    LGR_CO2_regex = r"\[CO2\]_ppm$"
    LGR_temp_regex = r"AmbT_C$"
    LGR_H2O_regex = r"\[H2O\]_ppm$"

    time_regex_field = r"(\d*):(\d*):(\d*)"
    time_regex_LGR = r"(\d*):(\d*):(\d*).(\d*)"

    collar_index = 0
    l_or_d_index = 1
    start_time_index = 2
    end_time_index = 3
    chamber_height_index = 4
    surface_area_index = 5

    LGR_time_index = 0
    LGR_CH4_index = 1
    LGR_CO2_index = 9
    LGR_H2O_index = 7
    LGR_temp_index = 19

    # for each flux, obtain from field data file the name, start and end times
    f = open(field_data, "r")
    x = next(f).replace('\t', ',').replace(';',',').split(",")
    for i in range(len(x)):
        x[i] = x[i].strip(' \t\n\r')
        if re.search(collar_regex, x[i], re.IGNORECASE):
            collar_index = i
        if re.search(l_or_d_regex, x[i], re.IGNORECASE):
            l_or_d_index = i
        if re.search(start_time_regex, x[i], re.IGNORECASE):
            start_time_index = i 
        if re.search(end_time_regex, x[i], re.IGNORECASE):
            end_time_index = i
        if re.search(start_temp_regex, x[i], re.IGNORECASE):
            start_temp_index = i
        if re.search(end_temp_regex, x[i], re.IGNORECASE):
            end_temp_index = i
        if re.search(chamber_height_regex, x[i], re.IGNORECASE):
            chamber_height_index = i
        if re.search(surface_area_regex, x[i], re.IGNORECASE):
            surface_area_index = i

    for line in f:
        x = line.replace('\t', ',').replace(';', ',').split(",")
        for i in range(len(x)):
            x[i] = x[i].strip(' \t\n\r')
        if not re.search(time_regex_field, x[start_time_index], re.IGNORECASE):
            raise Exception("Error: Start time for flux {} is not formatted correctly. Please ensure all times are in 24h HH:MM:SS format.".format(x[collar_index]))
        if not re.search(time_regex_field, x[end_time_index], re.IGNORECASE):
            raise Exception("Error: End time for flux {} is not formatted correctly. Please ensure all times are in 24h HH:MM:SS format.".format(x[collar_index]))
        fluxes.append(Flux(x[collar_index], x[l_or_d_index], x[start_time_index], x[end_time_index], float(x[chamber_height_index]), float(x[surface_area_index])))
    f.close()

    # use the raw LGR data as well as the parsed field data to obtain unpruned sets of times and concentrations for each flux
    f = open(LGR_data, "r")
    x = next(f).replace('\t', ',').replace(';',',').split(",")
    for i in range(len(x)):
        if re.search(LGR_time_regex, x[i], re.IGNORECASE):
            LGR_time_index = i
        if re.search(LGR_CH4_regex, x[i], re.IGNORECASE):
            LGR_CH4_index = i
        if re.search(LGR_CO2_regex, x[i], re.IGNORECASE):
            LGR_CO2_index = i
        if re.search(LGR_temp_regex, x[i], re.IGNORECASE):
            LGR_temp_index = i
        if re.search(LGR_H2O_regex, x[i], re.IGNORECASE):
            LGR_H2O_regex = i

    in_flux = 0
    times = []
    CH4 = []
    H2O = []
    temps = []
    if CO2_or_CH4.lower() == 'co2':
        index = LGR_CO2_index
    else:
        index = LGR_CH4_index

    # go line by line in LGR data, adding relevant times and concentrations to approrpiate fluxes
    for flux in fluxes:
        f.seek(6)
        print(flux.name)
        for line in f:
            x = line.replace('\t', ',').replace(';', ',').split(',')
            for k in range(len(x)):
                x[k] = x[k].strip(' \t\n\r')
            time_match = re.search(time_regex_LGR, x[LGR_time_index])
            if not time_match:
                continue
            else:
                time_seconds = "{}:{}:{}".format(time_match[1], time_match[2], time_match[3])

                if in_flux == 0: # if current line in LGR data isn't in a flux, check to see if it's the start time of the next flux
                    if time_seconds == flux.start_time:
                        in_flux = 1
                        start_seconds = float(time_match[1])*3600 + float(time_match[2])*60 + float(time_match[3])
                        times.append(0)
                        if CO2_or_CH4:
                            CH4.append(float(x[index])) #CO2
                        else:
                            CH4.append(float(x[index])*1000) # CH4, LGR reports in ppm, want ppb
                        H2O.append(float(x[LGR_H2O_index]))
                        temps.append(float(x[LGR_temp_index]))
                elif in_flux == 1: # if current line in LGR data is in a flux, append the time and gas concentration to flux
                    times.append(float(time_match[1])*3600 + float(time_match[2])*60 + float(time_match[3]) - start_seconds)
                    if CO2_or_CH4:
                        CH4.append(float(x[index])) #CO2
                    else:
                        CH4.append(float(x[index])*1000) # CH4, LGR reports in ppm, want ppb
                    H2O.append(float(x[LGR_H2O_index]))
                    temps.append(float(x[LGR_temp_index]))
                    if time_seconds == flux.end_time: # if current line in LGR data is the end time of the current flux, finalize times and concentrations sets and stop
                        
                        flux.times = times
                        flux.CH4 = CH4
                        flux.H2O = H2O
                        flux.pruned_times = times
                        flux.pruned_CH4 = CH4
                        flux.pruned_H2O = H2O
                        flux.original_length = len(times)
                        flux.temp = sum(temps)/len(temps)

                        times = []
                        CH4 = []
                        H2O = []
                        temps = []
                        in_flux = 0
                        start_seconds = 0

                        continue

    f.close()
    return fluxes


# process button press for plot
def on_press(event, i, fluxes, line_L, line_R, fig, ax1, ax2, cid, CO2_or_CH4):
    sys.stdout.flush()

    # if enter key pressed, cut data according to currently set cut bounds
    if event.key == 'enter':

        # get cut indices from left and right lines
        time_L = floor(line_L.line.get_xdata()[0])
        time_R = floor(line_R.line.get_xdata()[0])

        time_L_index = fluxes[i].pruned_times.index(time_L)
        time_R_index = fluxes[i].pruned_times.index(time_R)
        
        if (time_R_index - time_L_index + 1) >= len(fluxes[i].pruned_times):
            print("Error! Can't cut entire data set, please narrow your selection with the two red cursors")
        else:
            # if both indices exist, set CH4 delta as the gap between the last entry before and first entry after the cut
            # this is for the sake of maintaining a linear relationship, it adds a fixed offset to each entry after the cut
            # if the cut is at the boundary of the data, the offset will be set to zero (i.e. the cut isn't in the middle of the data)
            if time_L_index and time_R_index:
                CH4_delta = fluxes[i].pruned_CH4[time_L_index] - fluxes[i].pruned_CH4[time_R_index]
            else:
                CH4_delta = 0

            # obtain time offset as the time elapsed by the cut
            # for the sake of maintaining a linear relationship, adds a fixed offset
            time_delta = time_R - time_L
            fluxes[i].cuts.append([time_L_index, time_R_index])

            times = []
            CH4 = []
            H2O = []

            # if entry is not in the cut, add it to a new set
            for k in range(len(fluxes[i].pruned_times)):
                if fluxes[i].pruned_times[k] < time_L:
                    times.append(fluxes[i].pruned_times[k])
                    CH4.append(fluxes[i].pruned_CH4[k])
                    H2O.append(fluxes[i].pruned_H2O[k])
                if fluxes[i].pruned_times[k] > time_R:
                    times.append(fluxes[i].pruned_times[k] - time_delta)
                    CH4.append(fluxes[i].pruned_CH4[k] + CH4_delta)
                    H2O.append(fluxes[i].pruned_H2O[k])
            
            # update flux pruned sets with cut sets
            fluxes[i].pruned_times = times
            fluxes[i].pruned_CH4 = CH4
            fluxes[i].pruned_H2O = H2O

            # refresh the plot
            draw_plot(i, fluxes, fig, ax1, ax2, cid, CO2_or_CH4)

    # if r key pressed, reset data
    if event.key == 'r':
        fluxes[i].pruned_times = fluxes[i].times
        fluxes[i].pruned_CH4 = fluxes[i].CH4
        fluxes[i].pruned_H2O = fluxes[i].H2O
        fluxes[i].cuts = []
        draw_plot(i, fluxes, fig, ax1, ax2, cid, CO2_or_CH4)
    
    # right arrow key moves to the next flux, or exits if currently on the last flux
    if event.key == 'right':
        if i == len(fluxes) - 1:
            plt.ioff()
            plt.close('all')
            return 0
        else:
            draw_plot(i + 1, fluxes, fig, ax1, ax2, cid, CO2_or_CH4)
        
    # left arrow key moves to the previous flux, or does nothing if at the beginning
    if event.key == 'left':
        if i != 0:
            draw_plot(i - 1, fluxes, fig, ax1, ax2, cid, CO2_or_CH4)


#draw plot for i-th flux
def draw_plot(i, fluxes, fig, ax1, ax2, cid, CO2_or_CH4):

    ax1.clear()
    m, b, R2 = linear_regression(fluxes[i].pruned_times, fluxes[i].pruned_CH4)
    at = AnchoredText(
        r"$R^{2}$ = " + str(round(R2, 5)), prop=dict(size=15), frameon=True, loc='upper center')
    at.patch.set_boxstyle("round,pad=0.,rounding_size=0.2")
    at.patch.set_alpha(0.5)
    ax1.add_artist(at)
    ax1.plot(fluxes[i].pruned_times, fluxes[i].pruned_CH4, linewidth = 2.0)
    ax1.grid(True)
    line_L = draggable_lines(ax1, fluxes[i].pruned_times[0], [fluxes[i].pruned_times[0], fluxes[i].pruned_times[-1]], ax1.get_ylim())   # left draggable boundary line
    line_R = draggable_lines(ax1, fluxes[i].pruned_times[-1], [fluxes[i].pruned_times[0], fluxes[i].pruned_times[-1]], ax1.get_ylim())     # right draggable boundary line

    # if currently on the last flux, change header information, otherwise set title to user controls
    if i == len(fluxes) - 1:
        ax1.set(title = fluxes[i].name +  "\nLast flux! Press right arrow to finish, enter to cut data, r to reset cuts\nUse the mouse to drag peak bounds")
    else:      
        ax1.set(title = fluxes[i].name + '\nUse arrow keys to navigate fluxes, enter to cut data, r to reset cuts\nUse the mouse to drag cut bounds')

    if CO2_or_CH4 == 'co2':
        ax1.set(ylabel = "CO2 concentration (ppm)")  # y axis label
    else:
        ax1.set(ylabel = "CH4 concentration (ppb)")
    
    ax2.clear()
    ax2.plot(fluxes[i].pruned_times, fluxes[i].pruned_H2O, linewidth = 2.0)
    ax2.grid(True)

    ax2.set(xlabel = "Time (s)")
    ax2.set(ylabel = "H2O (ppm)")

    fig.canvas.mpl_disconnect(cid)
    cid = fig.canvas.mpl_connect('key_press_event', lambda event: on_press(event, i, fluxes, line_L, line_R, fig, ax1, ax2, cid, CO2_or_CH4))   # connect key press event  
    fig.canvas.draw()


# entry point for drawing the plots for the user to cut data
def prune(fluxes, CO2_or_CH4):
    i = 0
    fig, (ax1, ax2) = plt.subplots(2, 1)
    fig.set_size_inches(9,6)
    cid = ''
    plt.grid(True)
    plt.ion()
    draw_plot(i, fluxes, fig, ax1, ax2, cid, CO2_or_CH4)
    plt.show(block=False)
    while plt.get_fignums():
        fig.canvas.draw_idle()
        fig.canvas.start_event_loop(0.05)


# performs linear regression to generate linear gas concentration rate of change per minute
def flux_calculation(fluxes, CO2_or_CH4):
    for flux in fluxes:
        X = flux.pruned_times
        Y = flux.pruned_CH4

        m, b, R2 = linear_regression(X, Y)

        # calculates flux depending on CO2 vs. CH4
        vol = flux.surface_area * flux.chamber_height * 1000

        flux.RSQ = R2
        flux.RoC = m * 60
        if CO2_or_CH4.lower() == "co2":
            flux.flux = (flux.RoC*(vol/(0.0821*flux.temp))*(0.044*1440)/(flux.surface_area)*(12/44)/1000)
        else:
            flux.flux = (flux.RoC*(vol/(0.0821*flux.temp))*(0.016*1440)/(flux.surface_area)*(12/16)/1000000)


# calculates cut offsets for the sake of reporting
def offsets(fluxes):

    for flux in fluxes:

        flux.max_time = max(flux.pruned_times)
        flux.min_time = min(flux.pruned_times)

        # for each cut, rebuild pruned data sets with '' for cut indices
        cut_size = 0
        flux.data_loss = 100 - round(len(flux.pruned_times) / len(flux.times) * 100, 2)
        for cut in flux.cuts:
            for i in range(cut[0] + cut_size, cut[1] + 1 + cut_size):
                flux.pruned_times.insert(i, '')
                flux.pruned_CH4.insert(i, '')
                flux.pruned_H2O.insert(i, '')
            cut_size += cut[1] + cut[0] + 1

        # calculate the offset at each index
        for k in range(len(flux.times)):
            if flux.pruned_times[k] != '':
                flux.time_offsets.append(flux.times[k] - flux.pruned_times[k])
                flux.CH4_offsets.append(flux.CH4[k] - flux.pruned_CH4[k])
            else:
                flux.time_offsets.append(0)
                flux.CH4_offsets.append(0)


# outputs data to excel file
def outputData(fluxes, site, date, CO2_or_CH4):
    out = tkinter.filedialog.asksaveasfilename(defaultextension='.xlsx')
    workbook = xlsxwriter.Workbook(out)

    # summary worksheet, displays R^2, rate of change, flux, chamber volume, air temp for each flux
    worksheet = workbook.add_worksheet("Summary")
    worksheet.write_row(0, 0, ["Site:", site])
    worksheet.write_row(1, 0, ["Date:", date])
    if CO2_or_CH4.lower() == "co2":
        worksheet.write_column(3, 0, ["Flux name", '', "Chamber volume (L)", "Air temp (C)", '',  "RSQ", "Rate of change (CO2 [ppm/min])", "m (CO2 [ppm/sec])", "Flux of CO2 (g C m^-2 d^-1", "Data loss (%)", "Surface moisture", "Surface temperature", "PAR"])
    else:
        worksheet.write_column(3, 0, ["Flux name", '', "Chamber volume (L)", "Air temp (C)", '', "RSQ", "Rate of change (CH4 [ppb/min])", "m (CH4 [ppb/sec])", "Flux of CH4 (g C m^-2 d^-1", "Data loss (%)", "Surface moisture", "Surface temperature", "PAR"])
    for i in range(len(fluxes)):
        vol = fluxes[i].surface_area * fluxes[i].chamber_height * 1000
        worksheet.write_column(3, i + 1, [fluxes[i].name , '', vol, fluxes[i].temp, '', fluxes[i].RSQ, fluxes[i].RoC, fluxes[i].RoC/60, fluxes[i].flux, fluxes[i].data_loss])
        worksheet.set_column(i + 1, i + 1, len(fluxes[i].name ))
    worksheet.set_column(0, 0, len("Rate of change (CH4 [ppm/min])"))
    
    # create page for each flux, pages give a detailed breakdown of each fluxes data sets as well as the values that have been cut
    worksheets = {}
    for flux in fluxes:
        vol = flux.surface_area*flux.chamber_height * 1000
        if flux.name not in worksheets:
            worksheet = workbook.add_worksheet(flux.name)
            worksheets[flux.name] = 1
        else:
            worksheets[flux.name] += 1
            worksheet = workbook.add_worksheet(flux.name + " (%s)" %(str(worksheets[flux.name])))
        worksheet.write_row(0, 0, ["Name", flux.name])
        worksheet.write_row(1, 0, ["RSQ", flux.RSQ, '', '', "Chamber volume (L)", vol])
        if CO2_or_CH4.lower() == 'co2':
            worksheet.write_row(2, 0, ["Rate of change (CO2 [ppm/min]", flux.RoC, '', '', "Air temp (C)", flux.temp])
            worksheet.write_row(3, 0, ["Flux of CO2 (g C m^-2 d^-1)", flux.flux])
        else:
            worksheet.write_row(2, 0, ["Rate of change (CH4 [ppb/min]", flux.RoC, '', '', "Air temp (C)", flux.temp])
            worksheet.write_row(3, 0, ["Flux of CH4 (g C m^-2 d^-1)", flux.flux])
        worksheet.write_row(4, 0, ["Data loss (%)", flux.data_loss])

        worksheet.write(6, 0, "Original times (s)")
        worksheet.set_column(0, 0, len("Rate of change (CH4 [ppm/min])"))
        worksheet.write(6, 1, "Pruned times (s)")
        worksheet.set_column(1, 1, len("Pruned times (s)"))
        worksheet.write(6, 2, "Time offsets (s)")
        worksheet.set_column(2, 2, len("Time offsets (s)"))

        if CO2_or_CH4.lower() == "co2":
            worksheet.write(6, 4, "Original CO2 concentrations (ppm)")
            worksheet.set_column(4, 4, len("Original CO2 concentrations (ppm)"))
            worksheet.write(6, 5, "Pruned CO2 concentrations (ppm)")
            worksheet.set_column(5, 5, len("Pruned CO2 concentrations (ppm)"))
            worksheet.write(6, 6, "CO2 concentration offsets (ppm)")
            worksheet.set_column(6, 6, len("CO2 concentration offsets (ppm)"))
        else: 
            worksheet.write(6, 4, "Original CH4 concentrations (ppb)")
            worksheet.set_column(4, 4, len("Original CH4 concentrations (ppb)"))
            worksheet.write(6, 5, "Pruned CH4 concentrations (ppb)")
            worksheet.set_column(5, 5, len("Pruned CH4 concentrations (ppb)"))
            worksheet.write(6, 6, "CH4 concentration offsets (ppb)")
            worksheet.set_column(6, 6, len("CH4 concentration offsets (ppb)"))

        worksheet.write(6, 8, "Original H2O (ppm)")
        worksheet.set_column(8, 8, len("Original H2O (ppm)"))
        worksheet.write(6, 9, "Pruned H2O (ppm)")
        worksheet.set_column(9, 9, len("Pruned H2O (ppm)"))

        for i in range(len(flux.times)):
            worksheet.write_row(i + 7, 0, [flux.times[i], flux.pruned_times[i], flux.time_offsets[i], '', flux.CH4[i], flux.pruned_CH4[i], flux.CH4_offsets[i], '', flux.H2O[i], flux.pruned_H2O[i]])

        # generate chart showing cut values compared to kept values with offsets
        chart1 = workbook.add_chart({'type': 'line'})
        chart1.add_series({'values' : '=\'%s\'!E8:E%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A8:A%i'%(flux.name, len(flux.CH4) + 9), 'name': 'Cut values', 'line': {'color': 'red'}})
        chart1.add_series({'values' : '=\'%s\'!F8:F%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A9:A%i'%(flux.name, len(flux.CH4) + 9), 'name': 'Kept values', 'line': {'color': 'green'}})
        if CO2_or_CH4.lower() == "co2":
            chart1.set_y_axis({'interval_unit': 10, 'interval_tick': 2, 'name': 'CO2 concentration (ppm)'})
        else:
            chart1.set_y_axis({'interval_unit': 10, 'interval_tick': 2, 'name': 'CH4 concentration (ppb)'})
        chart1.set_x_axis({'name': 'Time (s)'})
        chart1.set_title({'name' : 'Concentration vs. Time'})
        chart1.set_size({'width': 800, 'height': 600})
        worksheet.insert_chart('L8', chart1)

        chart2 = workbook.add_chart({'type': 'line'})
        chart2.add_series({'values' : '=\'%s\'!I8:I%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A8:A%i'%(flux.name, len(flux.CH4) + 9), 'name': 'Cut values', 'line': {'color': 'red'}})
        chart2.add_series({'values' : '=\'%s\'!J8:J%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A9:A%i'%(flux.name, len(flux.CH4) + 9), 'name': 'Kept values', 'line': {'color': 'green'}})
        chart2.set_y_axis({'interval_unit': 10, 'interval_tick': 2, 'name': 'H2O (ppm)'})
        chart2.set_x_axis({'name': 'Time (s)'})
        chart2.set_title({'name': 'Humidity vs. Time'})
        chart2.set_size({'width': 800, 'height': 600})
        worksheet.insert_chart('L40', chart2)
    
    workbook.close()    
    return out


#########################################################################################################################
######################################## script execution starts here! ##################################################
#########################################################################################################################


def LGR():

    layout = [[sg.Text('LGR Flux Data Processing Tool', font='Any 36', background_color='#00A1A0')],
        [sg.Text("", background_color='#00A1A0')],
        [sg.Text('Field data file: (.csv, .txt)', size=(21, 1), background_color='#00A1A0'), sg.Input(key='-FIELD-'), sg.FileBrowse()],
        [sg.Text('LGR data file: (.csv, .txt)', size=(21, 1), background_color='#00A1A0'), sg.Input(key='-LGR-'), sg.FileBrowse()],
        [sg.Text('Gas to analyze:', size=(15, 1), background_color='#00A1A0'), sg.Radio('CO2', 'RADIO2', enable_events=True, default=False, key='-CO2-', background_color='#00A1A0'), sg.Radio('CH4', 'RADIO2',enable_events=True, default=True, key='-CH4-', background_color='#00A1A0')],
        [sg.Text("Site name:", size=(15, 1), background_color='#00A1A0'), sg.InputText(key='-SITE-')],
        [sg.Text("Date:", size=(15, 1), background_color='#00A1A0'), sg.InputText(key='-DATE-')],
        [sg.Text("", background_color='#00A1A0')],
        [sg.Submit(), sg.Cancel()]]


    # Create the window
    window = sg.Window("LGR", layout, margins=(80, 50), background_color='#00A1A0')
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
        elif event in '-SQUARE-':
            window[f'-COL1-'].update(visible=False)
            window[f'-COL2-'].update(visible=True)
            window[f'-SQUARE-'].update(True)
        elif event in '-CIRCLE-':
            window[f'-COL2-'].update(visible=False)
            window[f'-COL1-'].update(visible=True)
            window[f'-CIRCLE-'].update(True)

    if cancelled == False:
        if values['-CO2-']:
            CO2_or_CH4 = 'co2'
        else:
            CO2_or_CH4 = 'ch4'

        try:
            field_data = values['-FIELD-']
            LGR_data = values['-LGR-']
            site = values['-SITE-']
            date = values['-DATE-']
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error in inputted information")

        try:
            print("Reading input files")
            fluxes = input_data(field_data, LGR_data, CO2_or_CH4)

            print("Pruning data")
            prune(fluxes, CO2_or_CH4)

            print("Calculating fluxes")
            flux_calculation(fluxes, CO2_or_CH4)

            print("Calculating offsets")
            offsets(fluxes)

            print("Outputting data")
            out = outputData(fluxes, site, date, CO2_or_CH4)

        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise e
        
    window.close()
    return 0

