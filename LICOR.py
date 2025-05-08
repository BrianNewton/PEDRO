# Python v3.10.1
# matplotlib v3.5.1
# XlsxWriter v3.0.2

# Tool used to process raw LICOR data
# Enables users to prune raw data for ebullition events or inconsistent results

# For issues, suggestions or concerns contact btnewton@uwaterloo.ca


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

# Define global variable for the gas type being analyzed (CO2 or CH4)
LICOR_GAS = ''
# Dictionary of units for concentration of different gas types
gas_units = {
    'ch4': 'ppb',
    'co2': 'ppm',
    'n2o': 'ppb',
}
output_units = {
    'ch4': '(mg C m^-2 d^-1)',
    'co2': '(g C m^-2 d^-1)',
    'n2o': '(mg N2O m^-2 d^-1)'
}

# Flux object
class Flux:
    def __init__(self, name, light_or_dark, start_time, end_time, start_temp, end_temp, chamber_height, surface_area):

        light_regex = r"L|l|Light|light"
        dark_regex = r"D|d|Dark|dark"

        self.surface_area = surface_area

        if re.search(light_regex, light_or_dark):
            self.name = name + " " + "light"    # flux light name
        elif re.search(dark_regex, light_or_dark):
            self.name = name + " " + "dark"     # flux dark name
        else:
            self.name = name    # neither light nor dark name
            
        self.start_time = start_time    # LICOR flux start time
        self.end_time = end_time        # LICOR flux end time
        self.temp = (float(start_temp) + float(end_temp))/2 + 273.15    # average of start and end air temperature
        self.chamber_height = chamber_height

        self.original_length = 0    # original length of flux data
        self.data_loss = 0          # total percent of data set pruned

        self.times = []             # LICOR times
        self.pruned_times = []      # pruned LICOR times
        self.time_offsets = []      # offsets for each LICOR time index

        self.samples = []           # LICOR gas concentrations
        self.pruned_samples = []    # pruned LICOR concentrations
        self.sample_offsets = []    # offsets for each LICOR concentration index

        self.H2O = []
        self.pruned_H2O = []
        self.H2O_offsets = []

        # Lists for holding raw methane measurements. Will only be populated when LICOR_GAS == co2
        self.methane = []
        self.pruned_methane = []
        self.methane_offsets = []

        self.cuts = []      # every user data cut
            
        self.RSQ = 0        # R^2 for final rate calculation
        self.RoC = 0        # rate of change (concentration/minute)
        self.flux = 0       # final calculated flux


# parses raw LICOR data, as well as field data into flux objects
def input_data(field_data, licor_data):
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

    LICOR_time_regex = r"^TIME"
    LICOR_CH4_regex = r"CH4"
    LICOR_CO2_regex = r"CO2"
    LICOR_N2O_regex = r"N2O"
    LICOR_H2O_regex = r"H2O"

    time_regex = r"(\d*):(\d*):(\d*)"

    collar_index = 0
    l_or_d_index = 1
    start_time_index = 2
    end_time_index = 3
    start_temp_index = 4
    end_temp_index = 5
    chamber_height_index = 6
    surface_area_index = 7

    LICOR_time_index = 7
    LICOR_CH4_index = 10
    LICOR_CO2_index = 9
    LICOR_H2O_index = 8
    LICOR_N2O_index = 10

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
        if not re.search(time_regex, x[start_time_index], re.IGNORECASE):
            raise Exception("Error: Start time for flux {} is not formatted correctly. Please ensure all times are in 24h HH:MM:SS format.".format(x[collar_index]))
        if not re.search(time_regex, x[end_time_index], re.IGNORECASE):
            raise Exception("Error: End time for flux {} is not formatted correctly. Please ensure all times are in 24h HH:MM:SS format.".format(x[collar_index]))
        fluxes.append(Flux(x[collar_index], x[l_or_d_index], x[start_time_index], x[end_time_index], x[start_temp_index], x[end_temp_index], float(x[chamber_height_index]), float(x[surface_area_index])))
    f.close()

    # use the raw LICOR data as well as the parsed field data to obtain unpruned sets of times and concentrations for each flux
    f = open(licor_data, "r")
    x = next(f).replace('\t', ',').replace(';',',').split(",")
    while x[0] != "DATAH":
        x = next(f).replace('\t', ',').replace(';',',').split(",")
    print(x)
    # Determine the index (column) of the time, H2O, and sample gas
    for i in range(len(x)):
        if re.search(LICOR_time_regex, x[i], re.IGNORECASE):
            LICOR_time_index = i
        if re.search(LICOR_CH4_regex, x[i], re.IGNORECASE):
            LICOR_CH4_index = i
        if re.search(LICOR_CO2_regex, x[i], re.IGNORECASE):
            LICOR_CO2_index = i
        if re.search(LICOR_N2O_regex, x[i], re.IGNORECASE):
            LICOR_N2O_index = i
        if re.search(LICOR_H2O_regex, x[i], re.IGNORECASE):
            LICOR_H2O_index = i
    if LICOR_GAS == 'co2':
        index = LICOR_CO2_index
        methane_index = LICOR_CH4_index
    elif LICOR_GAS == 'ch4':
        index = LICOR_CH4_index
    else:
        index = LICOR_N2O_index

    in_flux = 0
    times = []
    samples = []
    H2O = []
    methane = []  # for raw methane measurements, only populated when LICOR_GAS == co2

    # go line by line in LICOR data, adding relevant times and concentrations to appropriate fluxes
    for flux in fluxes:
        f.seek(6)
        print(flux.name)
        for line in f:
            x = line.replace('\t', ',').replace(';', ',').split(',')
            if x[0] == "DATA":
                for k in range(len(x)):
                    x[k] = x[k].strip(' \t\n\r')
                if len(x) == 2:
                    continue
                if in_flux == 0:    # if current line in LICOR data isn't in a flux, check to see if it's the start time of the next flux
                    if x[LICOR_time_index] == flux.start_time:
                        in_flux = 1
                        start_seconds = float(x[1])
                        times.append(float(x[1]) - start_seconds)
                        samples.append(float(x[index]))
                        H2O.append(float(x[LICOR_H2O_index]))
                        if LICOR_GAS == 'co2':
                            methane.append(float(x[LICOR_CH4_index]))

                elif in_flux == 1:  # if current line in LICOR data is in a flux, append the time and gas concentration to flux
                    times.append(float(x[1]) - start_seconds)
                    samples.append(float(x[index]))
                    H2O.append(float(x[LICOR_H2O_index]))
                    if LICOR_GAS == 'co2':
                            methane.append(float(x[LICOR_CH4_index]))
                    if x[LICOR_time_index] == flux.end_time:  # if current line in LICOR data is the end time of the current flux, finalize times and concentrations sets and stop
                        flux.times = times
                        flux.samples = samples
                        flux.H2O = H2O
                        flux.methane = methane
                        flux.pruned_times = times
                        flux.pruned_samples = samples
                        flux.pruned_H2O = H2O
                        flux.pruned_methane = methane
                        flux.original_length = len(times)

                        times = []
                        samples = []
                        H2O = []
                        methane = []
                        in_flux = 0
                        start_seconds = 0 
                        continue
    f.close()
    return fluxes


def prune_list(data, time_L_index, time_R_index, time_L, time_R):
    """
    Shift the values in data by removing those within [time_L_index, time_R_index] while
    keeping the data continuous. Achieve continuity by shifting all datapoints after
    time_R_index by a delta equal to the gap between the last entry before and first entry
    after the cut.
    If the cut is at the boundary of the data (i.e. the cut isn't in the middle of the data),
    the offset (delta) will be zero.
    """
    if time_L_index and time_R_index:
        delta = data[time_L_index] - data[time_R_index]
    else:
        delta = 0

    return data[: time_L] + [datapoint + delta for datapoint in data[time_R+1 :]]


# process button press for plot
def on_press(event, i, fluxes, line_L, line_R, fig, ax1, ax2, ax3, cid):
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
            # track the upper and lower bounds of each cut in a list
            fluxes[i].cuts.append([time_L_index, time_R_index])

            # prune the time list to reflect the cut
            fluxes[i].pruned_times = prune_list(fluxes[i].pruned_times, time_L_index, time_R_index, time_L, time_R)

            # prune the samples, H2O, and optionally methane to reflect the cut
            fluxes[i].pruned_samples = prune_list(fluxes[i].pruned_samples, time_L_index, time_R_index, time_L, time_R)
            fluxes[i].pruned_H2O = prune_list(fluxes[i].pruned_H2O, time_L_index, time_R_index, time_L, time_R)
            if LICOR_GAS == 'co2':
                fluxes[i].pruned_methane = prune_list(fluxes[i].pruned_methane, time_L_index, time_R_index, time_L, time_R)

            # refresh the plot
            draw_plot(i, fluxes, fig, ax1, ax2, ax3, cid)

    # if r key pressed, reset data
    if event.key == 'r':
        fluxes[i].pruned_times = fluxes[i].times
        fluxes[i].pruned_samples = fluxes[i].samples
        fluxes[i].pruned_H2O = fluxes[i].H2O
        fluxes[i].pruned_methane = fluxes[i].methane
        fluxes[i].cuts = []
        draw_plot(i, fluxes, fig, ax1, ax2, ax3, cid)
    
    # right arrow key moves to the next flux, or exits if currently on the last flux
    if event.key == 'right':
        if i == len(fluxes) - 1:
            plt.ioff()
            plt.close('all')
            return 0
        else:
            ax1.clear()
            ax2.clear()
            if ax3:
                ax3.clear()
            draw_plot(i + 1, fluxes, fig, ax1, ax2, ax3, cid)
        
    # left arrow key moves to the previous flux, or does nothing if at the beginning
    if event.key == 'left':
        if i != 0:
            draw_plot(i - 1, fluxes, fig, ax1, ax2, ax3, cid)


#draw plot for i-th flux
def draw_plot(i, fluxes, fig, ax1, ax2, ax3, cid):
    """
    Function to add data from fluxes to the axes and draw figure.
    If LICOR_GAS != co2, ax3 will be None.
    """

    ax1.clear()
    m, b, R2 = linear_regression(fluxes[i].pruned_times, fluxes[i].pruned_samples)
    at = AnchoredText(
        r"$R^{2}$ = " + str(round(R2, 5)), prop=dict(size=12), frameon=True, loc='upper center')
    at.patch.set_boxstyle("round,pad=0.,rounding_size=0.2")
    at.patch.set_alpha(0.5)
    ax1.add_artist(at)
    ax1.plot(fluxes[i].pruned_times, fluxes[i].pruned_samples, linewidth = 2.0)
    ax1.grid(True)
    line_L = draggable_lines(ax1, fluxes[i].pruned_times[0], [fluxes[i].pruned_times[0], fluxes[i].pruned_times[-1]], ax1.get_ylim())   # left draggable boundary line
    line_R = draggable_lines(ax1, fluxes[i].pruned_times[-1], [fluxes[i].pruned_times[0], fluxes[i].pruned_times[-1]], ax1.get_ylim())     # right draggable boundary line

    # if currently on the last flux, change header information, otherwise set title to user controls
    if i == len(fluxes) - 1:
        ax1.set(title = fluxes[i].name +  "\nLast flux! Press right arrow to finish, enter to cut data, r to reset cuts\nUse the mouse to drag peak bounds")
    else:      
        ax1.set(title = fluxes[i].name + '\nUse arrow keys to navigate fluxes, enter to cut data, r to reset cuts\nUse the mouse to drag cut bounds')

    ax1.set(ylabel = f"{LICOR_GAS.upper()} concentration ({gas_units[LICOR_GAS]})")  # y axis label

    ax2.clear()
    ax2.plot(fluxes[i].pruned_times, fluxes[i].pruned_H2O, linewidth = 2.0)
    ax2.grid(True)

    ax2.set(ylabel = "H2O (ppm)")

    if LICOR_GAS == 'co2':
        ax3.clear()
        ax3.plot(fluxes[i].pruned_times, fluxes[i].pruned_methane, linewidth = 2.0)
        ax3.grid(True)

        ax3.set(xlabel = "Time (s)")                 # x axis label
        ax3.set(ylabel = "CH4 (ppb)")
    else:
        ax2.set(xlabel = "Time (s)")                 # x axis label

    fig.canvas.mpl_disconnect(cid)
    cid = fig.canvas.mpl_connect('key_press_event', lambda event: on_press(event, i, fluxes, line_L, line_R, fig, ax1, ax2, ax3, cid))   # connect key press event
    fig.canvas.draw()


# entry point for drawing the plots for the user to cut data
def prune(fluxes):
    i = 0
    # Create figure and axes for plots. If LICOR_GAS == co2, create
    # a third plot at the bottom for raw methane measurements.
    if LICOR_GAS == 'co2':
        fig, (ax1, ax2, ax3) = plt.subplots(3, 1)
        fig.set_size_inches(9, 7)
    else:
        fig, (ax1, ax2) = plt.subplots(2, 1)
        fig.set_size_inches(9,6)
        ax3 = None
    cid = ''
    plt.grid(True)
    plt.ion()
    draw_plot(i, fluxes, fig, ax1, ax2, ax3, cid)
    plt.show(block=False)
    while plt.get_fignums():
        fig.canvas.draw_idle()
        fig.canvas.start_event_loop(0.05)


# performs linear regression to generate linear gas concentration rate of change per minute
def flux_calculation(fluxes):
    for flux in fluxes:
        X = flux.pruned_times
        Y = flux.pruned_samples

        m, b, R2 = linear_regression(X, Y)

        # calculates flux depending on CO2 vs. CH4
        vol = flux.surface_area * flux.chamber_height * 1000

        flux.RSQ = R2
        flux.RoC = m * 60
        if LICOR_GAS == "co2":
            flux.flux = (flux.RoC*(vol/(0.0821*flux.temp))*(0.044*1440)/(flux.surface_area)*(12/44)/1000)
        elif LICOR_GAS == 'ch4':
            flux.flux = (flux.RoC*(vol/(0.0821*flux.temp))*(0.016*1440)/(flux.surface_area)*(12/16)/1000)
        else:  # LICOR_GAS == 'n2o'
            flux.flux = (flux.RoC*(vol/(0.0821*flux.temp))*(0.044*1440)/(flux.surface_area)/1000)


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
                flux.pruned_samples.insert(i, '')
                flux.pruned_H2O.insert(i, '')
                flux.pruned_methane.insert(i, '')
            cut_size += cut[1] + cut[0] + 1

        # calculate the offset at each index
        for k in range(len(flux.times)):
            if flux.pruned_times[k] != '':
                flux.time_offsets.append(flux.times[k] - flux.pruned_times[k])
                flux.sample_offsets.append(flux.samples[k] - flux.pruned_samples[k])
            else:
                flux.time_offsets.append(0)
                flux.sample_offsets.append(0)


# outputs data to excel file
def outputData(fluxes, site, date):
    out = tkinter.filedialog.asksaveasfilename(defaultextension='.xlsx')
    workbook = xlsxwriter.Workbook(out)

    # summary worksheet, displays R^2, rate of change, flux, chamber volume, air temp for each flux
    worksheet = workbook.add_worksheet("Summary")
    worksheet.write_row(0, 0, ["Site:", site])
    worksheet.write_row(1, 0, ["Date:", date])
    worksheet.write_column(3, 0, ["Flux name", '', "Chamber volume (L)", "Air temp (K)", '', "RSQ", f"Rate of change ({LICOR_GAS.upper()} [{gas_units[LICOR_GAS]}/min])", f"m ({LICOR_GAS.upper()} [{gas_units[LICOR_GAS]}/sec])", f"Flux of {LICOR_GAS.upper()} {output_units[LICOR_GAS]}", "Data loss (%)", "Surface moisture", "Surface temperature", "PAR"])
    for i in range(len(fluxes)):
        vol = fluxes[i].surface_area * fluxes[i].chamber_height * 1000
        worksheet.write_column(3, i + 1, [fluxes[i].name , '', vol, fluxes[i].temp, '', fluxes[i].RSQ, fluxes[i].RoC, fluxes[i].RoC/60, fluxes[i].flux, fluxes[i].data_loss])
        worksheet.set_column(i + 1, i + 1, len(fluxes[i].name ))
    worksheet.set_column(0, 0, len("Rate of change (CH4 [ppm/min])"))  # Gas type and units used only for length
    
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
        worksheet.write_row(2, 0, [f"Rate of change ({LICOR_GAS.upper()} [{gas_units[LICOR_GAS]}/min])", flux.RoC, '', '', "Air temp (K)", flux.temp])
        worksheet.write_row(3, 0, [f"Flux of {LICOR_GAS.upper()} {output_units[LICOR_GAS]}", flux.flux])
        worksheet.write_row(4, 0, ["Data loss (%)", flux.data_loss])

        worksheet.write(6, 0, "Original times (s)")
        worksheet.set_column(0, 0, len("Rate of change (CH4 [ppm/min])"))  # Gas type and units used only for length
        worksheet.write(6, 1, "Pruned times (s)")
        worksheet.set_column(1, 1, len("Pruned times (s)"))
        worksheet.write(6, 2, "Time offsets (s)")
        worksheet.set_column(2, 2, len("Time offsets (s)"))

        worksheet.write(6, 4, f"Original {LICOR_GAS.upper()} concentrations ({gas_units[LICOR_GAS]})")
        worksheet.set_column(4, 4, len(f"Original {LICOR_GAS.upper()} concentrations ({gas_units[LICOR_GAS]})"))
        worksheet.write(6, 5, f"Pruned {LICOR_GAS.upper()} concentrations ({gas_units[LICOR_GAS]})")
        worksheet.set_column(5, 5, len(f"Pruned {LICOR_GAS.upper()} concentrations ({gas_units[LICOR_GAS]})"))
        worksheet.write(6, 6, f"{LICOR_GAS.upper()} concentration offsets ({gas_units[LICOR_GAS]})")
        worksheet.set_column(6, 6, len(f"{LICOR_GAS.upper()} concentration offsets ({gas_units[LICOR_GAS]})"))

        worksheet.write(6, 8, "Original H2O (ppm)")
        worksheet.set_column(8, 8, len("Original H2O (ppm)"))
        worksheet.write(6, 9, "Pruned H2O (ppm)")
        worksheet.set_column(9, 9, len("Pruned H2O (ppm)"))

        for i in range(len(flux.times)):
            worksheet.write_row(i + 7, 0, [flux.times[i], flux.pruned_times[i], flux.time_offsets[i], '', flux.samples[i], flux.pruned_samples[i], flux.sample_offsets[i], '', flux.H2O[i], flux.pruned_H2O[i]])

        # generate chart showing cut values compared to kept values with offsets
        chart1 = workbook.add_chart({'type': 'line'})
        chart1.add_series({'values' : '=\'%s\'!E8:E%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A8:A%i'%(flux.name, len(flux.samples) + 9), 'name': 'Cut values', 'line': {'color': 'red'}})
        chart1.add_series({'values' : '=\'%s\'!F8:F%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A9:A%i'%(flux.name, len(flux.samples) + 9), 'name': 'Kept values', 'line': {'color': 'green'}})
        chart1.set_y_axis({'interval_unit': 10, 'interval_tick': 2, 'name': f"{LICOR_GAS.upper()} concentration ({gas_units[LICOR_GAS]})"})
        chart1.set_x_axis({'name': 'Time (s)'})
        chart1.set_title({'name' : 'Concentration vs. Time'})
        chart1.set_size({'width': 800, 'height': 600})
        worksheet.insert_chart('L8', chart1)

        chart2 = workbook.add_chart({'type': 'line'})
        chart2.add_series({'values' : '=\'%s\'!I8:I%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A8:A%i'%(flux.name, len(flux.samples) + 9), 'name': 'Cut values', 'line': {'color': 'red'}})
        chart2.add_series({'values' : '=\'%s\'!J8:J%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A9:A%i'%(flux.name, len(flux.samples) + 9), 'name': 'Kept values', 'line': {'color': 'green'}})
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


def LICOR():

    layout = [[sg.Text('LICOR Flux Data Processing Tool', font='Any 36', background_color='#DF954A')],
        [sg.Text("", background_color='#DF954A')],
        [sg.Text('Field data file: (.csv, .txt)', size=(21, 1), background_color='#DF954A'), sg.Input(key='-FIELD-'), sg.FileBrowse()],
        [sg.Text('LICOR data file: (.csv, .txt)', size=(21, 1), background_color='#DF954A'), sg.Input(key='-LICOR-'), sg.FileBrowse()],
        [
            sg.Text('Gas to analyze:', size=(15, 1), background_color='#DF954A'),
            sg.Radio('CO2', 'RADIO2', enable_events=True, default=False, key='-CO2-', background_color='#DF954A'),
            sg.Radio('CH4', 'RADIO2', enable_events=True, default=True, key='-CH4-', background_color='#DF954A'),
            sg.Radio('N2O', 'RADIO2', enable_events=True, default=False, key='-N2O-', background_color='#DF954A'),
        ],
        [sg.Text("Site name:", size=(15, 1), background_color='#DF954A'), sg.InputText(key='-SITE-')],
        [sg.Text("Date:", size=(15, 1), background_color='#DF954A'), sg.InputText(key='-DATE-')],
        [sg.Text("", background_color='#DF954A')],
        [sg.Submit(), sg.Cancel()]]


    # Create the window
    window = sg.Window("LICOR", layout, margins=(80, 50), background_color='#DF954A')
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
        # Set the global LICOR_GAS variable based on the radio button selected
        global LICOR_GAS
        if values['-CO2-']:
            LICOR_GAS = 'co2'
        elif values['-CH4-']:
            LICOR_GAS = 'ch4'
        else:  # values['-CH4-'] is True
            LICOR_GAS = 'n2o'

        try:
            field_data = values['-FIELD-']
            licor_data = values['-LICOR-']
            site = values['-SITE-']
            date = values['-DATE-']
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise e

        try:
            print("Reading input files")
            fluxes = input_data(field_data, licor_data)

            print("Pruning data")
            prune(fluxes)

            print("Calculating fluxes")
            flux_calculation(fluxes)

            print("Calculating offsets")
            offsets(fluxes)

            print("Outputting data")
            out = outputData(fluxes, site, date)

        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise e

    window.close()
    return 0