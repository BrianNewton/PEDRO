import traceback
import glob
import matplotlib.pyplot as plt
import matplotlib.lines as lines
from math import floor
import sys
import xlsxwriter
import tkinter
import tkinter.filedialog
from matplotlib.offsetbox import AnchoredText
from matplotlib.widgets import TextBox
import PySimpleGUI as sg
import os
import re

from utils import draggable_lines, linear_regression

class Flux:
    def __init__(self, name, CO2, times, PARs, temps, volume):
        self.name = name
        self.CO2 = CO2
        self.times = times
        self.temps = temps
        self.PARs = PARs
        self.temp = sum(temps)/len(temps)
        self.PAR = sum(PARs)/len(PARs)
        self.volume = volume

        self.pruned_CO2 = CO2
        self.pruned_times = times

        self.time_offsets = []
        self.CO2_offsets = []

        self.original_length = 0    # original length of flux data
        self.data_loss = 0          # total percent of data set pruned
        
        self.cuts = []

        self.RSQ = 0        # R^2 for final rate calculation
        self.RoC = 0        # rate of change (concentration/minute)
        self.NEE = 0


def on_press(event, i, fluxes, line_L, line_R, fig, ax, cid):
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
            # if both indices exist, set CO2 delta as the gap between the last entry before and first entry after the cut
            # this is for the sake of maintaining a linear relationship, it adds a fixed offset to each entry after the cut
            # if the cut is at the boundary of the data, the offset will be set to zero (i.e. the cut isn't in the middle of the data)
            if time_L_index and time_R_index:
                CO2_delta = fluxes[i].pruned_CO2[time_L_index] - fluxes[i].pruned_CO2[time_R_index]
            else:
                CO2_delta = 0

            # obtain time offset as the time elapsed by the cut
            # for the sake of maintaining a linear relationship, adds a fixed offset
            time_delta = time_R - time_L
            fluxes[i].cuts.append([time_L_index, time_R_index])

            times = []
            CO2 = []

            # if entry is not in the cut, add it to a new set
            for k in range(len(fluxes[i].pruned_times)):
                if fluxes[i].pruned_times[k] < time_L:
                    times.append(fluxes[i].pruned_times[k])
                    CO2.append(fluxes[i].pruned_CO2[k])
                if fluxes[i].pruned_times[k] > time_R:
                    times.append(fluxes[i].pruned_times[k] - time_delta)
                    CO2.append(fluxes[i].pruned_CO2[k] + CO2_delta)
            
            # update flux pruned sets with cut sets
            fluxes[i].pruned_times = times
            fluxes[i].pruned_CO2 = CO2

            # refresh the plot
            draw_plot(i, fluxes, fig, ax, cid)

    # if r key pressed, reset data
    if event.key == 'r':
        fluxes[i].pruned_times = fluxes[i].times
        fluxes[i].pruned_CO2 = fluxes[i].CO2
        fluxes[i].cuts = []
        draw_plot(i, fluxes, fig, ax, cid)
    
    # right arrow key moves to the next flux, or exits if currently on the last flux
    if event.key == 'right':
        if i == len(fluxes) - 1:
            plt.ioff()
            plt.close('all')
            return 0
        else:
            draw_plot(i + 1, fluxes, fig, ax, cid)
        
    # left arrow key moves to the previous flux, or does nothing if at the beginning
    if event.key == 'left':
        if i != 0:
            draw_plot(i - 1, fluxes, fig, ax, cid)


def input_data(folder, field_file):
    files = glob.glob(folder + "/*.TXT") + glob.glob(folder + "/*.txt")
    fluxes = []
    time_regex = r"(\d*):(\d*):(\d*)"

    f = open(field_file, "r")
    x = next(f).replace('\t', ',').replace(';',',').split(",")
    volumes = {}

    for line in f:
        x = line.replace('\t', ',').replace(';', ',').split(",")
        for i in range(len(x)):
            x[i] = x[i].strip(' \t\n\r')
        volumes[x[0]] = float(x[1])*float(x[2])


    for file in files:

        name = os.path.splitext(os.path.basename(file))[0]

        CO2 = []
        times = []
        PARs = []
        temps = []

        f = open(file, 'r')
        start = 1
        
    
        for line in f:
            line = line.split(',')

            time_match = re.search(time_regex, line[2])

            if line[0] != 'M3':
                continue

            if start == 1:
                start_time = float(time_match[1])*3600 + float(time_match[2])*60 + float(time_match[3])
                CO2.append(float(line[5].strip(' \t\n\r')))
                times.append(0)
                PARs.append(float(line[13].strip(' \t\n\r')))
                temps.append(float(line[15].strip(' \t\n\r')))
                start = 0

            else:
                current_time = float(time_match[1])*3600 + float(time_match[2])*60 + float(time_match[3])

                CO2.append(float(line[5].strip(' \t\n\r')))
                times.append(current_time - start_time)
                PARs.append(float(line[13].strip(' \t\n\r')))
                temps.append(float(line[15].strip(' \t\n\r')))

        fluxes.append(Flux(name, CO2, times, PARs, temps, volumes[name]))

    return fluxes


def draw_plot(i, fluxes, fig, ax, cid):

    fig.clear()
    ax = fig.add_subplot()

    print(fluxes[i].pruned_times)
    print(fluxes[i].pruned_CO2)
    m, b, R2 = linear_regression(fluxes[i].pruned_times, fluxes[i].pruned_CO2)
    at = AnchoredText(
        r"$R^{2}$ = " + str(round(R2, 5)), prop=dict(size=15), frameon=True, loc='upper center')
    at.patch.set_boxstyle("round,pad=0.,rounding_size=0.2")
    ax.add_artist(at)
    
    plt.plot(fluxes[i].pruned_times, fluxes[i].pruned_CO2, linewidth = 2.0)
    line_L = draggable_lines(ax, fluxes[i].pruned_times[0], [fluxes[i].pruned_times[0], fluxes[i].pruned_times[-1]], plt.gca().get_ylim())   # left draggable boundary line
    line_R = draggable_lines(ax, fluxes[i].pruned_times[-1], [fluxes[i].pruned_times[0], fluxes[i].pruned_times[-1]], plt.gca().get_ylim())     # right draggable boundary line

    # if currently on the last flux, change header information, otherwise set title to user controls
    if i == len(fluxes) - 1:
        ax.set(title = fluxes[i].name +  "\nLast flux! Press right arrow to finish, enter to cut data, r to reset cuts\nUse the mouse to drag peak bounds")
    else:      
        ax.set(title = fluxes[i].name + '\nUse arrow keys to navigate fluxes, enter to cut data, r to reset cuts\nUse the mouse to drag cut bounds')

    ax.set(xlabel = "Time (s)")                 # x axis label

    ax.set(ylabel = "CO2 concentration (ppm)")  # y axis label
    fig.canvas.mpl_disconnect(cid)
    cid = fig.canvas.mpl_connect('key_press_event', lambda event: on_press(event, i, fluxes, line_L, line_R, fig, ax, cid))   # connect key press event  
    ax.grid(True)

    fig.canvas.draw()


def prune(fluxes):
    i = 0
    fig, ax = plt.subplots()
    fig.set_size_inches(9,6)
    cid = ''
    plt.grid(True)
    plt.ion()
    draw_plot(i, fluxes, fig, ax, cid)
    plt.show(block=False)
    while plt.get_fignums():
        fig.canvas.draw_idle()
        fig.canvas.start_event_loop(0.05)


def flux_calculation(fluxes):
    for flux in fluxes:
        X = flux.pruned_times
        Y = flux.pruned_CO2

        m, b, R2 = linear_regression(X, Y)

        flux.RSQ = R2
        flux.RoC = m * 60

        adjusted_volume = (flux.volume*273.15)/(flux.temp + 273.15)*1000
        flux.NEE = ((m*44.01)/(22.414))*(adjusted_volume/0.58 ** 2)*(86400/1000000)

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
                flux.pruned_CO2.insert(i, '')
            cut_size += cut[1] + cut[0] + 1

        # calculate the offset at each index
        for k in range(len(flux.times)):
            if flux.pruned_times[k] != '':
                flux.time_offsets.append(flux.times[k] - flux.pruned_times[k])
                flux.CO2_offsets.append(flux.CO2[k] - flux.pruned_CO2[k])
            else:
                flux.time_offsets.append(0)
                flux.CO2_offsets.append(0)

def output_data(fluxes, date):

    out = tkinter.filedialog.asksaveasfilename(defaultextension='.xlsx')
    workbook = xlsxwriter.Workbook(out)

    worksheet = workbook.add_worksheet("Summary")
    worksheet.write_row(1, 0, ["Date:", date])
    worksheet.write_column(3, 0, ["Flux name", '', "Chamber volume (L)", "Air temp (K)", '',  "RSQ", "Rate of change (CO2 [ppm/min])", "m (CO2 [ppm/sec])", "NEE (g CO2 m^-2 d^-1)", "Data loss (%)", "PAR", "Surface moisture", "Surface temperature"])

    for i in range(len(fluxes)):
        print(fluxes[i].name)
        worksheet.write_column(3, i + 1, [fluxes[i].name , '', fluxes[i].volume, fluxes[i].temp, '', fluxes[i].RSQ, fluxes[i].RoC, fluxes[i].RoC/60, fluxes[i].NEE, fluxes[i].data_loss, fluxes[i].PAR])
        worksheet.set_column(i + 1, i + 1, len(fluxes[i].name ))
    worksheet.set_column(0, 0, len("Rate of change (CH4 [ppm/min])"))


    worksheets = {}
    for flux in fluxes:
        if flux.name not in worksheets:
            worksheet = workbook.add_worksheet(flux.name)
            worksheets[flux.name] = 1
        else:
            worksheets[flux.name] += 1
            worksheet = workbook.add_worksheet(flux.name + " (%s)" %(str(worksheets[flux.name])))
        worksheet.write_row(0, 0, ["Name", flux.name])
        worksheet.write_row(1, 0, ["RSQ", flux.RSQ, '', '', "Chamber volume (L)", flux.volume])      
        worksheet.write_row(2, 0, ["Rate of change (CO2 [ppm/min]", flux.RoC, '', '', "Air temp (K)", flux.temp])
        worksheet.write_row(3, 0, ["NEE (g CO2 m^-2 d^-1)", flux.NEE])
        worksheet.write_row(4, 0, ["Data loss (%)", flux.data_loss])

        worksheet.write(6, 0, "Original times (s)")
        worksheet.set_column(0, 0, len("Rate of change (CH4 [ppm/min])"))
        worksheet.write(6, 1, "Pruned times (s)")
        worksheet.set_column(1, 1, len("Pruned times (s)"))
        worksheet.write(6, 2, "Time offsets (s)")
        worksheet.set_column(2, 2, len("Time offsets (s)"))

        worksheet.write(6, 4, "Original CO2 concentrations (ppm)")
        worksheet.set_column(4, 4, len("Original CO2 concentrations (ppm)"))
        worksheet.write(6, 5, "Pruned CO2 concentrations (ppm)")
        worksheet.set_column(5, 5, len("Pruned CO2 concentrations (ppm)"))
        worksheet.write(6, 6, "CO2 concentration offsets (ppm)")
        worksheet.set_column(6, 6, len("CO2 concentration offsets (ppm)"))

        worksheet.write(6, 8, "PARs")
        worksheet.write(6, 9, "Temperatures (K)")
        worksheet.set_column(8, 8, len("Temperatures (K)"))
        worksheet.set_column(9, 9, len("Temperatures (K)"))

        for i in range(len(flux.times)):
            worksheet.write_row(i + 7, 0, [flux.times[i], flux.pruned_times[i], flux.time_offsets[i], '', flux.CO2[i], flux.pruned_CO2[i], flux.CO2_offsets[i], '', flux.PARs[i], flux.temps[i]])

        # generate chart showing cut values compared to kept values with offsets
        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({'values' : '=\'%s\'!E8:E%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A8:A%i'%(flux.name, len(flux.CO2) + 9), 'name': 'Cut values', 'line': {'color': 'red'}})
        chart.add_series({'values' : '=\'%s\'!F8:F%i'%(flux.name, len(flux.times) + 9), 'categories' : '=\'%s\'!A9:A%i'%(flux.name, len(flux.CO2) + 9), 'name': 'Kept values', 'line': {'color': 'green'}})
        chart.set_y_axis({'interval_unit': 10, 'interval_tick': 2, 'name': 'CO2 concentration (ppm)'})
        chart.set_x_axis({'name': 'Time (s)'})
        chart.set_title({'name' : 'Concentration vs. Time'})
        chart.set_size({'width': 800, 'height': 600})
        worksheet.insert_chart('L8', chart)


    workbook.close()

    return 0



#########################################################################################################################
######################################## script execution starts here! ##################################################
#########################################################################################################################


def IRGA():

    layout = [[sg.Text('IRGA EGM-5 Data Processing Tool', font='Any 36', background_color='#E833FF')],
        [sg.Text("", background_color='#E833FF')],
        [sg.Text('Field data file: (.csv, .txt)', size=(21, 1), background_color='#E833FF'), sg.Input(key='-FIELD-'), sg.FileBrowse()],
        [sg.Text('IRGA files folder:', size=(15, 1), background_color='#E833FF'), sg.Input(key='-FOLDER-'), sg.FolderBrowse()],
        [sg.Text("Date:", size=(15, 1), background_color='#E833FF'), sg.InputText(key='-DATE-')],
        [sg.Text("", background_color='#E833FF')],
        [sg.Submit(), sg.Cancel()]]


    # Create the window
    window = sg.Window("IRGA EGM5", layout, margins=(80, 50), background_color='#E833FF')
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
            field_file = values['-FIELD-']
            folder = values['-FOLDER-']
            date = values['-DATE-']
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error in inputted information")

        try:
            fluxes = input_data(folder, field_file)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error parsing data files: Ensure your field data and LICOR data are formatted correctly")
        
        try:
            prune(fluxes)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error pruning data: Verify times are correct in field data file")

        try:
            flux_calculation(fluxes)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error generating linear model: This shouldn't happen, contact me at btnewton@uwaterloo.ca")

        try:  
            offsets(fluxes)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error generating linear offsets: This shouldn't happen, contact me at btnewton@uwaterloo.ca")
        
        try:
            out = output_data(fluxes, date)
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error outputting data: Ensure the chosen location is valid")

    window.close()
    return 0


if __name__ == "__main__":

    os.system('cls' if os.name == 'nt' else 'clear')
    print("================================================================================")
    print("========================== LICOR data processing tool ==========================")
    print("================================================================================\n")
    
    # get field data and LICOR data files from user prompts
    root = tkinter.Tk()
    root.withdraw()
    print("Choose field data file:")
    field_data = tkinter.filedialog.askopenfilename(title = "Choose field data file")