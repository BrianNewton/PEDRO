import traceback
import glob
import matplotlib.pyplot as plt
import matplotlib.lines as lines
from math import floor
import sys
import tkinter
import tkinter.filedialog
from matplotlib.offsetbox import AnchoredText
import PySimpleGUI as sg
import os
import re

from utils import draggable_lines

class Flux:
    def __init__(self, CO2, times, PARs, temps):
        self.CO2 = CO2
        self.times = times
        self.temp = sum(temps)/len(temps)
        self.PARs = PARs

        self.pruned_CO2 = CO2
        self.pruned_times = times

        self.time_offsets = []
        self.CO2_offsets = []

        self.original_length = 0    # original length of flux data
        self.data_loss = 0          # total percent of data set pruned
        
        self.cuts = []

        self.RSQ = 0        # R^2 for final rate calculation
        self.RoC = 0        # rate of change (concentration/minute)


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


def input_data(folder):
    files = glob.glob(folder + "/*.txt")
    fluxes = []
    time_regex = r"(\d*):(\d*):(\d*)"

    for file in files:
        CO2 = []
        times = []
        PARs = []
        temps = []

        f = open(file, 'r')

        line = f.readline().strip('\n').split(',')
        time_match = re.search(time_regex, line[2])
        start_time = float(time_match[1])*3600 + float(time_match[2])*60 + float(time_match[3])

        CO2.append(float(line[5].strip(' \t\n\r')))
        times.append(0)
        PARs.append(float(line[13].strip(' \t\n\r')))
        temps.append(float(line[15].strip(' \t\n\r')))

        for line in f:
            line = line.split(',')

            time_match = re.search(time_regex, line[2])
            current_time = float(time_match[1])*3600 + float(time_match[2])*60 + float(time_match[3])

            CO2.append(float(line[5].strip(' \t\n\r')))
            times.append(current_time - start_time)
            PARs.append(float(line[13].strip(' \t\n\r')))
            temps.append(float(line[15].strip(' \t\n\r')))

        fluxes.append(Flux(CO2, times, PARs, temps))

    return fluxes


def draw_plot(i, fluxes, fig, ax, cid):

    fig.clear()
    ax = fig.add_subplot()

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


def linear_regression():
    return 0

def flux_calculation():
    return 0

def offsets():
    return 0

def output_data():
    return 0



#########################################################################################################################
######################################## script execution starts here! ##################################################
#########################################################################################################################


def IRGA():

    layout = [[sg.Text('IRGA EGM-5 Data Processing Tool', font='Any 36', background_color='#E833FF')],
        [sg.Text("", background_color='#E833FF')],
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
            folder = values['-FOLDER-']
            date = values['-DATE-']
        except Exception as e:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error in inputted information")

        try:
            fluxes = input_data(folder)
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
            out = outputData(fluxes, site, date, CO2_or_CO2)
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