from math import nan
import sys
import matplotlib.pyplot as plt
import matplotlib.lines as lines
import numpy as np
import random
import csv
import datetime
import os
import re
import tkinter
import tkinter.filedialog
import xlsxwriter
from os.path import  join
import PySimpleGUI as sg
import traceback



# draggable lines for user cuts
class draggable_lines:
    def __init__(self, ax, start_coordinate, x_bounds, y_bounds):
        self.ax = ax
        self.c = ax.get_figure().canvas
        self.x_bounds = x_bounds
        self.y_bounds = y_bounds
        self.press = None

        self.line = lines.Line2D([start_coordinate, start_coordinate], y_bounds, color='r', picker=5)

        self.ax.add_line(self.line)
        self.c.draw_idle()
        self.sid = self.c.mpl_connect('button_press_event', self.on_press)
        self.sid = self.c.mpl_connect('motion_notify_event', self.on_motion)
        self.sid = self.c.mpl_connect('button_release_event', self.on_release)


    def on_press(self, event):
        if abs(event.xdata - self.line.get_xdata()[0]) < 3:
            self.press = (self.line.get_xdata()[0], event.xdata)
        return
    
    def on_motion(self, event):
        if self.press == None:
            return 
        try: 
            x0, xpress = self.press
            dx = event.xdata - xpress
            if self.x_bounds[0] >= (x0 + dx):
                self.line.set_xdata([self.x_bounds[0],self.x_bounds[0]])
            elif (x0 + dx) >= self.x_bounds[1]:
                self.line.set_xdata([self.x_bounds[1], self.x_bounds[1]])
            else:
                self.line.set_xdata([x0+dx, x0+dx])
        except:
            pass
    
    def on_release(self, event):
        self.press = None
        self.c.draw_idle()


# sample object, one created for each individual sample
class sample:
    def __init__(self, name, time):
        self.name = name                # sample name
        self.start_time = float(time)        # sample start time
        self.end_time = nan                  # sample end time
        self.peak_start_time = float(time)   # peak start time
        self.peak_end_time = nan             # peak end time
        self.times = []                 # FMA times between sample start and end times
        self.concentrations = []        # FMA concentrations between sample start and end times
        self.area = nan                 # sample peak area
        self.CH4 = nan                  # sample CH4 concentration


def input_data(sample_data, LICOR_data):
    samples = []
    LICOR = []

    time_regex = r"(\d*):(\d*):(\d*)"

    LICOR_time_regex = r"^TIME"
    LICOR_CH4_regex = r"CH4"
    LICOR_CO2_regex = r"CO2"

    f = open(LICOR_data, "r")
    x = next(f).replace('\t', ',').replace(';',',').split(",")
    while x[0] != "DATAH":
        x = next(f).replace('\t', ',').replace(';',',').split(",")
    print(x)
    for i in range(len(x)):
        if re.search(LICOR_time_regex, x[i], re.IGNORECASE):
            LICOR_time_index = i
        if re.search(LICOR_CH4_regex, x[i], re.IGNORECASE):
            LICOR_CH4_index = i
        if re.search(LICOR_CO2_regex, x[i], re.IGNORECASE):
            LICOR_CO2_index = i

    for line in f:
        x = line.replace('\t', ',').replace(';', ',').split(',')
        for k in range(len(x)):
            x[k] = x[k].strip(' \t\n\r')
        if x[0] == "DATA":
            time_match = re.search(time_regex, x[LICOR_time_index])
            time = float(time_match[1])*3600 + float(time_match[2])*60 + float(time_match[3])
            LICOR.append([time, float(x[LICOR_CH4_index])/1000, x[LICOR_CO2_index]])

    f = open(sample_data, "r") 
    for line in f:
        x = line.split(",")
        for i in range(len(x)):
            x[i] = x[i].strip(' \t\n\r')

        sample_time_match =re.search(time_regex, x[1])
        sample_time = float(sample_time_match[1])*3600 + float(sample_time_match[2])*60 + float(sample_time_match[3])

        samples.append(sample(x[0], sample_time))
    f.close()

    return samples, LICOR



# add end time to each sample
def process_samples(samples, LICOR):

    time_regex = r"(\d*):(\d*):(\d*)"

    for i in range(len(samples)):
        if i < len(samples) - 1:
            if samples[i + 1].peak_start_time - samples[i].peak_start_time > 180:
                samples[i].peak_end_time = samples[i].peak_start_time + 180
            else:
                samples[i].peak_end_time = samples[i + 1].peak_start_time
        else:
            samples[i].peak_end_time = samples[i].peak_start_time + 180


# process button press for plot
def on_press(event, i, samples, line_L, line_R, LICOR, fig, ax, cid):
    sys.stdout.flush()

    # when changing samples, obtain newly set start and end times from the positions of the user set bounds
    samples[i].peak_start_time = line_L.line.get_xdata()[0]
    samples[i].peak_end_time = line_R.line.get_xdata()[0]

    # right arrow key moves to the next sample, or exits if currently on the last sample
    if event.key == 'right' or event.key == 'space':
        if i == len(samples) - 1:
            plt.close()
            return 0
        else:
            draw_plot(i + 1, samples, LICOR, fig, ax, cid)
            #draw_plot(i + 1, samples, FMA)
        
    # left arrow key moves to the previous sample, or does nothing if at the beginning
    if event.key == 'left':
        if i != 0:
            draw_plot(i - 1, samples, LICOR, fig, ax, cid)


#draw plot for i-th sample
def draw_plot(i, samples, LICOR, fig, ax, cid):
    
    # if sample hasn't already been processed
    if len(samples[i].times) == 0:
        for j in range(len(LICOR)):
            if abs(float(samples[i].peak_start_time) - float(LICOR[j][0])) < 1:
                LICOR_start = j
            if abs(float(samples[i].peak_end_time) - float(LICOR[j][0])) < 1:
                LICOR_end = j
        for k in range(LICOR_start, LICOR_end):
            samples[i].times.append(float(LICOR[k][0]))
            samples[i].concentrations.append(float(LICOR[k][1]))

    fig.clear()
    ax = fig.add_subplot()
    plt.plot(samples[i].times, samples[i].concentrations, linewidth = 2.0)

    line_L = draggable_lines(ax, samples[i].peak_start_time, [samples[i].start_time, samples[i].start_time + len(samples[i].concentrations)], plt.gca().get_ylim())   # left draggable boundary line
    line_R = draggable_lines(ax, samples[i].peak_end_time, [samples[i].start_time, samples[i].start_time + len(samples[i].concentrations)], plt.gca().get_ylim())     # right draggable boundary line

    # if currently on the last sample, change header information, otherwise set title to user controls
    if i == len(samples) - 1:
        ax.set(title = samples[i].name + "\nLast sample! Press right arrow to finish\nUse the mouse to drag peak bounds")
    else:      
        ax.set(title = samples[i].name + '\nUse arrow keys to navigate samples\nUse the mouse to drag peak bounds')

    ax.set(xlabel = "Time (s)")                 # x axis label
    ax.set(ylabel = "CH4 concentration (ppm)")  # y axis label
    fig.canvas.mpl_disconnect(cid)
    cid = fig.canvas.mpl_connect('key_press_event', lambda event: on_press(event, i, samples, line_L, line_R, LICOR, fig, ax, cid))   # connect key press event  
    ax.grid(True)
    fig.canvas.draw()


# obtains peak bounds from user interactive plots
def obtain_peaks(samples, LICOR):
    i = 0
    fig, ax = plt.subplots()
    fig.set_size_inches(6,6)
    cid = fig.canvas.mpl_connect('key_press_event', lambda event: on_press(event, i, samples, line_L, line_R, LICOR, fig, ax, cid))   # connect key press event
    plt.grid(True)
    plt.ion()
    draw_plot(i, samples, LICOR, fig, ax, cid)
    plt.show(block=False)
    while plt.get_fignums():
        fig.canvas.draw_idle()
        fig.canvas.start_event_loop(0.05)


# using the new user set start and end times, trim all extraneous data outside these bounds
# also cut out samples from overall FMA data set, for random baseline sampling used to generate linear model
def standardize(samples, LICOR):
    for j in range(len(samples)):
        LICOR_start = 0           # sample start index within sample time and concentration sets (not FMA data indices)
        LICOR_end = 0             # sample end index within sample time and concentration sets (not FMA data incides)
        LICOR_cut_L = 0           # sample start index within FMA data set
        LICOR_cut_R = 0           # sample end index within FMA data set
        start_temp = np.inf     # temporary variable for finding closest FMA time to user set start time
        end_temp = np.inf       # temporary variable for finding closest FMA time to user set end time

        # trims sample time and concentration sets
        #TODO: This can be smarter, will come back later
        for i in range(len(samples[j].times)):
            if abs(samples[j].peak_start_time - samples[j].times[i]) < start_temp:
                start_temp = abs(samples[j].peak_start_time - samples[j].times[i])
                LICOR_start = i
            if abs(samples[j].peak_end_time - samples[j].times[i]) < end_temp:
                end_temp = abs(samples[j].peak_end_time - samples[j].times[i])
                LICOR_end = i
        #/TODO

        # trim sample times and concentrations 
        samples[j].times = samples[j].times[LICOR_start: LICOR_end + 1]
        samples[j].concentrations = samples[j].concentrations[LICOR_start: LICOR_end + 1]

        # cuts out sample from overall FMA data (used for baseline sampling in generating linear model)
        for k in range(len(LICOR)):
            if float(LICOR[k][0]) == samples[j].times[0]:
                LICOR_cut_L = k
            if float(LICOR[k][0]) == samples[j].times[-1]:
                LICOR_cut_R = k
        LICOR = LICOR[0: LICOR_cut_L] + LICOR[LICOR_cut_R + 1: -1]


# integrates each sample to get peak areas
def peak_areas(samples):
    for i in range(len(samples)):

        # determines whether the sample has a positive or negative peak based on the value of the middle relative to the beginning
        if samples[i].concentrations[int(len(samples[i].concentrations)/2)] < samples[i].concentrations[0]: # negative peak
            baseline = max(samples[i].concentrations)
        else: # positive peak
            baseline = min(samples[i].concentrations)
        area = 0

        # integrate concentration data points with respect to time
        for j in range(len(samples[i].times)):
            if j < len(samples[i].times) - 1:
                area = area + (samples[i].concentrations[j] - baseline)*(samples[i].times[j + 1] - samples[i].times[j])
        samples[i].area = area


# performs linear regression to generate linear relationship between peak areas and CH4 concentration
def linear_model(samples, LICOR):
    X = []
    Y = []

    
    one_ppm_regex = r"1\s?ppm"
    five_ppm_regex = r"5\s?ppm"
    fifty_ppm_regex = r"50\s?ppm"


    # extracts standards from samples
    for i in range(len(samples)):
        if re.search(one_ppm_regex, samples[i].name):
            Y.append(1)
            X.append(samples[i].area)
        if re.search(five_ppm_regex, samples[i].name):
            Y.append(5)
            X.append(samples[i].area)
        if re.search(fifty_ppm_regex, samples[i].name):
            Y.append(50.6)
            X.append(samples[i].area)

    # for each set of standards, generates a random non-sample baseline data point
    numStandards = len(X)

    if numStandards == 0:
        return(1, 1, 1)

    for i in range(numStandards):
        Y.append(float(LICOR[random.randint(0, len(LICOR))][1]))
        X.append(0)

    # simple linear regression
    sumX = sum(X)
    sumY = sum(Y)
    meanX = sumX/len(X)
    meanY = sumY/len(Y)

    SSx = 0 # sum of squares
    SP = 0  # sum of products

    for i in range(len(X)):
        SSx = SSx + (X[i] - meanX) ** 2
        SP = SP + (X[i] - meanX)*(Y[i] - meanY)

    # generates slope and intercept based on standards and baseline samples
    m = SP/SSx
    b = meanY - m * meanX

    SS_res = 0
    SS_t = 0

    for i in range(len(X)):
        SS_res = SS_res + (Y[i] - X[i]*m - b) ** 2
        SS_t = SS_t + (Y[i] - meanY) ** 2
    
    R2 = 1 - SS_res/SS_t    # coefficient of determination

    return(m, b, R2)


# For each sample uses the determined linear model to determine sample CH4 concentration
def concentrations(samples, m, b):
    for i in range(len(samples)):
        CH4 = samples[i].area * m + b
        samples[i].CH4 = CH4


# outputs data to csv file
def outputData(samples, m, b, R2):
    out = tkinter.filedialog.asksaveasfilename(defaultextension='.xlsx')
    workbook = xlsxwriter.Workbook(out)

    worksheet = workbook.add_worksheet("Results")
    worksheet.write_row(0, 0, ["Date: ", str(datetime.datetime.now().replace(microsecond=0))])
    worksheet.write_row(2, 0, ["Linear model:"])
    worksheet.write_row(3, 0, ["m:", m])
    worksheet.write_row(4, 0, ["b:", b])
    worksheet.write_row(5, 0, ["R2:", R2])
    worksheet.write_row(7, 0, ["Sample name", "Start time (s)", "End time (s)", "Peak area", "CH4 concentration (ppm)"])

    row = 8

    for i in range(len(samples)):
        worksheet.write_row(row, 0, [samples[i].name, samples[i].peak_start_time, samples[i].peak_end_time, samples[i].area, "=D{}*B4+B5".format(str(row + 1))])
        row += 1

    workbook.close()
    return out


#########################################################################################################################
######################################## script execution starts here! ##################################################
#########################################################################################################################

# parse input data -> obtain peaks from plots -> integrate peaks -> linear regression -> CH4 calculation -> output

def LICOR_Samples():

    layout = [[sg.Text('LICOR sample Data Processing Tool', font='Any 36', background_color='#01A100')],
        [sg.Text("", background_color='#01A100')],
        [sg.Text('Sample data file: (.csv, .txt)', size=(21, 1), background_color='#01A100'), sg.Input(key='-SAMPLES-'), sg.FileBrowse()],
        [sg.Text('LICOR data file: (.csv, .txt, .data)', size=(21, 1), background_color='#01A100'), sg.Input(key='-LICOR-'), sg.FileBrowse()],
        [sg.Text("", background_color='#01A100')],
        [sg.Submit(), sg.Cancel()]]

    #layout = [[sg.Column(layout, key='-COL1-'), sg.Column(layout2, visible=False, key='-COL2-'), sg.Column(layout3, visible=False, key='-COL3-')],
     #     [sg.Button('Cycle Layout'), sg.Button('1'), sg.Button('2'), sg.Button('3'), sg.Button('Exit')]]


    # Create the window
    window = sg.Window("LICOR samples", layout, margins=(80, 50), background_color='#01A100')
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
        sample_data = values['-SAMPLES-']
        LICOR_data = values['-LICOR-']

        try:
            print("Reading input files")
            samples, LICOR = input_data(sample_data, LICOR_data)
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error reading data files, ensure they're formatted correctly")

        try:
            print("Processing data")
            process_samples(samples, LICOR)
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error processing samples: This shouldn't happen, contact me at btnewton@uwaterloo.ca")

        try:
            print("Obtaining peak bounds")
            obtain_peaks(samples, LICOR)
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error obtaining peaks: Ensure peak bounds were chosen correctly")

        try:
            print("Trimming data")
            standardize(samples, LICOR)
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error standardizing results: This shouldn't happen, contact me at btnewton@uwaterloo.ca")

        try:
            print("Calculating peak areas")
            peak_areas(samples)
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error obtaining peak areas: This shouldn't happen, contact me at btnewton@uwaterloo.ca")

        try:
            print("Generating linear model")
            m, b, R2 = linear_model(samples, LICOR)
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error obtaining peak areas: This shouldn't happen, contact me at btnewton@uwaterloo.ca")

        try:
            print("Calculating concentrations")
            concentrations(samples, m, b)
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error obtaining peak areas: This shouldn't happen, contact me at btnewton@uwaterloo.ca")

        try:
            print("Outputting data")
            out = outputData(samples, m, b, R2) 
        except:
            window.close()
            print(traceback.format_exc())
            raise Exception("Error outputting results: Ensure chosen location is valid")

    window.close()
    return 0


if __name__ == "__main__":

    os.system('cls' if os.name == 'nt' else 'clear')
    print("================================================================================")
    print("======================= LICOR SAMPLE data processing tool ======================")
    #print("================================================================================")
    #print("======= For any questions or concerns please email btnewton@uwaterloo.ca =======")
    print("================================================================================\n")
    
    root = tkinter.Tk()
    root.withdraw()
    print("Choose sample timing file:")
    sample_data = tkinter.filedialog.askopenfilename(title = "Choose sample timing file")
    print("Choose LICOR data file:")
    FMA_data = tkinter.filedialog.askopenfilename(title = "Choose LICOR data file")


    # Parse sample and FMA data from passed text files
    try:
        print("Start data parsing")
        samples, FMA = input_data(sample_data, FMA_data)
    except:
        print("Error parsing data, please ensure your data files are named correctly")
        sys.exit(1)
    else:
        print("Data parsing complete")


    # Add end times to each sample, for the sake of plotting
    try:
        print("Start sample processing")
        process_samples(samples, FMA)
    except:
        print("Error processing data, please ensure data files are formatted correctly")
        sys.exit(1)
    else:
        print("Sample processing complete")


    # Plot each sample time range, have user set peak bounds
    try:
        print("Obtain peak bounds")
        obtain_peaks(samples, FMA)
    except:
        print("Error obtaining peak bounds")
        sys.exit(1)
    else:
        print("Peak bounds determination complete")


    # Trim extraneous data outside of user set bounds
    try:
        print("Trimming sample data")
        standardize(samples, FMA)
    except:
        print("Error trimming sample data")
        sys.exit(1)
    else:
        print("Data trimming complete")


    # Determine peak areas
    try:
        print("Determining peak areas")
        peak_areas(samples)
    except:
        print("Error determining peak areas")
        sys.exit(1)
    else:  
        print("Peak area determination complete")


    # Generate linear model of the form, CH4 ppm = m * area + b
    try:
        print("Generating linear model")
        m, b, R2 = linear_model(samples, FMA)
    except:
        print("Error generating linear model")
        sys.exit(1)
    else:
        print("Linear model generation complete")

    # Calculate CH4 concentration of each sample
    try:
        print("Calculating CH4 concentrations")
        concentrations(samples, m, b)
    except:
        print("Error calculating CH4 concentrations")
        sys.exit(1)
    else:
        print("CH4 calculation complete")


    out = outputData(samples, m, b, R2)
    # Output data in csv file
    #try:
    #    print("Outputting results")
    #    out = outputData(samples, m, b, R2, folder)
    #except:
    #    print("Error creating output file")
    #    sys.exit(1)
    #else:
    #    print("Outputting results complete")

    print("FMA data analysis complete! The results have been outputted to: " + out)