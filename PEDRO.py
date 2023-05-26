from FMA import FMA
from GC import GC
from IRGA_EGM5 import IRGA
from LICOR import LICOR
from LICOR_Samples import LICOR_Samples
from LGR import LGR
from LGR_Samples import LGR_Samples
import traceback
import PySimpleGUI as sg

def error_screen(e):
    layout = [[sg.Text("Error", font='Any 48', background_color="#7A1E74")],
    [sg.Text("The following error occured:", background_color="#7A1E74")],
    [sg.Text(background_color="#7A1E74")],
    [sg.Text(e, background_color="#7A1E74")],
    [sg.Text(background_color="#7A1E74")],
    [sg.Button("Go back")]]

    window = sg.Window("ERROR", layout, margins=(60, 20), background_color = '#7A1E74')
    while True:
        event, values = window.read()
        if event == "Go back" or event == sg.WIN_CLOSED:
            break

    window.close()

layout = [[sg.Text("P.E.D.R.O.", font='Any 48', background_color="#DF4A4A")],
    [sg.Text("\"Peatland Equipment Data Re-Organizer\"", font ='italic', background_color="#DF4A4A")],
    [sg.Text(background_color="#DF4A4A")],
    [sg.Text("Select an option below:", background_color="#DF4A4A")], 
    [sg.Button("FMA"), sg.Button("LICOR (flux)"), sg.Button("LICOR (samples)"), sg.Button("GC")],
    [sg.Button("LGR (flux)"), sg.Button("LGR (samples)"), sg.Button("IRGA EGM-5")],
    [sg.Text(background_color="#DF4A4A")],
    [sg.Text("v1.0.0", background_color="#DF4A4A")]]

# Create the window
window = sg.Window("PEDRO", layout, margins=(60, 20), background_color="#DF4A4A")

# Create an event loop
while True:
    event, values = window.read()
    # End program if user closes window or
    # presses the OK button
    try:
        if event == "QUIT" or event == sg.WIN_CLOSED:
            break
        elif event == "LICOR (flux)":
            window.Hide()
            print("===== LICOR (flux) =====")
            LICOR()
            window.UnHide()

        elif event == "LICOR (samples)":
            window.Hide()
            print("===== LICOR (samples) =====")
            LICOR_Samples()
            window.UnHide()

        elif event == "LGR (flux)":
            window.Hide()
            print("===== LGR (flux) =====")
            LGR()
            window.UnHide()

        elif event == "LGR (samples)":
            window.Hide()
            print("===== LGR (samples) =====")
            LGR_Samples()
            window.UnHide()

        elif event == "FMA":
            window.Hide()
            print("===== FMA =====")
            FMA()
            window.UnHide()

        elif event == "GC":
            window.Hide()
            print("===== GC =====")
            GC()
            window.UnHide()

        elif event == "IRGA EGM-5":
            window.Hide()
            print("===== IRGA =====")
            IRGA()
            window.UnHide()

    except Exception as e:
        print(traceback.format_exc())
        error_screen(e)

window.close()


