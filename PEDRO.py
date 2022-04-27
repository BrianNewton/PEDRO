from FMA import FMA
from LICOR import LICOR
from GC import GC
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
    [sg.Text("Select an option bellow:", background_color="#DF4A4A")], [sg.Button("FMA"), sg.Button("LICOR"), sg.Button("GC"), sg.Button("IRGA (coming soon!)")]]

# Create the window
window = sg.Window("PEDRO", layout, margins=(60, 20), background_color="#DF4A4A")

# Create an event loop
while True:
    event, values = window.read()
    # End program if user closes window or
    # presses the OK button
    if event == "QUIT" or event == sg.WIN_CLOSED:
        break
    elif event == "LICOR":
        window.Hide()
        try:
            print("===== LICOR =====")
            LICOR()
        except Exception as e:
            print(traceback.format_exc())
            error_screen(e)
        window.UnHide()
    elif event == "FMA":
        window.Hide()
        try:
            print("===== FMA =====")
            FMA()
        except Exception as e:
            print(traceback.format_exc())
            error_screen(e)
        window.UnHide()
    elif event == "GC (unstable)":
        window.Hide()
        try:
            print("===== GC =====")
            GC()
        except Exception as e:
            print(traceback.format_exc())
            error_screen(e)
        window.UnHide()

window.close()


