from FMA import FMA
from LICOR import LICOR
#import GC
import PySimpleGUI as sg

layout = [[sg.Text("P.E.D.R.O.", font='Any 48', background_color="#DF4A4A")],
    [sg.Text("\"Peatland Equipment Data Re-Organizer\"", background_color="#DF4A4A")],
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
        LICOR()
        window.UnHide()
    elif event == "FMA":
        window.Hide()
        FMA()
        window.UnHide()
    elif event == "GC":
        GC()

window.close()
