import PySimpleGUI as sg

layout = [[sg.Text("Hello PySimpleGUI")],
          [sg.Button("OK")]
]

window = sg.Window("PySimpleGUI Demo", layout)

while True:
    event, values = window.read()
    if event == "OK" or event == sg.WINDOW_CLOSED:
        break

window.close()
