import PySimpleGUI as sg
import pandas as pd
import pywhatkit


def complete_windows():
    CompleteLayout = [[sg.Text('The Absent Messages are sent to the corresponding students.', font=("Inter", 25))]]
    windows = sg.Window("Attendance Marker", CompleteLayout, keep_on_top=True, size=(1200, 500))
    while True:
        event, values = windows.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break


sg.theme('LightGrey1')

list_Numbers = []

list_Num = sg.Listbox(values=list_Numbers, enable_events=True, key='-LISTBOX-', size=(40, 20))

LeftColumn = [
    [sg.Image('C:/Users/umesh/OneDrive/Desktop/Umesh/logo_Final (1).png', size=(75, 75)),
     sg.Text("National\nEngineering\nCollege", font=("Inter", 12))],
    [sg.T('Enter the Register Number :')],
    [sg.In(key='INPUT', do_not_clear=False)],
    [sg.Button('Add')],
    [sg.Text("", size=(0, 1), key='OUTPUT')],
    [sg.Text("View the Typed Register Number Details :")]

]

RightColumn = [
    [sg.T('The Selected Student Detail :')],
    [list_Num],
    [sg.Button('Remove'), sg.Button('Send')],
    [sg.Button('Exit')]
]

WindowsLayout = [
    [
        sg.Column(LeftColumn),
        sg.VerticalSeparator(),
        sg.Column(RightColumn)
    ]
]

windows = sg.Window("Attendance Marker", WindowsLayout, keep_on_top=True)

while True:
    event, values = windows.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    elif event == 'Add':
        name = values['INPUT']
        if name not in list_Numbers:
            list_Numbers.append(name)
        windows['-LISTBOX-'].update(list_Numbers)
        values['INPUT'] = ""

    elif event == 'Remove':
        if len(list_Numbers) > 0:
            val = list_Num.get()[0]
            list_Numbers.remove(val)
            windows['-LISTBOX-'].update(list_Numbers)

    elif event == 'Send':
        windows.close()
        Student = pd.read_csv("C:/Users/umesh/OneDrive/Desktop/Umesh/Student_Data.csv")
        Send_List = []
        Register = Student['Reg_No'].tolist()
        for i in range(len(list_Numbers)):
            for j in range(len(Register)):
                if list_Numbers[i] == str(Register[j]):
                    Send_List.append(j)
        for i in Send_List:
            message = "Dear Parent,\nYour Ward " + str(Student['Name'][i]) + ", " + str(
                Student['Reg_No'][i]) + " of " + str(
                Student['Tutor'][i]) + " tutor Ward is absent for the day. Kindly inform to " + str(
                Student['Tutor'][i]) + " as early as possible if not informed.\nThank You\nRegards\n" + str(
                Student['Department'][i]) + " Department\nNational Engineering College, Kovilpatti"
            pywhatkit.sendwhatmsg_instantly("+91" + str(Student['Phone'][i]), message, wait_time=15, tab_close=True,
                                            close_time=30)
        complete_windows()

# Excel sheet below the text selected
