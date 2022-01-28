import win32com.client
import glob
import time
from tkinter import filedialog
from tkinter import *

import os


def list_of_files(path, extension):
    """
    Returns list of paths to files with extension needed.

    Parameters
    ----------
    path : str
        Path to general folder (top level) to search in.
    extension : str
        File extension to search (for instance .xlsx or .dxf).

    Returns
    -------
    paths_list : list
        List of paths to files with extension needed

    """
    paths_list = glob.glob(path + f"/**/*.{extension}", recursive=True)

    return paths_list


def choose_folder():
    global folder_selected
    folder_selected = filedialog.askdirectory()
    try:
        text_first.delete(1.0, END)
        text_first.insert(1.0, folder_selected)
    except:
        text_first.insert(1.0, folder_selected)
    return folder_selected


def choose_file():
    global file_to_convert
    file_to_convert = filedialog.askopenfilename()
    try:
        text_second.delete(1.0, END)
        text_second.insert(1.0, file_to_convert)
    except:
        text_second.insert(1.0, file_to_convert)
    return file_to_convert


def connect_to_acad():
    """
    Connects to AutoCAD application.

    Returns
    -------
    wincad_obj : wincad_obj
        AutoCAD application object.

    """
    wincad_obj = win32com.client.Dispatch("AutoCAD.Application")

    return wincad_obj


def delaying_execution_and_return(function):
    """
    Delays execution of function for args[-1] seconds

    Parameters
    ----------
    function : func
        Function to be decorated.

    Returns
    -------
    wrapper : func
        Decorated function.

    """

    def wrapper(*args):
        try:
            res_inter = function(*args)
            return res_inter
        except:
            time.sleep(args[-1])
            res_inter = function(*args)
            return res_inter

    return wrapper


def delaying_execution(function):
    """
    Delays execution of function for args[-1] seconds

    Parameters
    ----------
    function : func
        Function to be decorated.

    Returns
    -------
    wrapper : func
        Decorated function.

    """

    def wrapper(*args):
        try:
            function(*args)
        except:
            time.sleep(args[-1])
            function(*args)

    return wrapper


@delaying_execution_and_return
def open_dwg(path_to_file, wincad_obj, delay):
    """
    Opens .dwg (AutoCAD drawing) file.

    Parameters
    ----------
    path_to_file : str
        Path to file.
    wincad_obj : wincad_obj
        AutoCAD application object.

    Returns
    -------
    acad_doc : acad_doc
        AutoCAD document object.

    """

    acad_doc = wincad_obj.Documents.Open(path_to_file)

    return acad_doc


@delaying_execution
def send_com_to_acad(autocad_document, command_for_acad, delay):
    """
    Sends direct command to AutoCAD command line.
    Parameters
    ----------
    autocad_document : acad_doc_object
        AutoCAD drawing.
    command_for_acad : str
        Text of command to be entered to AutoCAD command line.
    delay : int
        Delay of execution, sec in case of exception.
    """

    autocad_document.SendCommand(command_for_acad)


@delaying_execution
def acad_save_as(autocad_document, path, ext_code, delay):
    """
    Saves drawing.
    Parameters
    ----------
    autocad_document : acad_doc_object
        AutoCAD drawing.
    path : str
        Patch to directory where file should be saved (and filename).
    ext_code : int
        Extension code type.
    delay : int
        Delay of execution, sec in case of exception.
    """

    autocad_document.saveas(path, ext_code)


@delaying_execution
def acad_doc_close(autocad_document, delay):
    """
    Closing AutoCAD document.
    Parameters
    ----------
    autocad_document : acad_doc_object
        AutoCAD drawing.
    delay : int
        Delay of execution, sec in case of exception.
    """

    autocad_document.close()


def save_and_close(path_file, acad_doc, delay_const):
    '''Saving and closing AutoCAD file.
       Function takes "path_file" - path to file,
                      "acad_doc" - AutoCAD object,
                      "delay_const"  - delay parameter,  type: integer, given in seconds.'''

    if path_file[-4:] == '.dxf':
        new_name = path_file[:-4]+'_'
        acad_save_as(acad_doc,new_name , 61, delay_const)
        acad_doc_close(acad_doc, delay_const)
        os.remove(path_file, dir_fd=None)
        os.rename(new_name+'.dxf',path_file)
        
        try:
            os.remove(path_file + '.dwg', dir_fd=None)
        except:
            pass
    elif path_file[-4:] == '.dwg':
        acad_save_as(acad_doc, path_file, 60, delay_const)
        acad_doc_close(acad_doc, delay_const)
    else:
        print(f'Unrecognized file extension: {path_file[-4:]}')

@delaying_execution_and_return
def get_insunits(doc, delay_const):
    result = doc.GetVariable('INSUNITS')
    return result

def convertation(target_file,wincad_obj):
    acad_doc = open_dwg(target_file, wincad_obj, delay_par)
    insunits = get_insunits(acad_doc, delay_par)
    if insunits == 1 or insunits == 0:
        send_com_to_acad(acad_doc, "-DWGUNITS\n1\n4\n4\n\n\n", delay_par)
        send_com_to_acad(acad_doc, "DIMUNIT\n4\n", delay_par)
        send_com_to_acad(acad_doc, "LTSCALE\n'CAL\n(10/254)\n", delay_par)
        send_com_to_acad(acad_doc, "PSVPSCALE\n'CAL\n(10/254)\n", delay_par)
        send_com_to_acad(acad_doc, "AI_SELALL\nSCALE\n0,0,0\n'CAL\n(10/254)\n", delay_par)
        send_com_to_acad(acad_doc, "Z\nE\n", delay_par)
        save_and_close(target_file, acad_doc, delay_par)
    else:
        send_com_to_acad(acad_doc, "-DWGUNITS\n1\n5\n4\n\n\n\n\n", delay_par)
        send_com_to_acad(acad_doc, "DIMUNIT\n4\n", delay_par)
        send_com_to_acad(acad_doc, "Z\nE\n", delay_par)
        save_and_close(target_file, acad_doc, delay_par)

def resave_to_imperial():
    wincad = connect_to_acad()
    message = 'good'
    if file_to_convert != '':
        convertation(file_to_convert,wincad)
        bak_files_list = list_of_files(os.path.dirname(folder_selected), 'bak')


    elif folder_selected != '':
        if var_dwg.get() == 1:
            dwg_files_list = list_of_files(folder_selected, 'dwg')
            for file in dwg_files_list:
                convertation(file,wincad)

        if var_dxf.get() == 1:
            dxf_files_list = list_of_files(folder_selected, 'dxf')
            for file in dxf_files_list:
                convertation(file,wincad)

        bak_files_list = list_of_files(folder_selected, 'bak')

    else:
        message = 'bad'

    try:
        [os.remove(file) for file in bak_files_list]
    except:
        pass

    if message == 'good':
        canvas.create_text(20, 125, text='DONE!')
    elif message == 'bad':
        canvas.create_text(400, 125, text='CHECK YOUR INPUT!')

delay_par = 7
file_to_convert = ''
folder_selected = ''

tk = Tk()
tk.title('Resave to imperial units')  # title of the window
tk.resizable(0, 0)
tk.wm_attributes('-topmost', 1)

canvas = Canvas(tk, width=500, height=140)
canvas.pack()

text_first = Text(canvas, width=53, height=0.5, font=('Arial', 10))
text_first.place(x=10, y=57.5)

text_second = Text(canvas, width=53, height=0.5, font=('Arial', 10))
text_second.place(x=10, y=12.5)

btn_folder = Button(tk, text='Choose the folder!', command=choose_folder)
btn_folder.place(x=392.5, y=55)

btn_file = Button(tk, text='Choose the file!', command=choose_file)
btn_file.place(x=392.5, y=10)

btn_sort = Button(tk, text='Resave to imperial units!', command=resave_to_imperial)
btn_sort.place(x=180, y=110)

l1 = Label(tk, text="OR")
l1.place(x=210, y=35)

var_dwg = IntVar()
var_dxf = IntVar()
dwg_checkbox = Checkbutton(tk, text='DWG', variable=var_dwg)
dwg_checkbox.place(x=10, y=80)

dxf_checkbox = Checkbutton(tk, text='DXF', variable=var_dxf)
dxf_checkbox.place(x=75, y=80)

tk.mainloop()
