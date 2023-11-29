# -*- coding: utf-8 -*-
"""
Created on Thu Feb 24 11:53:16 2022

@author: BAC5MC
"""

#%%
# rename from folder list

# from IPython import get_ipython
# get_ipython().magic('reset -sf')

import PySimpleGUI as sg

import os.path, sys
from tkinter import Tcl
import xlrd #xlrd regebbi valtozattal:  pip install xlrd==1.2.0
from PIL import Image, ImageTk
import io
import fitz  # PyMuPDF, imported as fitz for backward compatibility reasons

# file_path = "my_file.pdf"
# doc = fitz.open(file_path)  # open document
# for page in doc:
#     pix = page.get_pixmap()  # render page to an image
#     pix.save(f"page_{i}.png")
    
# from pdf2image import convert_from_path

def get_img_data(f, maxsize=(400, 250), first=False):
    """Generate image data using PIL
    """
    
    print("WHAT: ",f[-3:])
    if f[-3:]=='pdf':
        print("PDF??")
        doc=fitz.open(f)
        img= doc[0].get_pixmap()
        print("prob 0")
        img=img.pil_tobytes(format="PNG", optimize=True)
        img=Image.open(io.BytesIO(img))
        # img = convert_from_path(f,size=(400,250))
    else:
        img = Image.open(f)
        
    img.thumbnail(maxsize)
    if first:                     # tkinter is inactive the first time
        print("prob?")
        bio = io.BytesIO()
        img.save(bio, format="PNG")
        del img
        return bio.getvalue()
    return ImageTk.PhotoImage(img)

def zoom_at(img, first=False):
    print("WHAAAT 2")
    if img[-3:]=='pdf':
        print("PDF??")
        doc=fitz.open(img)
        img= doc[0].get_pixmap()
        img=img.pil_tobytes()
        # img = convert_from_path(f,size=(400,250))
    else:
        img = Image.open(img)
    # img = Image.open(img)
    img = img.crop((50, 0, 
                    200, 20))
    x=int(150*2.7)
    y=int(20*2.7)
    img = img.resize((x, y), Image.LANCZOS)
    if first:                     # tkinter is inactive the first time
        bio = io.BytesIO()
        img.save(bio, format="PNG")
        del img
        return bio.getvalue()
    return ImageTk.PhotoImage(img)

# im=zoom_at(img)
# im.show()

# First the window layout in 2 columns
TT=sg.theme('Green Mono')
T_number=str(000)
S_number=str(000)
# d=sg.ListOfLookAndFeelValues()
# f=sg.SetOptions(button_color)
file_list_column = [
    [
        sg.Text("Folder with source names"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse(),
    ],
    [        
     sg.Text("Excel with source names "),
     sg.In(size=(25, 1), enable_events=True, key="-EXCEL-"),
     sg.FileBrowse("   File  ")
     ],
    [
     sg.Checkbox("Search for ''Identifier'' column in excel", key='Identifier',default=True)
     ],
    [
        sg.Listbox(
        values=[], enable_events=True, size=(40, 20), key="-FILE LIST-")
        ,sg.Text(S_number, key="S_number")
    ],

]
# filennn=r"\\10.11.114.114\tc-mc_reports$\2022\TS22-01029\TS22-01029.04_Noise_and_vibration_measurement\Results\Vibration_Level\Vib_Level_SN-037B12_sweep.png"
file_list_column2 = [
    [
        sg.Text("Folder to rename"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER2-"),
        sg.FolderBrowse(),
    ],
    [sg.Text("                                                       "),
    sg.Text("Vera <3",font = ("Arial, 8"), text_color='green', background_color='#A8C1B4',visible=False, key="-love-")],

    [
        sg.Listbox(values=[], enable_events=True, size=(40, 20), key="-FILE LIST2-"),
        sg.Text(T_number, key="T_number")
    ],
    ]
        # sg.Image(data=get_img_data(filennn,  first=True))
images=[
        
        sg.Image(key="-IMAGE_zoom-")],[
        sg.Image(key="-IMAGE-")
    
]

# ----- Full layout -----

layout = [

    [
        [
        #sg.Text("Title",size=(15, 1), font=("Verdana",11),text_color='Black',background_color='Yellow', justification='right
        sg.Text("Title",size=(7, 1),justification='center', text_color='white', background_color='#6D9F85'),
        sg.Input(size=(20, 1),key="-TITLE-B"),
        sg.Checkbox('Before name', key="-BEFORE-", default=False, size=(10,1)),
        sg.Input(size=(20, 1),key="-TITLE-A"),
        sg.Checkbox('After name', key="-AFTER-", default=True),
        sg.Button("Rename", button_color=('green'),enable_events=True,key="-RENAME-"),
        sg.Button("Undo", button_color=('green'), enable_events=True, key="-UNDO-", disabled=True),
        sg.Text("                                                                                                                               "),
        sg.Button(" ",border_width=0, size=(1, 1), button_color=('#A8C1B4'), enable_events=True, key="-EasterEgg-"),
        ],
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(file_list_column2),
        sg.Column(images)
    ]
]


window = sg.Window("File re-namer", layout)


# Run the Event Loop

while True:

    event, values = window.read()

    if event == "Exit" or event == sg.WIN_CLOSED:

        break

    # Folder name was filled in, make a list of files in the folder

    if event == "-FOLDER-":

        folder = values["-FOLDER-"]

        try:

            # Get list of files in folder
            file_list = Tcl().call('lsort', '-dict', os.listdir(folder))
            ff = [
                f
                for f in file_list
                if os.path.isfile(os.path.join(folder, f))]
            file_list = Tcl().call('lsort', '-dict', ff)
            #file_list = Tcl().call('lsort', '-dict', os.listdir(folder))
            try:
                file_list=list(file_list)
                file_list.remove("Thumbs.db")
                file_list=tuple(file_list)
            except:
                print("No Thumbs.db in folders")
            S_number=str(len(file_list))
            key_folder=True
            key_file=False

        except Exception as aa:
            print(aa)
            file_list = []

        # fnames =Tcl().call('lsort', '-dict', os.listdir(folder))
        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
        ]

        window["-FILE LIST-"].update(fnames)
        window["S_number"].update(S_number)
        window["-EXCEL-"].update([])
        window['-UNDO-'].update(disabled=False)
        # %
    
    if event == "-EXCEL-":

        excel_file = values["-EXCEL-"]

        try:

            # Get list of files in folder
            #file_list_excel = pd.read_excel(excel_file, header=None)
            file_list_excel=[]
            wb = xlrd.open_workbook(excel_file)
            sheet = wb.sheet_by_index(0)
            print(sheet)
            Z=0
            SS=False
            for colx in range(sheet.ncols):
                non_emptycells=[i for i,x in enumerate(sheet.col(colx)) if x.ctype != 0]
                print(len(non_emptycells))
            if values['Identifier'] == True:
                for i in range(sheet.ncols):
                    if sheet.cell_value(0,i)=="Identifier":
                        print("COL", i)
                        SS=True
                if SS==True:
                    for i in range(sheet.ncols):
                        if sheet.cell_value(0,i)=="Identifier":
                            print("COL2", i)
                            Z=i
                            
                    non_emptycells=[i for i,x in enumerate(sheet.col(Z)) if x.ctype != 0]
                    ROW=len(non_emptycells)
                    
                    for k in range(1,ROW):
                        print("ROW", k)
                        file_list_excel.append(sheet.cell_value(k,Z))
                        print(file_list_excel)
                elif SS==False:
                     file_list_excel='No_Identifier'
            elif values['Identifier'] == False:
                
                non_emptycells=[i for i,x in enumerate(sheet.col(0)) if x.ctype != 0]
                ROW=len(non_emptycells)  
                
                for i in range(0,ROW):                    
                    file_list_excel.append(sheet.cell_value(i,0))
                    
            # print(file_list_excel)
            
            S_number=str(len(file_list_excel))
            key_folder=False
            key_file=True
            
        except Exception as e_excel:
            print("excel hiba", e_excel)
            file_list_excel = []


        fnames_excel = [
            f
            for f in file_list_excel  ]
        # new_name = "".join(i for i in fnames_excel if i not in "\/:*?<>|")
        empty_excel_error=False
        empty_excel_error2=False
        for ii in range(len(fnames_excel)):
            if fnames_excel[ii]=='':
                empty_excel_error=True
        if len(file_list_excel)==0:
            empty_excel_error2=True
        if empty_excel_error==True:
            print('Exit1')
            sg.popup_error('Empty name in excel')  # Shows red error button
        if empty_excel_error2==True:
            print('Exit2')
            sg.popup_error('Empty excel')  # Shows red error button
            
        print(fnames_excel)
        window["-FILE LIST-"].update(fnames_excel)
        window["S_number"].update(S_number)
        window["-FOLDER-"].update([])
        window['-UNDO-'].update(disabled=True)
        # %
        
    if event == "-FOLDER2-":

        folder2 = values["-FOLDER2-"]

        try:

            # Get list of files in folder

            # file_list2 = os.listdir(folder2)
            file_list2 = Tcl().call('lsort', '-dict', os.listdir(folder2))
            print(file_list2)
            ff2 = [
                f
                for f in file_list2
                
                if os.path.isfile(os.path.join(folder2, f))]
                
            file_list2 = Tcl().call('lsort', '-dict', ff2)
            
            try:
                file_list2=list(file_list2)
                file_list2.remove("Thumbs.db")
                print(file_list2)
                # print("tt")
                file_list2=tuple(file_list2)
                # print("ttt")
            except:
                print("No Thumbs.db in folders")
            T_number=str(len(file_list2))
            # print("T_num",T_number)

        except:
            file_list2 = []
            print(file_list2)
        # fnames2 = Tcl().call('lsort', '-dict', os.listdir(folder2))
        fnames2 = [
            f
            for f in file_list2
            if os.path.isfile(os.path.join(folder2, f))
            #and f.lower().endswith((".png", ".gif"))
        ]

        window["-FILE LIST2-"].update(fnames2)
        window["T_number"].update(T_number)
        try:
            window["-IMAGE-"].update(data=get_img_data((folder2 + '\\' + fnames2[0]),  first=True))
            window["-IMAGE_zoom-"].update(zoom_at((folder2 + '\\' + fnames2[0]),  first=True))
        except Exception as e_i:
            print(e_i)
        
        
    if event == "-FILE LIST2-":
        try:
            file_list2 = Tcl().call('lsort', '-dict', os.listdir(folder2))
            ff2 = [
                f
                for f in file_list2
                if os.path.isfile(os.path.join(folder2, f))]
            file_list2 = Tcl().call('lsort', '-dict', ff2)
            try:
                file_list2=list(file_list2)
                file_list2.remove("Thumbs.db")
                file_list2=tuple(file_list2)
            except:
                print("No Thumbs.db in folders")
            fnames_ = [
    
                ff
    
                for ff in file_list2
    
                if os.path.isfile(os.path.join(folder2, ff))]
            try:
                ff = values["-FILE LIST2-"][0]            # selected filename
                print("ff:",ff)
                filename = os.path.join(folder2, ff)  # read this file
                print("filename:",filename)
                i = fnames_.index(ff)                 # update running index
                print("i: ",i)
            # else:
            #     filename = os.path.join(folder2, fnames_[i])
    
            # update window with new image
                try:
                    window["-IMAGE-"].update(data=get_img_data(folder2 + '\\' + fnames_[i],  first=True))
                    window["-IMAGE_zoom-"].update(zoom_at(filename,  first=True))
                except Exception as e_ii:
                    print(e_ii)
                
            except Exception as eeee:
                print("img not ok")
                print(eeee)
        except Exception as eeeee:
            print("no folder jet")
            print(eeeee)
            
    if event == "-RENAME-":
        try:
            print("Rename button pushed")
            # title=values["-TITLE-"]

            title_A=values["-TITLE-A"]
            title_B=values["-TITLE-B"]
            if key_folder is True:
                window['-UNDO-'].update(disabled=False)
                # rename based on file name list in folder
                folder_to=folder2+"\\"
                print('FOLDERS??')
                print('identifier False')
                folder_with_names=folder+"\\"

                source_name=[]
                source_name=os.listdir(folder_with_names)
                ss = [
                    s
                    for s in source_name
                    if os.path.isfile(os.path.join(folder_with_names, s))]
                source_name = Tcl().call('lsort', '-dict', ss)

                print(source_name)
                
                counted_names=[]
                counted_names=os.listdir(folder_to)
                cc = [
                    c
                    for c in counted_names
                    if os.path.isfile(os.path.join(folder_to, c))]
                counted_names = Tcl().call('lsort', '-dict', cc)
                
                #remove Thumbs.db
                try:
                    source_name=list(source_name)
                    source_name.remove("Thumbs.db")
                    # source_name=counted_names
                    source_name=tuple(source_name)
                except:
                    print("No Thumbs.db in source name")   
                try:   
                    counted_names=list(counted_names)
                    counted_names.remove("Thumbs.db")
                    counted_names=tuple(counted_names)
                except:
                    print("No Thumbs.db in counted name")

                # remove the file type from source
                source_name= [x[:-4] for x in source_name]
                
                # ask for file type
                kiterjesztes=counted_names[0][-4:]
                print('Lists are ready')
                
                #check the number of files
                if (len(counted_names) == len(source_name)):
                    print ( folder_to + " and " + folder_with_names + " have the same number of files")
                    count = 0
                    
                    print("whynot ok?")
                    T1=Tcl().call('lsort', '-dict', os.listdir(folder_to))
                    print(T1)
                    try:
                        T1=list(T1)
                        T1.remove("Thumbs.db")
                        T1=tuple(T1)
                    except:
                        print("Thumbs.db is not in folder1")
                    for file in T1:
                    # iterate all files from a directory
                    # for file in Tcl().call('lsort', '-dict', os.listdir(folder_to)):
                        # construct current name using file name and path
                        old_name = os.path.join(folder_to, file)
                        # Construct new file name
                        status_title_A=values["-AFTER-"]
                        status_title_B=values["-BEFORE-"]
                        title_A = "".join(i for i in title_A if i not in "\/:*?<>|")
                        title_B = "".join(i for i in title_B if i not in "\/:*?<>|")
                        source_name[count] = "".join(i for i in source_name[count] if i not in "\/:*?<>|")
                        if status_title_A==False and status_title_B==False:
                            new_name = folder_to  + source_name[count] + kiterjesztes
                        elif status_title_A==True and status_title_B==False:
                            new_name = folder_to  + source_name[count] + "_" + title_A + kiterjesztes
                        elif status_title_A==False and status_title_B==True:
                            new_name = folder_to + title_B + "_" + source_name[count] + kiterjesztes
                        elif status_title_A==True and status_title_B==True:
                            new_name = folder_to + title_B + "_" + source_name[count]+ "_" + title_A + kiterjesztes

                        #     if title=='':
                        #         new_name = folder_to  + source_name[count] + kiterjesztes
                        #     else:
                        #         new_name = folder_to  + source_name[count] + "_" + title + kiterjesztes
                        # else:
                        #     if title=='':
                        #         new_name = folder_to + source_name[count] + kiterjesztes
                        #     else:
                        #         new_name = folder_to + title + "_" + source_name[count] + kiterjesztes
                    
                        # Renaming the file
                        os.rename(old_name, new_name)
                        count += 1
                    
                    print('All Files Renamed')
                    print('New Names are')
                    # verify the result
                    res = os.listdir(folder_to)
                    res= Tcl().call('lsort', '-dict', res)
                    print(res)
                    
                    try:
                        # Get list of files in folder
                        file_list3 = Tcl().call('lsort', '-dict', os.listdir(folder2))
                        ff3 = [
                            f
                            for f in file_list3
                            if os.path.isfile(os.path.join(folder2, f))]
                        file_list3 = Tcl().call('lsort', '-dict', ff3)
                        try:
                            file_list3=list(file_list3)
                            file_list3.remove("Thumbs.db")
                            file_list3=tuple(file_list3)
                        except:
                            print("No Thumbs.db in folders")
                        T_number=str(len(file_list3))
        
                    except:
                        file_list3 = []
                    # fnames3= Tcl().call('lsort', '-dict', os.listdir(folder2))
                    fnames3 = [
                        f
                        for f in file_list3
                        if os.path.isfile(os.path.join(folder2, f))
                    ]
                    window["-FILE LIST2-"].update(fnames3)
                    window["T_number"].update(T_number)

                else:
                    print('Exit3')
                    sg.popup_error('Not the same file number')  # Shows red error button
                    #sys.exit()
            else:
                window['-UNDO-'].update(disabled=True)
                folder_to=folder2+"\\"
                file_with_names=excel_file
                source_name=[]
                # file_list_excel = pd.read_excel(excel_file, header=None)
                file_list_excel=[]
                wb = xlrd.open_workbook(excel_file)
                sheet = wb.sheet_by_index(0)
                if values['Identifier'] == True:
                    for i in range(sheet.ncols):
                        if sheet.cell_value(0,i)=="Identifier":
                            Z=i
                    for i in range(1,sheet.nrows):
                        file_list_excel.append(sheet.cell_value(i,Z))
                elif values['Identifier'] == False:
                    for i in range(0,sheet.nrows):
                        file_list_excel.append(sheet.cell_value(i,0))
                print("111",file_list_excel)
                
                # for i in range(sheet.nrows):
                #     file_list_excel.append(sheet.cell_value(i,0))
                source_name = [
                    f
                    for f in file_list_excel]
                
                counted_names=[]
                counted_names=os.listdir(folder_to)
                cc = [
                    c
                    for c in counted_names
                    if os.path.isfile(os.path.join(folder_to, c))]
                counted_names = Tcl().call('lsort', '-dict', cc)
                
                #remove Thumbs.db
                try:
                    source_name=list(source_name)
                    source_name.remove("Thumbs.db")
                    # source_name=counted_names
                    source_name=tuple(source_name)
                except:
                    print("No Thumbs.db in source name")   
                try:   
                    counted_names=list(counted_names)
                    counted_names.remove("Thumbs.db")
                    counted_names=tuple(counted_names)
                except:
                    print("No Thumbs.db in counted name")
                
                
                # remove the file type from source
                #source_name= [x[:-4] for x in source_name]
                
                # ask for file type
                kiterjesztes=counted_names[0][-4:]
                print('Lists are ready')
                
                #check the number of files
                if (len(counted_names) == len(source_name)):
                    print ( folder_to + " and " + file_with_names + " have the same number of files")
                    count = 0
                    # iterate all files from a directory
                    
                    T1=Tcl().call('lsort', '-dict', os.listdir(folder_to))
                    print(T1)
                    try:
                        T1=list(T1)
                        T1.remove("Thumbs.db")
                        T1=tuple(T1)
                    except:
                        print("Thumbs.db is not in folder1")
                    for file in T1:
                    # for file in Tcl().call('lsort', '-dict', os.listdir(folder_to)):
                        # construct current name using file name and path
                        old_name = os.path.join(folder_to, file)
                        # Construct new file name
                        status_title_A=values["-AFTER-"]
                        status_title_B=values["-BEFORE-"]
                        title_A = "".join(i for i in title_A if i not in "\/:*?<>|")
                        title_B = "".join(i for i in title_B if i not in "\/:*?<>|")
                        source_name[count] = "".join(i for i in str(source_name[count]) if i not in "\/:*?<>|")
                        if status_title_A==False and status_title_B==False:
                            new_name = folder_to  + source_name[count] + kiterjesztes
                        elif status_title_A==True and status_title_B==False:
                            new_name = folder_to  + source_name[count] + "_" + title_A + kiterjesztes
                        elif status_title_A==False and status_title_B==True:
                            new_name = folder_to + title_B + "_" + source_name[count] + kiterjesztes
                        elif status_title_A==True and status_title_B==True:
                            new_name = folder_to + title_B + "_" + source_name[count]+ "_" + title_A + kiterjesztes
                        
                        # Renaming the file
                        os.rename(old_name, new_name)
                        count += 1
                    
                    print('All Files Renamed')
                    print('New Names are')
                    # verify the result
                    res = os.listdir(folder_to)
                    res= Tcl().call('lsort', '-dict', res)
                    print(res)
            # rename based on fil name list in excel file
            
                    try:
                        # Get list of files in folder
        
                        # file_list3 = os.listdir(folder2)
                        file_list3 = Tcl().call('lsort', '-dict', os.listdir(folder2))
                        ff3 = [
                            f
                            for f in file_list3
                            if os.path.isfile(os.path.join(folder2, f))]
                        file_list3 = Tcl().call('lsort', '-dict', ff3)
                        try:
                            file_list3=list(file_list3)
                            file_list3.remove("Thumbs.db")
                            file_list3=tuple(file_list3)
                        except:
                            print("No Thumbs.db in folders")
                        T_number=str(len(file_list3))
        
                    except:
        
                        file_list3 = []
        
                    # fnames3= Tcl().call('lsort', '-dict', os.listdir(folder2))
                    fnames3 = [
                        f
                        for f in file_list3
                        if os.path.isfile(os.path.join(folder2, f))
                    ]
                    window["-FILE LIST2-"].update(fnames3)
                    window["T_number"].update(T_number)
                else:
                    print('Exit4')
                    sg.popup_error('Not the same file number')  # Shows red error button
                    #sys.exit()
        except Exception as e:
            print("not ok")
            sg.popup_error('Can not rename')  # Shows red error button
            print("expection:", e)
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            
    if event == "-UNDO-":
        
        try:
            print("Undo button pushed")
            folder_to=folder2+"\\"
            count=0
            file_listX = Tcl().call('lsort', '-dict', os.listdir(folder_to))
            ffX = [
                f
                for f in file_listX
                if os.path.isfile(os.path.join(folder_to, f))]
            file_listX = Tcl().call('lsort', '-dict', ffX)
            try:
                file_listX=list(file_listX)
                file_listX.remove("Thumbs.db")
                file_listX=tuple(file_listX)
            except:
                print("No Thumbs.db in file_listX")
            for file in file_listX:
            # for file in Tcl().call('lsort', '-dict', os.listdir(folder_to)):
                # construct current name using file name and path
                old_name2 = os.path.join(folder_to, file)
                # Construct old file name
                new_name2 = folder_to  + fnames2[count]
            
                # Renaming the file
                os.rename(old_name2, new_name2)
                count += 1
            try:
    
                # Get list of files in folder

                file_list4 = Tcl().call('lsort', '-dict', os.listdir(folder2))
                ff4 = [
                    f
                    for f in file_list4
                    if os.path.isfile(os.path.join(folder2, f))]
                file_list4 = Tcl().call('lsort', '-dict', ff4)
                try:
                    file_list4=list(file_list4)
                    file_list4.remove("Thumbs.db")
                    file_list4=tuple(file_list4)
                except:
                    print("No Thumbs.db in folders")
                T_number=str(len(file_list4))
    
            except:
                file_list4 = []

            fnames4 = [
                f
                for f in file_list4
                if os.path.isfile(os.path.join(folder2, f))
            ]
            
            window["-FILE LIST2-"].update(fnames4)
            window["T_number"].update(T_number)
            try:
                window["-IMAGE-"].update(data=get_img_data((folder2 + '\\' + fnames4[0]),  first=True))
                window["-IMAGE_zoom-"].update(zoom_at((folder2 + '\\' + fnames4[0]),  first=True))
            except Exception as e_i:
                print(e_i)
        except Exception as e:
            print("wrong undo")
            print(e)
    
    if event == "-EasterEgg-":
        try:
            print("EE button pushed")
            status=window["-love-"]
            #vis=status.visible
            if status.visible is True:
                window["-love-"].update(visible=False)
            else:
                window["-love-"].update(visible=True)
            #print("valami"+str(status.visible))
        except:
            print("nyuszi hopp")


window.close()
