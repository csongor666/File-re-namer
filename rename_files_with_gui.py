# -*- coding: utf-8 -*-
#%% what to do?
# rename from folder list

# from IPython import get_ipython
# get_ipython().magic('reset -sf')

import PySimpleGUI as sg

import os.path
#import sys
from tkinter import Tcl

# First the window layout in 2 columns

TT=sg.theme('Green Mono')
# d=sg.ListOfLookAndFeelValues()
# f=sg.SetOptions(button_color)
file_list_column = [

    [

        sg.Text("Folder with source names"),

        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),

        sg.FolderBrowse(),

    ],

    [

        sg.Listbox(

            values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"

        )

    ],

]

file_list_column2 = [

    [

        sg.Text("Folder to rename"),

        sg.In(size=(25, 1), enable_events=True, key="-FOLDER2-"),

        sg.FolderBrowse(),

    ],

    [

        sg.Listbox(

            values=[], enable_events=True, size=(40, 20), key="-FILE LIST2-"

        )

    ],

]

# ----- Full layout -----

layout = [

    [
        [
        #sg.Text("Title",size=(15, 1), font=("Verdana",11),text_color='Black',background_color='Yellow', justification='right
        sg.Text("Title",size=(7, 1),justification='center', text_color='white', background_color='#6D9F85'),
        sg.Input(size=(39, 1),key="-TITLE-"),
        sg.Button("Rename", button_color=('green'),enable_events=True,key="-RENAME-"),
        sg.Button("Undo", button_color=('green'), enable_events=True, key="-UNDO-"),
        sg.Text("                                                            "),
        sg.Button(" ",border_width=0, size=(1, 1), button_color=('#A8C1B4'), enable_events=True, key="-EasterEgg-"),
        sg.Text("Vera <3",font = ("Arial, 8"), text_color='green', background_color='#A8C1B4',visible=False, key="-love-")],
        
        sg.Column(file_list_column),

        sg.VSeperator(),

        sg.Column(file_list_column2),

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

            file_list = os.listdir(folder)

        except:

            file_list = []


        fnames = [

            f

            for f in file_list

            if os.path.isfile(os.path.join(folder, f))

            #and f.lower().endswith((".png", ".gif"))

        ]

        window["-FILE LIST-"].update(fnames)
        
    if event == "-FOLDER2-":

        folder2 = values["-FOLDER2-"]

        try:

            # Get list of files in folder

            file_list2 = os.listdir(folder2)

        except:

            file_list2 = []


        fnames2 = [

            f

            for f in file_list2

            if os.path.isfile(os.path.join(folder2, f))

            #and f.lower().endswith((".png", ".gif"))

        ]

        window["-FILE LIST2-"].update(fnames2)
        
    if event == "-RENAME-":
        try:
            print("Rename button pushed")
            # title=values["-TITLE-"]

            title=values["-TITLE-"]
            folder_to=[folder2+"\\"]
            folder_with_names=folder+"\\"
            for i in range(0,len(folder_to)):
                source_name=[]
                source_name=os.listdir(folder_with_names)
                source_name=Tcl().call('lsort', '-dict', source_name)
                
                counted_names=[]
                counted_names=os.listdir(folder_to[i])
                counted_names=Tcl().call('lsort', '-dict', counted_names)
                
                #remove Thumbs.db
                try:
                    source_name=list(source_name)
                    source_name=counted_names.remove("Thumbs.db")
                    source_name=tuple(source_name)
                    
                    counted_names=list(counted_names)
                    counted_names=counted_names.remove("Thumbs.db")
                    counted_names=tuple(counted_names)
                except:
                    print("No Thumbs.db in folders")
                
                
                # remove the file type from source
                source_name= [x[:-4] for x in source_name]
                
                # ask for file type
                kiterjesztes=counted_names[0][-4:]
                print('Lists are ready')
                
                #check the number of files
                if (len(counted_names) == len(source_name)):
                    print ( folder_to[i] + " and " + folder_with_names + " have the same number of files")
                    count = 0
                    # iterate all files from a directory
                    for file in Tcl().call('lsort', '-dict', os.listdir(folder_to[i])):
                        # construct current name using file name and path
                        old_name = os.path.join(folder_to[i], file)
                        # Construct old file name
                        new_name = folder_to[i]  + source_name[count] + title + kiterjesztes
                    
                        # Renaming the file
                        os.rename(old_name, new_name)
                        count += 1
                    
                    print('All Files Renamed')
                    print('New Names are')
                    # verify the result
                    res = os.listdir(folder_to[i])
                    res= Tcl().call('lsort', '-dict', res)
                    print(res)
                    try:

                        # Get list of files in folder

                        file_list3 = os.listdir(folder2)

                    except:

                        file_list3 = []


                    fnames3 = [

                        f

                        for f in file_list3

                        if os.path.isfile(os.path.join(folder2, f))
                    ]
                    window["-FILE LIST2-"].update(fnames3)
                else:
                    print('Exit')
                    sg.popup_error('Not the same file number')  # Shows red error button
                    #sys.exit()
        except:
            print("not ok")
            
    if event == "-UNDO-":
        
        try:
            print("Undo button pushed")
            folder_to=[folder2+"\\"]
            #folder_with_names=folder+"\\"
            count=0
            for file in Tcl().call('lsort', '-dict', os.listdir(folder_to[i])):
                # construct current name using file name and path
                old_name2 = os.path.join(folder_to[i], file)
                # Construct old file name
                new_name2 = folder_to[i]  + fnames2[count]
            
                # Renaming the file
                os.rename(old_name2, new_name2)
                count += 1
            try:
    
                # Get list of files in folder
    
                file_list4 = os.listdir(folder2)
    
            except:
    
                file_list4 = []
    
    
            fnames4 = [
    
                f
    
                for f in file_list4
    
                if os.path.isfile(os.path.join(folder2, f))
            ]
            window["-FILE LIST2-"].update(fnames4)
        except:
            print("wrong undo")
    
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
