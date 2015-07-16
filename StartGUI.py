from Tkinter import *
from ttk import *
import os

myGUI = Tk()
nameTest = StringVar()
globalLabel=Label(myGUI, text= "This folder doesn't exist. Enter a new name.")
FolderInputLabel=Label(myGUI, text= "Enter the name of the folder for the test.")
title = "Multiple Choice with Elimination"
StandardLanguage = "Language"
Language = "English"
FolderInput=Entry(myGUI, textvariable=nameTest)

def initiate():
    InputText = nameTest.get()
    globalLabel.pack_forget()
    if InputText != "verwerkingMeerkeuzeEliminatie_python" and InputText != "" and os.path.exists("../"+InputText):
        myGUI.geometry('320x100+300+150')
        myGUI.destroy()
    else:
        myGUI.geometry('320x100+300+150')
        globalLabel.pack()

submitButton=Button(myGUI, text = "Submit", command = initiate)

def changeLanguage():
    global globalLabel
    global FolderInputLabel
    global title
    global StandardLanguage
    if Language == "English":
        submitButton.pack_forget()
        FolderInput.pack_forget()
        FolderInputLabel.pack_forget()
        globalLabel.pack_forget()
        myGUI.geometry('320x70+300+150')
        globalLabel=Label(myGUI, text= "This folder doesn't exist. Enter a new name.")
        FolderInputLabel=Label(myGUI, text= "Enter the name of the folder for the test.")
        title = "Multiple Choice with Elimination"
        StandardLanguage = "Language"
        myGUI.title(title)
        FolderInputLabel.pack()
        FolderInput.pack()
        submitButton.pack()
    elif Language == "Dutch":
        submitButton.pack_forget()
        FolderInput.pack_forget()
        FolderInputLabel.pack_forget()
        globalLabel.pack_forget()
        myGUI.geometry('320x70+300+150')
        globalLabel=Label(myGUI, text= "Deze map bestaat niet. Geef een nieuwe naam in.")
        FolderInputLabel=Label(myGUI, text= "Geef de naam van de map van de test in.")
        title = "Meerkeuze met Eliminatie"
        StandardLanguage = "Taal"
        myGUI.title(title)
        FolderInputLabel.pack()
        FolderInput.pack()
        submitButton.pack()

def dutch():
    global Language
    Language = "Dutch"
    changeLanguage()

def english():
    global Language
    Language = "English"
    changeLanguage()

def StartGUI (number):
    global Language
    if number == 1:
        myGUI.geometry('320x70+300+150')
        myGUI.title(title)
        FolderInputLabel.pack()
        FolderInput.pack()
        submitButton.pack()
        GUImenu = Menu(myGUI)
        languageMenu = Menu(GUImenu)
        languageMenu.add_command(label="English", command = english)
        languageMenu.add_command(label="Nederlands", command = dutch)
        GUImenu.add_cascade(label=StandardLanguage,menu=languageMenu)
        myGUI.config(menu=GUImenu)
        myGUI.lift()
        myGUI.mainloop()
        return nameTest.get()
    elif number == 2:
        pass
        return
    elif number == 3:
        pass
        return
    else:
        pass
        return
