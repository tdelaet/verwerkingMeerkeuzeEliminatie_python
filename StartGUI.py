from Tkinter import *
import os


myGUI = Tk()
nameTest = StringVar()
globalLabel=Label(myGUI, text= "This map doesn't exist. Enter a new name.")
def initiate():
    InputText = nameTest.get()
    globalLabel.pack_forget()
    if InputText != "verwerkingMeerkeuzeEliminatie_python" and os.path.exists("../"+InputText):
        myGUI.destroy()
    else:
        globalLabel.pack()

def StartGUI (number):
    if number == 1:
        myGUI.geometry('250x250+500+300')
        myGUI.title('Verwerking Meerkeuze Eliminatie')
        Label(myGUI, text= "Enter the name of the map for the test").pack()
        Entry(myGUI, textvariable=nameTest).pack()
        Button(myGUI, activebackground = 'black', activeforeground ='white', text = "Submit", command = initiate).pack()
        myGUI.lift()
        myGUI.mainloop()
        return nameTest.get()
#Menu's in de GUI??
##GUImenu = Menu(myGUI)
##mymenu = Menu(GUImenu)
##mymenu.add_command(label="FileMap")
##mymenu.add_command(label="Questions")
##mymenu.add_command(label="AnswerPossibilities")
##mymenu.add_command(label="Other inputs")
##GUImenu.add_cascade(label="Options",menu=mymenu)
##myGUI.config(menu=GUImenu)


