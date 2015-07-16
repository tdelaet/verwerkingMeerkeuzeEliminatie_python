from Tkinter import *
from ttk import *
import os

def InputGUI ():
        myGUI = Tk()
        ##Opbouw GUI venster
        nameTest = StringVar()
        globalLabel=Label(myGUI, text= "This folder doesn't exist. Submit a new name.")
        FolderInputLabel=Label(myGUI, text= "Submit the name of the folder for the test.")
        title = "Multiple Choice with Elimination"
        FolderInput=Entry(myGUI, textvariable=nameTest)
        FolderSubmissionError = Label(myGUI, text ="The program will not work correctly because the interaction window was closed.")
        ##Actie voor knop
        def initiate():
            InputText = nameTest.get()
            globalLabel.pack_forget()
            if InputText != "verwerkingMeerkeuzeEliminatie_python" and InputText != "" and os.path.exists("../"+InputText):
                myGUI.geometry('320x100+300+150')
                myGUI.destroy()##Functie myGUI.quit() geeft veel kans op freeze van python => niet gebruiken
            else:
                myGUI.geometry('320x100+300+150')
                globalLabel.pack()
        submitButton=Button(myGUI, text = "Submit", command = initiate)
        ##Definitie weergave GUI
        myGUI.geometry('320x70+300+150')
        myGUI.title(title)
        FolderInputLabel.pack()
        FolderInput.pack()
        submitButton.pack()
        myGUI.lift()
        myGUI.mainloop()
        return nameTest.get()

Vragen = 0
Alternatieven = 0
TotaleScore = 0
output = {'questions':Vragen, 'alternatives':Alternatieven, 'totalscore':TotaleScore}

def VragenGUI():
        myGUI = Tk()
        ##Opbouw GUI venster
        numQuestions = IntVar()
        numAlternatives = IntVar()##Kan later een lijst worden op basis van aantal vragen
        maxTotalScore = IntVar()
        title = "Multiple Choice with Elimination"
        QuestionsLabel=Label(myGUI, text = "Enter the total number of questions below.")
        AlternativesLabel=Label(myGUI, text = "Enter the number of alternatives per question.")
        TotalScoreLabel=Label(myGUI, text = "Enter the total score students can obtain from this test.")
        OverviewLabel = Label(myGUI, text ="Overview of your input:")
        QuestionValueLabel = Label(myGUI, text = "")
        AlternativesValueLabel = Label (myGUI, text = "")
        TotalScoreValueLabel = Label (myGUI, text = "")
        QuestionsInput = Entry(myGUI, textvariable=numQuestions)
        AlternativesInput = Entry(myGUI, textvariable=numAlternatives)
        TotalScoreInput = Entry(myGUI, textvariable=maxTotalScore)
        def getQuestions():
                global Vragen
                Vragen = numQuestions.get()
                submitButton.pack_forget()
                QuestionsInput.pack_forget()
                QuestionsLabel.pack_forget()
                submitButton.config(command = getAlternatives)
                AlternativesLabel.pack()
                AlternativesInput.pack()
                submitButton.pack()
        def getAlternatives():
                global Alternatieven
                Alternatieven = numAlternatives.get()
                submitButton.pack_forget()
                AlternativesInput.pack_forget()
                AlternativesLabel.pack_forget()
                submitButton.config(command = getTotalScore)
                TotalScoreLabel.pack()
                TotalScoreInput.pack()
                submitButton.pack()
        def getTotalScore():
                global TotaleScore
                TotaleScore = maxTotalScore.get()
                submitButton.pack_forget()
                TotalScoreInput.pack_forget()
                TotalScoreLabel.pack_forget()
                submitButton.config(command = finishQuestionInput)
                myGUI.geometry('180x110+300+150')
                OverviewLabel.place(x=0,y=0)
                QuestionValueLabel.config(text ="Your test has " + str(numQuestions.get()) + " questions.")
                QuestionValueLabel.place(x=0,y=20)
                AlternativesValueLabel.config(text ="Each question has " + str(numAlternatives.get()) + " alternatives.")
                AlternativesValueLabel.place(x=0,y=40)
                TotalScoreValueLabel.config(text ="The maximum total score is " + str(maxTotalScore.get()))
                TotalScoreValueLabel.place(x=0,y=60)
                OkButton.place(x=0,y=80)
                AgainButton.place(x=100,y=80)
        def finishQuestionInput():
                global output
                output = {'questions':Vragen, 'alternatives':Alternatieven, 'totalscore':TotaleScore}
                myGUI.destroy()
        def adjustInput():
                global output
                myGUI.destroy()
                output = VragenGUI()
        submitButton=Button(myGUI, text = "Submit", command = getQuestions)
        OkButton=Button(myGUI, text = "OK", command = finishQuestionInput)
        AgainButton=Button(myGUI, text = "Adjust input", command = adjustInput)
        myGUI.geometry('320x70+300+150')
        myGUI.title(title)
        QuestionsLabel.pack()
        QuestionsInput.pack()
        submitButton.pack()
        myGUI.lift()
        myGUI.mainloop()
        return output
