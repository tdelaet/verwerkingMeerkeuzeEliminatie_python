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
        nameTest.set('2015-juni')
        ##Actie voor knop (ingeven van map waar data staan)
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
        numQuestions = IntVar()
        numQuestions.set(26)
        questionAlternatives = []
#Alternatives can be different for every question. For now the output will only return the number of alternatives for question 1.
        numAlternatives = dict()
        AlternativesInput = dict()                                          
        maxTotalScore = IntVar()
        maxTotalScore.set(20)
        title = "Multiple Choice with Elimination"
        QuestionsLabel=Label(myGUI, text = "Submit the total number of questions below.")
        AlternativesLabel1=Label(myGUI, text = "Submit the number of alternatives per question. The default value is set to 4.")
        AlternativesLabel2=Label(myGUI, text = "Change the default by changing the value for Question 1, then press 'Copy Question 1'.")
        TotalScoreLabel=Label(myGUI, text = "Submit the total score students can obtain from this test.")
        OverviewLabel = Label(myGUI, text ="Overview of your input:")
        QuestionValueLabel = Label(myGUI, text = "")
        AlternativesValueLabel = dict()
        TotalScoreValueLabel = Label (myGUI, text = "")
        QuestionsNegativeLabel = Label (myGUI, text = "The number of questions must be positive.")
        QuestionsInput = Entry(myGUI, textvariable=numQuestions)
        TotalScoreInput = Entry(myGUI, textvariable=maxTotalScore)
        def getQuestions(): #Submit total number of questions. Default value is 26.
                global Vragen
                QuestionsNegativeLabel.pack_forget()
                Vragen = numQuestions.get()
                if numQuestions.get() < 0:
                        myGUI.geometry('260x90+300+150')
                        QuestionsNegativeLabel.pack()
                else:
                        submitButton.pack_forget()
                        QuestionsInput.pack_forget()
                        QuestionsLabel.pack_forget()
                        submitButton.config(command = getAlternatives)
                        AlternativesLabel1.pack()
                        AlternativesLabel2.pack()
                        for x in range(0, numQuestions.get()):
                                questionAlternatives.append(Label(myGUI,text='Question ' + str(x+1)))
                                numAlternatives[x]=IntVar()
                                numAlternatives[x].set(4)
                                AlternativesInput[x]=Entry(myGUI, textvariable=numAlternatives[x]) 
                        myGUI.geometry('500x'+str(70+22*((numQuestions.get()+1)/2))+'+300+150')
                        for x in range(0, int(round((numQuestions.get()+1)/2))):
                                questionAlternatives[x].place(x=0,y=22*(x+2))
                                AlternativesInput[x].place(x=100,y=22*(x+2))
                        for x in range(int(round((numQuestions.get()+1)/2)),numQuestions.get()):
                                questionAlternatives[x].place(x=250,y=22*(x-round((numQuestions.get()+1)/2)+2))
                                AlternativesInput[x].place(x=350,y=22*(x-round((numQuestions.get()+1)/2)+2))
                        copyButton.place(x=100,y=22*(round((numQuestions.get()+1)/2)+2))
                        submitButton.place(x=350,y=22*(round((numQuestions.get()+1)/2)+2))
        def getAlternatives(): #Submit alternatives for each questions. Default value is 4.
                global Alternatieven
                Alternatieven = numAlternatives[0].get()
                submitButton.place_forget()
                copyButton.place_forget()
                AlternativesInput[0].pack_forget()
                AlternativesLabel1.pack_forget()
                AlternativesLabel2.pack_forget()
                myGUI.geometry('320x70+300+150')
                for x in range(0, numQuestions.get()):
                        questionAlternatives[x].place_forget()
                        AlternativesInput[x].place_forget()
                submitButton.config(command = getTotalScore)
                TotalScoreLabel.pack()
                TotalScoreInput.pack()
                submitButton.pack()
        def changeAlternatives(): #Change default value of number of alternatives for all questions.
                alternatives = numAlternatives[0].get()
                for x in range(1, numQuestions.get()):
                        numAlternatives[x].set(alternatives)
        def getTotalScore(): #Submit total score the test is on.
                global TotaleScore
                TotaleScore = maxTotalScore.get()
                submitButton.pack_forget()
                TotalScoreInput.pack_forget()
                TotalScoreLabel.pack_forget()
                submitButton.config(command = finishQuestionInput)
                myGUI.geometry('430x'+str(90+((numQuestions.get()+1)/2)*20)+'+300+150')
                OverviewLabel.place(x=0,y=0)
                QuestionValueLabel.config(text ="Your test has " + str(numQuestions.get()) + " questions.")
                QuestionValueLabel.place(x=0,y=20)
                for x in range(0, int(round((numQuestions.get()+1)/2))):
                        AlternativesValueLabel[x]=Label(myGUI, text =questionAlternatives[x].cget("text") + " has " + str(numAlternatives[x].get()) + " alternatives.")
                        AlternativesValueLabel[x].place(x=0,y=(40+20*x))
                for x in range(int(round((numQuestions.get()+1)/2)),numQuestions.get()):
                        AlternativesValueLabel[x]=Label(myGUI, text =questionAlternatives[x].cget("text") + " has " + str(numAlternatives[x].get()) + " alternatives.")
                        AlternativesValueLabel[x].place(x=200,y=(40+20*(x-round((numQuestions.get()+1)/2))))
                TotalScoreValueLabel.config(text ="The maximum total score is " + str(maxTotalScore.get()) + ".")
                TotalScoreValueLabel.place(x=0,y=40+20*(round((numQuestions.get()+1)/2)))
                OkButton.place(x=0,y=60+20*(round((numQuestions.get()+1)/2)))
                AgainButton.place(x=100,y=60+20*(round((numQuestions.get()+1)/2)))
        def finishQuestionInput(): #Return number of questions, alternatives and score
                global output
                output = {'questions':Vragen, 'alternatives':Alternatieven, 'totalscore':TotaleScore}
                myGUI.destroy()
        def adjustInput(): #Recursively restart entering data
                global output
                myGUI.destroy()
                output = VragenGUI()
#Building the initial GUI
        submitButton=Button(myGUI, text = "Submit", command = getQuestions)
        copyButton=Button(myGUI, text = "Copy Question 1", command = changeAlternatives)
        OkButton=Button(myGUI, text = "OK", command = finishQuestionInput)
        AgainButton=Button(myGUI, text = "Adjust input", command = adjustInput)
        myGUI.geometry('260x70+300+150')
        myGUI.title(title)
        QuestionsLabel.pack()
        QuestionsInput.pack()
        submitButton.pack()
        myGUI.lift()
        myGUI.mainloop()
        return output
