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
Alternatieven = dict()
Permutaties = 0
TotaleScore = 0
output = {'questions':Vragen, 'alternatives':Alternatieven, 'permutations':Permutaties, 'totalscore':TotaleScore}

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
        permutations = IntVar()
        permutations.set(4)
        title = "Multiple Choice with Elimination"
        QuestionsLabel=Label(myGUI, text = "Submit the total number of questions below.")
        AlternativesLabel1=Label(myGUI, text = "Submit the number of alternatives per question. The default value is set to 4.")
        AlternativesLabel2=Label(myGUI, text = "Change the default by changing the value for Question 1, then press 'Copy Question 1'.")
        TotalScoreLabel=Label(myGUI, text = "Submit the total score students can obtain from this test.")
        PermutationsLabel=Label(myGUI, text = "Submit the total number of permutations in your test.")
        OverviewLabel = Label(myGUI, text ="Overview of your input:")
        QuestionValueLabel = Label(myGUI, text = "")
        AlternativesValueLabel = dict()
        PermutationsValueLabel = Label(myGUI, text ="")
        TotalScoreValueLabel = Label (myGUI, text = "")
        QuestionsNegativeLabel = Label (myGUI, text = "The number of questions must be positive.")
        QuestionsInput = Entry(myGUI, textvariable=numQuestions)
        TotalScoreInput = Entry(myGUI, textvariable=maxTotalScore)
        PermutationsInput = Entry(myGUI, textvariable=permutations)
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
                        myGUI.geometry('750x'+str(70+22*((numQuestions.get()-1)/3+1))+'+300+150')
                        counterCol = 0
                        counterRow = 0
                        for x in range(0,numQuestions.get()):
                                questionAlternatives[x].place(x=counterCol*250, y = 22*(counterRow+2))
                                AlternativesInput[x].place(x=100+250*counterCol, y = 22*(counterRow+2))
                                if counterCol == 2:
                                        counterRow +=1
                                        counterCol = 0
                                else:
                                        counterCol +=1
                        copyButton.place(x=100,y=22*(round((numQuestions.get()-1)/3)+3))
                        submitButton.place(x=600,y=22*(round((numQuestions.get()-1)/3)+3))
        def getAlternatives(): #Submit alternatives for each questions. Default value is 4.
                global Alternatieven
                for x in range (0,numQuestions.get()):
                        Alternatieven[x+1] = numAlternatives[x].get()
                submitButton.place_forget()
                copyButton.place_forget()
                AlternativesInput[0].pack_forget()
                AlternativesLabel1.pack_forget()
                AlternativesLabel2.pack_forget()
                myGUI.geometry('320x70+300+150')
                for x in range(0, numQuestions.get()):
                        questionAlternatives[x].place_forget()
                        AlternativesInput[x].place_forget()
                submitButton.config(command = getPermutations)
                PermutationsLabel.pack()
                PermutationsInput.pack()
                submitButton.pack()
        def changeAlternatives(): #Change default value of number of alternatives for all questions.
                alternatives = numAlternatives[0].get()
                for x in range(1, numQuestions.get()):
                        numAlternatives[x].set(alternatives)
        def getPermutations():
                global Permutaties
                Permutaties = permutations.get()
                submitButton.pack_forget()
                PermutationsLabel.pack_forget()
                PermutationsInput.pack_forget()
                submitButton.config(command = getTotalScore)
                TotalScoreLabel.pack()
                TotalScoreInput.pack()
                submitButton.pack()
        def getTotalScore(): #Submit total score the test is on.
                global TotaleScore
                TotaleScore = maxTotalScore.get()
                submitButton.pack_forget()
                TotalScoreInput.pack_forget()
                TotalScoreLabel.pack_forget()
                submitButton.config(command = finishQuestionInput)
                OverviewLabel.place(x=0,y=0)
                QuestionValueLabel.config(text ="Your test has " + str(numQuestions.get()) + " questions.")
                QuestionValueLabel.place(x=0,y=20)
                myGUI.geometry('600x'+str(130+((numQuestions.get()-1)/3)*20)+'+300+150')
                counterCol = 0
                counterRow = 0
                for x in range(0,numQuestions.get()):
                        AlternativesValueLabel[x]=Label(myGUI, text =questionAlternatives[x].cget("text") + " has " + str(numAlternatives[x].get()) + " alternatives.")
                        AlternativesValueLabel[x].place(x=200*counterCol,y=(40+20*counterRow))
                        if counterCol == 2:
                                counterRow +=1
                                counterCol = 0
                        else:
                                counterCol +=1
                PermutationsValueLabel.config(text="The number of permutations in your test is " + str(permutations.get())+".")
                PermutationsValueLabel.place(x=0,y=40+20*(round((numQuestions.get()-1)/3+1)))
                TotalScoreValueLabel.config(text ="The maximum total score is " + str(maxTotalScore.get()) + ".")
                TotalScoreValueLabel.place(x=0,y=40+20*(round((numQuestions.get()-1)/3+2)))
                OkButton.place(x=0,y=60+20*(round((numQuestions.get()-1)/3+2)))
                AgainButton.place(x=100,y=60+20*(round((numQuestions.get()-1)/3+2)))
        def finishQuestionInput(): #Return number of questions, alternatives and score
                global output
                output = {'questions':Vragen, 'alternatives':Alternatieven, 'permutations':Permutaties, 'totalscore':TotaleScore}
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
