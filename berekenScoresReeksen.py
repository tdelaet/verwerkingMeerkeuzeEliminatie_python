# -*- coding: utf-8 -*-
"""
Created on Mon Mar 31 13:58:48 2014

@author: tdelaet
"""

from xlrd import open_workbook, colname
import string
import numpy
import matplotlib.pyplot as plt
from xlwt import Workbook, easyxf


#nameFile = "OMRoutputEN_together.xlsx"
nameFile = "test.xlsx"
nameSheet = "outputScan"
numQuestions = 10
numAlternatives = 4
maxTotalScore = 20


correctAnswers = [["B","B","D","B","A","A","B","D","A","C"], #sessie 1
["B","D","B","A","A","B","D","A","C","B"]] #sessie 2

twoOptions=["onmogelijk","mogelijk"]

plt.close("all")


#letters of answer alternatives
alternatives = list(string.ascii_uppercase)[0:numAlternatives]

############################
#create list of expected content of scan file
content = ["Deelnemersnummer","Deelnemersnummer_bar","vragenreeks"]
for question in xrange(1,numQuestions+1):
    for alternative in alternatives:
        name = "vraag" + str(question) + alternative
        content.append(name)
#print content
###########################
        
def all_indices(value, qlist):
    indices = []
    idx = -1
    while True:
        try:
            idx = qlist.index(value, idx+1)
            indices.append(idx)
        except ValueError:
            break
    return indices


# write to excel_file
outputbook = Workbook()

style_header = easyxf("font: bold on; align: horiz center; border: bottom medium")
style_header_borderRight = easyxf("font: bold on; align: horiz center; border: right medium ")
style_correctAnswer = easyxf('font: color blue')
style_specialAttention  = easyxf('pattern: pattern solid, fore_colour red')
style_correctAnswerSpecialAttention = easyxf('font: color blue;''pattern: pattern solid, fore_colour red')

style_border_header = easyxf('border: left thick, top thick, bottom thick, right thick')


##TODO: behaalde score in andere reeksen en totale score
#number of series
numSeries = len(correctAnswers)

## check errors in input file
# check if the number of questions is equal to the length of the list of correct answers
for serie in xrange(numSeries):
    if numQuestions != len(correctAnswers[serie]):
        print "The length of the list of correct answers is not equal to the number of questions"
        #sys.exit()
    for answer in correctAnswers[serie]:
        if not(answer in alternatives):
            print "The answer " + answer +  " of series " + str(serie+1) + " is not in the list of possible answer alternatives"
            #sys.exit()


# read file and get sheet
book= open_workbook(nameFile)
sheet = book.sheet_by_name(nameSheet)

#number of rows and columns
num_rows = sheet.nrows;
num_cols = sheet.ncols;

#number of participants = number of rows-1
numParticipants = num_rows-1;

#Load the first row => name indicating 
firstRow =  sheet.row(0) 
firstRowValues = sheet.row_values(0)
#print firstRowValues

content_colNrs = [0] * len(content);
#Look for column numbers of expected content
index =0
for x in content:
    content_colNrs[index] = firstRowValues.index(x)
    index = index+1
#print content_colNrs

# Calculate the score for each question
scoreQuestions= numpy.zeros((numSeries,numParticipants,numQuestions))
numOnmogelijkQuestionsAlternatives= numpy.zeros(numQuestions*numAlternatives)
numMogelijkQuestionsAlternatives= numpy.zeros(numQuestions*numAlternatives)


#loop over question
counter = 0;
for question in xrange(1,numQuestions+1):
    #print "vraag" + str(question)

    #loop over alternatives
    for alternative in alternatives:
        name = "vraag" + str(question) + alternative
        #print name
        #find column number in which question alternatives are given
        colNr = content_colNrs[content.index(name)]
        #get the answers for the participants (so skip for row with name of first row)
        columnAlternative=sheet.col_values(colNr,1,num_rows)
        indicesNotMogelijkOnmogelijk = [x for x in range(0,len(columnAlternative)) if not(columnAlternative[x] in twoOptions)]
        if len(indicesNotMogelijkOnmogelijk): #there are entries that do not correspond to the options
            row_withErrors = [x+1 for x in indicesNotMogelijkOnmogelijk]
            print "ERROR FOUND " + "row index " + str(row_withErrors) + " column index: " + colname(colNr)       
            ##TODO:STOP PROGRAM
        #get the indices of the particpants that answered onmogelijk
        indicesOnmogelijk = [x for x in range(0,len(columnAlternative)) if columnAlternative[x]=="onmogelijk"]
        numOnmogelijkQuestionsAlternatives[counter]=len(indicesOnmogelijk)
        indicesMogelijk = [x for x in range(0,len(columnAlternative)) if columnAlternative[x]=="mogelijk"]  
        numMogelijkQuestionsAlternatives[counter]=len(indicesMogelijk)        
        counter+=1
        #only needed to calculate score for different series
        #loop over series
        for serie in xrange(1,numSeries+1):
            #print "correct answer is " + correctAnswers[serie][question-1]
            if  correctAnswers[serie-1][question-1] == alternative:         # if the alternative is the correct answer => wrongly excluded so -1
                scoreQuestions[serie-1,indicesOnmogelijk,question-1]-=1.0
            else: # if the alternative is NOT the correct answer => correctly excluded so +1/(numAlternatives-1)
                scoreQuestions[serie-1,indicesOnmogelijk,question-1]+= 1.0/(float(numAlternatives)-1.0)
        
#print scoreQuestions
name = "vragenreeks"
#get the column in which the vragenreeks is stored
colNrSerie = content_colNrs[content.index(name)]
#get the series for the participants (so skip for row with name of first row)
columnSeries=sheet.col_values(colNrSerie,1,num_rows)

#print columnSeries
scoreQuestionsIndicatedSeries= numpy.zeros((numParticipants,numQuestions))
for participant in xrange(numParticipants):
    serieIndicated = columnSeries[participant]
    scoreQuestionsIndicatedSeries[participant,:] = scoreQuestions[serieIndicated-1,participant,:] 
    
#To calculate the total score only use the score for the series indicated by the student
totalScore = scoreQuestionsIndicatedSeries.sum(axis=1)/numQuestions*maxTotalScore
# set negative scores to 0
totalScore[totalScore < 0]=0
totalScore = numpy.round(totalScore)
#print totalScore

# Get the overall average and median score
averageScore = numpy.zeros(numSeries)
medianScore = numpy.zeros(numSeries)

#print averageScore
#print medianScore
    
# write for each participant the total score for the different series

sheetC = outputbook.add_sheet('ScoreVerschillendeSeries')

#column counter
columnCounter = 0;

#deelnemersnummers
deelnemersBarColNr= content_colNrs[content.index("Deelnemersnummer_bar")]
deelnemers=sheet.col_values(deelnemersBarColNr,columnCounter,num_rows)
#print deelnemers
sheetC.write(0, 0,"deelnemersnummer", style=style_header_borderRight) 
for i in xrange(1,len(deelnemers)):
    sheetC.write(i,columnCounter,deelnemers[i], style=style_header_borderRight)
columnCounter+=1;

#total score for indicated series
sheetC.write(0,columnCounter,"totale score aangeduide serie ",style=style_header)
offsetRow=1
for i in xrange(len(totalScore)):
    sheetC.write(i+offsetRow,columnCounter,totalScore[i])
columnCounter+=1;

#total score for different series
for serie in xrange(1,numSeries+1):
    print serie
    sheetC.write(0,columnCounter,"totale score serie " + str(serie),style=style_header)
    offsetRow=1
    totalScoreSerie = scoreQuestions[serie-1].sum(axis=1)/numQuestions*maxTotalScore
    totalScoreSerie[totalScoreSerie < 0]=0
    totalScoreSerie = numpy.round(totalScoreSerie)

    for i in xrange(len(totalScoreSerie)):
        # if the series is the same as the one indicated 
        if (serie == columnSeries[i]):
            sheetC.write(i+offsetRow,columnCounter,totalScoreSerie[i],style=style_correctAnswer)
        else:# the series is different fromthe one indicated 
            if (totalScoreSerie[i]>totalScore[i]):  # if score on other serie than the one indicated is higer        
                sheetC.write(i+offsetRow,columnCounter,totalScoreSerie[i],style=style_specialAttention)
            else:
                sheetC.write(i+offsetRow,columnCounter,totalScoreSerie[i])
    columnCounter+=1;    


## more calculations
for serie in xrange(1,numSeries+1):
    #get all the participants that have indicated this seris
    participantsOfSerie = all_indices(serie,columnSeries)
    numParticipantsSerie = len(participantsOfSerie)
    print "number of participants serie " + str(serie) + " is " + str(numParticipantsSerie)
    #get the scores of the questions of the participants that are in the serie
    scoreQuestionsSerie = scoreQuestions[serie-1,participantsOfSerie]
    #print scoreQuestionsSerie
    #get the total score of that serie
    totalScoreSerie = totalScore[participantsOfSerie]
    averageScoreSerie = sum(totalScoreSerie)/float(numParticipantsSerie)
    medianScoreSerie = numpy.median(totalScoreSerie)
    #print totalScoreSerie
    
    #averageScoreQuestions
    averageScoreQuestions = scoreQuestionsSerie.sum(axis=0)/float(numParticipants)
    ##TODO:STD

    #find upper/middle/lower third of the students
    #sort students according to score
    orderedDeelnemers = sorted(range(len(totalScoreSerie)),key=totalScoreSerie.__getitem__) 
    third = int(numpy.ceil(numParticipantsSerie/3))
    indicesUpper = orderedDeelnemers[numParticipantsSerie-third:numParticipantsSerie]
    indicesLower = orderedDeelnemers[0:third]
    indicesMiddle= orderedDeelnemers[third:numParticipantsSerie-third]
    numUpper = len(indicesUpper)
    numLower = len(indicesLower)
    numMiddle = len(indicesMiddle)
    #print totalScore[indicesUpper]
    #print totalScore[indicesLower]
    #print totalScore[indicesMiddle]
    #print len(indicesUpper)
    #print len(indicesLower)
    #print len(indicesMiddle)
    #print len(orderedDeelnemers)


    #averageScoreQuestions Upper/Lower/Middle
    totalScoreUpper = totalScoreSerie[indicesUpper]
    totalScoreMiddle = totalScoreSerie[indicesMiddle]
    totalScoreLower = totalScoreSerie[indicesLower]
    averageScoreUpper = sum(totalScoreUpper)/float(numUpper)
    averageScoreMiddle = sum(totalScoreMiddle)/float(numMiddle)
    averageScoreLower = sum(totalScoreLower)/float(numLower)
    scoreQuestionsUpper =  scoreQuestionsSerie[indicesUpper,:]
    scoreQuestionsMiddle =  scoreQuestionsSerie[indicesMiddle,:]
    scoreQuestionsLower =  scoreQuestionsSerie[indicesLower,:]
    averageScoreQuestionsUpper = scoreQuestionsUpper.sum(axis=0)/float(numUpper)
    averageScoreQuestionsMiddle = scoreQuestionsMiddle.sum(axis=0)/float(numMiddle)
    averageScoreQuestionsLower = scoreQuestionsLower.sum(axis=0)/float(numLower)

    numOnmogelijkQuestionsAlternatives= numpy.zeros(numQuestions*numAlternatives)
    numOnmogelijkQuestionsAlternativesUpper= numpy.zeros(numQuestions*numAlternatives)
    numMogelijkQuestionsAlternativesUpper= numpy.zeros(numQuestions*numAlternatives)
    numOnmogelijkQuestionsAlternativesMiddle= numpy.zeros(numQuestions*numAlternatives)
    numMogelijkQuestionsAlternativesMiddle= numpy.zeros(numQuestions*numAlternatives)
    numOnmogelijkQuestionsAlternativesLower= numpy.zeros(numQuestions*numAlternatives)
    numMogelijkQuestionsAlternativesLower= numpy.zeros(numQuestions*numAlternatives)

    
    #number of onmogelijk in upper group per question
    #loop over question
    counter = 0;
    for question in xrange(1,numQuestions+1):
        print "vraag" + str(question)
        correctAnswer = correctAnswers[serie-1][question-1]
        #print "correct answer is " + correctAnswer
        #loop over alternatives
        for alternative in alternatives:
            name = "vraag" + str(question) + alternative
            #print name
            #find column number in which question alternatives are given
            colNr = content_colNrs[content.index(name)]
            #get the answers for the participants (so skip for row with name of first row)
            columnAlternative=sheet.col_values(colNr,1,num_rows)
            columnAlternative = [columnAlternative[x] for x in participantsOfSerie]
            ##ERROR: only get the ones for the one in this series
            indicesOnmogelijk = [x for x in range(0,len(columnAlternative)) if columnAlternative[x]=="onmogelijk"]
            numOnmogelijkQuestionsAlternatives[counter]=len(indicesOnmogelijk)
            indicesMogelijk = [x for x in range(0,len(columnAlternative)) if columnAlternative[x]=="mogelijk"]  
            numMogelijkQuestionsAlternatives[counter]=len(indicesMogelijk)    
            #get the indices of the particpants that answered onmogelijk
            indicesOnmogelijkUpper = [x for x in indicesUpper if columnAlternative[x]=="onmogelijk"]
            indicesOnmogelijkMiddle = [x for x in indicesMiddle if columnAlternative[x]=="onmogelijk"]
            indicesOnmogelijkLower = [x for x in indicesLower if columnAlternative[x]=="onmogelijk"]
            print len(indicesOnmogelijkUpper)
            print len(indicesOnmogelijkMiddle)
            #rint len(indicesOnmogelijkLower)
            #indicesOnmogelijk = [x for x in range(0,len(columnAlternative)) if columnAlternative[x]=="onmogelijk"]
            #print len(indicesOnmogelijk)
            numOnmogelijkQuestionsAlternativesUpper[counter]=len(indicesOnmogelijkUpper)
            numOnmogelijkQuestionsAlternativesMiddle[counter]=len(indicesOnmogelijkMiddle)
            numOnmogelijkQuestionsAlternativesLower[counter]=len(indicesOnmogelijkLower)
            
            indicesMogelijkUpper = [x for x in indicesUpper if columnAlternative[x]=="mogelijk"]  
            indicesMogelijkMiddle = [x for x in indicesMiddle if columnAlternative[x]=="mogelijk"]  
            indicesMogelijkLower = [x for x in indicesLower if columnAlternative[x]=="mogelijk"]  
            numMogelijkQuestionsAlternativesUpper[counter]=len(indicesMogelijkUpper)        
            numMogelijkQuestionsAlternativesMiddle[counter]=len(indicesMogelijkMiddle)        
            numMogelijkQuestionsAlternativesLower[counter]=len(indicesMogelijkLower)        
            counter+=1

    #plot statistics for each question
    #plot the histogram of the total score
    plt.figure()
    n, bins, patches = plt.hist(totalScoreSerie,bins=numpy.arange(0-0.5,maxTotalScore+1,1))
    plt.title("histogram total score")
    plt.xlabel("score (max " + str(maxTotalScore)+ ")")
    plt.xlim([0-0.5,maxTotalScore+0.5])
    plt.ylabel("number of students")
    figManager = plt.get_current_fig_manager()
    figManager.window.showMaximized()
    plt.savefig('histogramGeheel.png', bbox_inches='tight')
    
    # plot the quintiles of the total score
    plt.figure()
    n, bins, patches = plt.hist(totalScoreSerie,bins=numpy.arange(0-0.5,maxTotalScore+1,maxTotalScore/5.0))
    plt.title("histogram total score")
    plt.xlabel("score (max " + str(maxTotalScore)+ ")")
    plt.xlim([0-0.5,maxTotalScore+0.5])
    plt.ylabel("number of students")
    figManager = plt.get_current_fig_manager()
    figManager.window.showMaximized()
    plt.savefig('quintielenGeheel.png', bbox_inches='tight')
    
    #plot histogram for different questions
    numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
    #print numColsPict
    numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) +1
    #print numRowsPict
    fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
    fig.tight_layout() # Or equivalently,  "plt.tight_layout()"
    
    for question in xrange(1,numQuestions+1):
        ax = plt.subplot(numRowsPict,numColsPict,question)
        n, bins, patches = plt.hist(scoreQuestionsSerie[:,question-1],bins=numpy.arange(-1-1.0/6.0,1+1.0/3.0,1.0/3.0))
        plt.xticks(numpy.arange(-1, 1+1.0/3.0, 2*1.0/3.0))
        plt.title("vraag " + str(question))
        plt.xlabel("score (max " + str(maxTotalScore)+ ")")
        plt.xlim([-1-1.0/3.0,1+1.0/3.0])
        plt.ylabel("number of students")
    figManager = plt.get_current_fig_manager()
    figManager.window.showMaximized()    
    plt.savefig('histogramVragen.png', bbox_inches='tight')
    
    ##sheet: global paramters
    sheetC = outputbook.add_sheet('globaleParameters_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    #gemiddelde verdeling scores per vraag
    rowCounter = 0 
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=style_header)
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,numParticipantsSerie)
    
    rowCounter+=1
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=style_header)
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,averageScoreSerie)
    
    rowCounter+=1
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"mediaan ",style=style_header)
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,medianScoreSerie)
    
    rowCounter+=1
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=style_header)
    columnCounter+=1  
    aantalGeslaagd = sum(score>= maxTotalScore/2.0 for score in totalScoreSerie)
    sheetC.write(rowCounter,columnCounter,float(aantalGeslaagd)/float(numParticipantsSerie))

    #Sheet: Score per Deelnemer
    sheetC = outputbook.add_sheet('ScorePerDeelnemer_serie' + str(serie) )
    
    #column counter
    columnCounter = 0;
    
    #deelnemersnummers
    deelnemersBarColNr= content_colNrs[content.index("Deelnemersnummer_bar")]
    deelnemers=sheet.col_values(deelnemersBarColNr,start_rowx=1, end_rowx=num_rows)
    deelnemers= [deelnemers[x] for x in participantsOfSerie]
    #print deelnemers
    sheetC.write(0, 0,"deelnemersnummer", style=style_header_borderRight) 
    for i in xrange(0,len(deelnemers)):
        sheetC.write(i+1,columnCounter,deelnemers[i], style=style_header_borderRight)
    columnCounter+=1;
        
    #totale score
    sheetC.write(0,columnCounter,"totale score",style=style_header)
    offsetRow=1
    for i in xrange(len(totalScoreSerie)):
        sheetC.write(i+offsetRow,columnCounter,totalScoreSerie[i])
    columnCounter+=1;

    #score per vraag
    for question in xrange(1,numQuestions+1):
        sheetC.write(0,columnCounter,"vraag " + str(question),style=style_header)
        offsetRow=1
        for i in xrange(len(scoreQuestionsSerie)):
            sheetC.write(i+offsetRow,columnCounter,scoreQuestionsSerie[i,question-1])
        columnCounter+=1;
    
    #Sheet: Aantal keren onmogelijk per alternatief 
    sheetC = outputbook.add_sheet('OnmogelijkAantal_serie' + str(serie) )
    #column counter
    columnCounter = 0;
    
    
    #aantal keren onmogelijk per vraag en per antwoordalternatief
    counter=0
    #write alternative names on top
    columnCounter = 1
    for alternative in alternatives:
        sheetC.write(0,columnCounter,alternative,style=style_header)
        columnCounter+=1
        counter+=1
    counter=0
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #sheet2.write_merge(counter,counter,columnCounter,columnCounter+numAlternatives,"vraag" + str(question))
        #loop over alternatives
        sheetC.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                sheetC.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter],style_correctAnswer)
            else:
                # test of groep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                if (numOnmogelijkQuestionsAlternatives[counter] < numOnmogelijkQuestionsAlternatives[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                    sheetC.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter],style_specialAttention)
                else:
                    sheetC.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter])
            columnCounter+=1
            counter+=1

    
        
    #sheet: percentage onmogelijk per alternatief 
    sheetC = outputbook.add_sheet('OnmogelijkPerc_serie' + str(serie))
    #column counter
    columnCounter = 0;
    
    
    #percentage onmogelijk per vraag en per antwoordalternatief
    counter=0
    #write alternative names on top
    columnCounter = 1
    for alternative in alternatives:
        sheetC.write(0,columnCounter,alternative,style=style_header)
        columnCounter+=1
        counter+=1
    counter=0
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #loop over alternatives
        sheetC.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                sheetC.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter]/numParticipants,style_correctAnswer)
            else:
                 # test of groep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                if (numOnmogelijkQuestionsAlternatives[counter] < numOnmogelijkQuestionsAlternatives[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                    sheetC.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter]/numParticipants,style_specialAttention)
                else:
                    sheetC.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter]/numParticipants)
            columnCounter+=1
            counter+=1
    
    #sheet: Aantal keren mogelijk per alternatief 
    sheetC = outputbook.add_sheet('MogelijkAantal_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    #aantal keren mogelijk per vraag en per antwoordalternatief
    counter=0
    #write alternative names on top
    columnCounter = 1
    for alternative in alternatives:
        sheetC.write(0,columnCounter,alternative,style=style_header)
        columnCounter+=1
        counter+=1
    counter=0
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #sheet2.write_merge(counter,counter,columnCounter,columnCounter+numAlternatives,"vraag" + str(question))
        #loop over alternatives
        sheetC.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                sheetC.write(question,columnCounter,numMogelijkQuestionsAlternatives[counter],style_correctAnswer)
            else:
                 # test of groep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                if (numMogelijkQuestionsAlternatives[counter] > numMogelijkQuestionsAlternatives[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                    sheetC.write(question,columnCounter,numMogelijkQuestionsAlternatives[counter],style_specialAttention)
                else:
                    sheetC.write(question,columnCounter,numMogelijkQuestionsAlternatives[counter])
            columnCounter+=1
            counter+=1
            
    #sheet: Percentage mogelijk per alternatief 
    sheetC = outputbook.add_sheet('MogelijkPerc_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    #percentage mogelijk per vraag en per antwoordalternatief
    counter=0
    #write alternative names on top
    columnCounter = 1
    for alternative in alternatives:
        sheetC.write(0,columnCounter,alternative,style=style_header)
        columnCounter+=1
        counter+=1
    counter=0
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #loop over alternatives
        sheetC.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                sheetC.write(question,columnCounter,numMogelijkQuestionsAlternatives[counter]/numParticipants,style_correctAnswer)
            else:
                 # test of groep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                if (numMogelijkQuestionsAlternatives[counter] > numMogelijkQuestionsAlternatives[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                    sheetC.write(question,columnCounter,numMogelijkQuestionsAlternatives[counter]/numParticipants,style_specialAttention)
                else:
                    sheetC.write(question,columnCounter,numMogelijkQuestionsAlternatives[counter]/numParticipants)
            columnCounter+=1
            counter+=1
    
    #sheet: aantal onmogelijk per alternatief voor upper/middle/lower group
    sheetC = outputbook.add_sheet('OnMogelijkAantalUML_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    columnCounterAlternative = 1
    
    for alternative in alternatives:
        sheetC.write_merge(0,0,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(1,columnCounter,"upper",style=style_header)
        sheetC.write(1,columnCounter+1,"middle",style=style_header)
        sheetC.write(1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3
    
    offsetRow = 1
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #loop over alternatives
        sheetC.write(question+offsetRow,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                 # test of uppergroep het correcte antwoord meer onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper[counter] > numOnmogelijkQuestionsAlternativesLower[counter]):
                    sheetC.write(question+offsetRow,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter],style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(question+offsetRow,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper[counter],style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter],style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter],style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord minder onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper[counter] < numOnmogelijkQuestionsAlternativesLower[counter]):
                    sheetC.write(question+offsetRow,columnCounter,numOnmogelijkQuestionsAlternativesUpper[counter],style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter],style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter],style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                    if (numOnmogelijkQuestionsAlternativesUpper[counter] < numOnmogelijkQuestionsAlternativesUpper[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                        sheetC.write(question+offsetRow,columnCounter,numOnmogelijkQuestionsAlternativesUpper[counter],style_specialAttention)
                    else:
                        sheetC.write(question+offsetRow,columnCounter,numOnmogelijkQuestionsAlternativesUpper[counter])
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter])
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter])
            columnCounter+=3
            counter+=1
    
    
    
    #sheet: Percentage mogelijk per alternatief voor upper/middle/lower group
    sheetC = outputbook.add_sheet('OnmogelijkPercUML_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    columnCounterAlternative = 1
    
    for alternative in alternatives:
        sheetC.write_merge(0,0,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(1,columnCounter,"upper",style=style_header)
        sheetC.write(1,columnCounter+1,"middle",style=style_header)
        sheetC.write(1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3
    
    offsetRow = 1
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #loop over alternatives
        sheetC.write(question+offsetRow,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord meer onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper[counter]/numUpper > numOnmogelijkQuestionsAlternativesLower[counter]/numLower):
                    sheetC.write(question+offsetRow,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper[counter]/numUpper,style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter]/numMiddle,style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter]/numLower,style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(question+offsetRow,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper[counter]/numUpper,style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter]/numMiddle,style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter]/numLower,style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord minder onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper[counter]/numUpper < numOnmogelijkQuestionsAlternativesLower[counter]/numLower):
                    sheetC.write(question+offsetRow,columnCounter,numOnmogelijkQuestionsAlternativesUpper[counter]/numUpper,style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter]/numMiddle,style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter]/numLower,style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                    if (numOnmogelijkQuestionsAlternativesUpper[counter] < numOnmogelijkQuestionsAlternativesUpper[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                        sheetC.write(question+offsetRow,columnCounter,numOnmogelijkQuestionsAlternativesUpper[counter]/numUpper,style_specialAttention)
                    else:
                        sheetC.write(question+offsetRow,columnCounter,numOnmogelijkQuestionsAlternativesUpper[counter]/numUpper)
                    sheetC.write(question+offsetRow,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle[counter]/numMiddle)
                    sheetC.write(question+offsetRow,columnCounter+2,numOnmogelijkQuestionsAlternativesLower[counter]/numLower)
            columnCounter+=3
            counter+=1

    
    #sheet: aantal mogelijk per alternatief voor upper/middle/lower group
    sheetC = outputbook.add_sheet('MogelijkAantalUML_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    columnCounterAlternative = 1
    
    for alternative in alternatives:
        sheetC.write_merge(0,0,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(1,columnCounter,"upper",style=style_header)
        sheetC.write(1,columnCounter+1,"middle",style=style_header)
        sheetC.write(1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3
    
    offsetRow = 1
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #loop over alternatives
        sheetC.write(question+offsetRow,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord minder mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper[counter]< numMogelijkQuestionsAlternativesLower[counter]):
                    sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter],style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter],style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter],style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter],style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord meer mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper[counter] > numMogelijkQuestionsAlternativesLower[counter]):
                    sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter],style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter],style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter],style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                    if (numMogelijkQuestionsAlternativesUpper[counter] > numMogelijkQuestionsAlternativesUpper[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                        sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter],style_specialAttention)
                    else:
                        sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter])
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter])
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter])
            columnCounter+=3
            counter+=1
    
        
    
    #sheet: Percentage mogelijk per alternatief voor upper/middle/lower group
    sheetC = outputbook.add_sheet('MogelijkPercentageUML_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    columnCounterAlternative = 1
    
    for alternative in alternatives:
        sheetC.write_merge(0,0,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(1,columnCounter,"upper",style=style_header)
        sheetC.write(1,columnCounter+1,"middle",style=style_header)
        sheetC.write(1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3
    
    offsetRow = 1
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        correctAnswer = correctAnswers[serie-1][question-1]
        #loop over alternatives
        sheetC.write(question+offsetRow,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord minder mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper[counter]/numUpper < numMogelijkQuestionsAlternativesLower[counter]/numLower):
                    sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter]/numUpper,style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter]/numMiddle,style_correctAnswerSpecialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter]/numLower,style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter]/numUpper,style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter]/numMiddle,style_correctAnswer)
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter]/numLower,style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord meer mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper[counter]/numUpper > numMogelijkQuestionsAlternativesLower[counter]/numLower):
                    sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter]/numUpper,style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter]/numMiddle,style_specialAttention)
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter]/numLower,style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                    if (numMogelijkQuestionsAlternativesUpper[counter] > numMogelijkQuestionsAlternativesUpper[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                        sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter]/numUpper,style_specialAttention)
                    else:
                        sheetC.write(question+offsetRow,columnCounter,numMogelijkQuestionsAlternativesUpper[counter]/numUpper)
                    sheetC.write(question+offsetRow,columnCounter+1,numMogelijkQuestionsAlternativesMiddle[counter]/numMiddle)
                    sheetC.write(question+offsetRow,columnCounter+2,numMogelijkQuestionsAlternativesLower[counter]/numLower)
            columnCounter+=3
            counter+=1
            
            
    #sheet: gemiddelde score per vraag
    sheetC = outputbook.add_sheet('GemScorePerVraag_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    #gemiddelde score per vraag
    counter=0
    #write all/upper/middle/lower on top
    columnCounter = 1
    sheetC.write(0,columnCounter,"all",style=style_header)
    columnCounter+=1
    sheetC.write(0,columnCounter,"upper",style=style_header)
    columnCounter+=1
    sheetC.write(0,columnCounter,"middle",style=style_header)
    columnCounter+=1
    sheetC.write(0,columnCounter,"lower",style=style_header)
    columnCounter+=1
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        sheetC.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1        
        if averageScoreQuestions[question-1]<0:
            sheetC.write(question,columnCounter,averageScoreQuestions[question-1],style=style_specialAttention)        
        else:
            sheetC.write(question,columnCounter,averageScoreQuestions[question-1])                
        columnCounter+=1 
        if averageScoreQuestionsUpper[question-1]<=averageScoreQuestionsLower[question-1] or averageScoreQuestionsUpper[question-1]<=averageScoreQuestionsMiddle[question-1]:
            sheetC.write(question,columnCounter,averageScoreQuestionsUpper[question-1],style=style_specialAttention)
            columnCounter+=1
            sheetC.write(question,columnCounter,averageScoreQuestionsMiddle[question-1],style=style_specialAttention)
            columnCounter+=1
            sheetC.write(question,columnCounter,averageScoreQuestionsLower[question-1],style=style_specialAttention)    
        else:
            sheetC.write(question,columnCounter,averageScoreQuestionsUpper[question-1])
            columnCounter+=1
            sheetC.write(question,columnCounter,averageScoreQuestionsMiddle[question-1])
            columnCounter+=1
            sheetC.write(question,columnCounter,averageScoreQuestionsLower[question-1])
    columnCounter=0
    sheetC.write(numQuestions+1,columnCounter,"total",style=style_header_borderRight)
    columnCounter+=1
    sheetC.write(numQuestions+1,columnCounter,averageScoreSerie)   
    columnCounter+=1
    sheetC.write(numQuestions+1,columnCounter,averageScoreUpper)
    columnCounter+=1
    sheetC.write(numQuestions+1,columnCounter,averageScoreMiddle)
    columnCounter+=1
    sheetC.write(numQuestions+1,columnCounter,averageScoreLower)  
    columnCounter+=1

    
    
    ##sheet: histogram per question
    sheetC = outputbook.add_sheet('HistScorePerVraag_serie'+str(serie))
    #column counter
    columnCounter = 0;
    
    
    #gemiddelde verdeling scores per vraag
    counter=0
    #write different scores
    
    possibleScores=numpy.arange(-1.0,1+1.0/2.0,1.0/3.0)
    #print possibleScores
    
    columnCounter = 1
    for possibleScore in possibleScores[0:len(possibleScores)-1]:
        sheetC.write(0,columnCounter,possibleScore,style=style_header)
        columnCounter+=1
    sheetC.write(0,columnCounter,"gemiddelde",style=style_header)
        
    for question in xrange(1,numQuestions+1):
        columnCounter=0
        sheetC.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        hist,bins = numpy.histogram(scoreQuestions[:,question-1],bins=possibleScores-1.0/6.0)
        columnCounter+=1    
        for n in hist:        
            if (hist[0]>hist[len(hist)-1] or hist[0]+hist[1]>hist[len(hist)-1]+hist[len(hist)-2]): #more confident in wrong answer than confident in correct answer
                sheetC.write(question,columnCounter,n,style=style_specialAttention)
            else:
                sheetC.write(question,columnCounter,n)
            columnCounter+=1
        if averageScoreQuestions[question-1]<0:
            sheetC.write(question,columnCounter,averageScoreQuestions[question-1],style=style_specialAttention)        
        else:
            sheetC.write(question,columnCounter,averageScoreQuestions[question-1]) 
    

outputbook.save('output.xls') 