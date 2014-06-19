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


nameFile = "OMRoutputEN_sessie2.xlsx"
nameSheet = "outputScan"
numQuestions = 10
numAlternatives = 4
maxTotalScore = 20

#correctAnswers = ["B","B","D","B","A","A","B","D","A","C"]#sessie 1
correctAnswers = ["B","D","B","A","A","B","D","A","C","B"]#sessie 2
twoOptions=["onmogelijk","mogelijk"]

plt.close("all")





##TODO: behaalde score in andere reeksen en totale score

#letters of answer alternatives
alternatives = list(string.ascii_uppercase)[0:numAlternatives]
print alternatives

if numQuestions != len(correctAnswers):
    print "The length of the list of correct answers is not equal to the number of questions"
    #sys.exit()
for answer in correctAnswers:
    print answer
    if not(answer in alternatives):
        print "The answer " + answer +  " is not in the list of possible answer alternatives"
        #sys.exit()


#create list of expected content
content = ["Deelnemersnummer","Deelnemersnummer_bar","vragenreeks"]
for question in xrange(1,numQuestions+1):
    for alternative in alternatives:
        name = "vraag" + str(question) + alternative
        content.append(name)
#print content


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
scoreQuestions= numpy.zeros((numParticipants,numQuestions))
numOnmogelijkQuestionsAlternatives= numpy.zeros(numQuestions*numAlternatives)
numMogelijkQuestionsAlternatives= numpy.zeros(numQuestions*numAlternatives)


#loop over question
counter = 0;
for question in xrange(1,numQuestions+1):
    #print "vraag" + str(question)
    correctAnswer = correctAnswers[question-1]
    #print "correct answer is " + correctAnswer
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
        if  correctAnswer == alternative:         # if the alternative is the correct answer => wrongly excluded so -1
            scoreQuestions[indicesOnmogelijk,question-1]-=1.0
        else: # if the alternative is NOT the correct answer => correctly excluded so +1/(numAlternatives-1)
            scoreQuestions[indicesOnmogelijk,question-1]=scoreQuestions[indicesOnmogelijk,question-1] + 1.0/(float(numAlternatives)-1.0)

#print scoreQuestions
totalScore = scoreQuestions.sum(axis=1)/numQuestions*maxTotalScore
# set negative scores to 0
totalScore[totalScore < 0]=0
totalScore = numpy.round(totalScore)
averageScore = sum(totalScore)/float(numParticipants)
medianScore = numpy.median(totalScore)
#print totalScore

#averageScoreQuestions
averageScoreQuestions = scoreQuestions.sum(axis=0)/float(numParticipants)
##TODO:STD

#find upper/middle/lower third of the students
#sort students according to score
orderedDeelnemers = sorted(range(len(totalScore)),key=totalScore.__getitem__) 
third = int(numpy.ceil(numParticipants/3))
indicesUpper = orderedDeelnemers[numParticipants-third:numParticipants]
indicesLower = orderedDeelnemers[0:third]
indicesMiddle= orderedDeelnemers[third:numParticipants-third]
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
totalScoreUpper = totalScore[indicesUpper]
totalScoreMiddle = totalScore[indicesMiddle]
totalScoreLower = totalScore[indicesLower]
averageScoreUpper = sum(totalScoreUpper)/float(numUpper)
averageScoreMiddle = sum(totalScoreMiddle)/float(numMiddle)
averageScoreLower = sum(totalScoreLower)/float(numLower)
scoreQuestionsUpper =  scoreQuestions[indicesUpper,:]
scoreQuestionsMiddle =  scoreQuestions[indicesMiddle,:]
scoreQuestionsLower =  scoreQuestions[indicesLower,:]
averageScoreQuestionsUpper = scoreQuestionsUpper.sum(axis=0)/float(numUpper)
averageScoreQuestionsMiddle = scoreQuestionsMiddle.sum(axis=0)/float(numMiddle)
averageScoreQuestionsLower = scoreQuestionsLower.sum(axis=0)/float(numLower)

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
    #print "vraag" + str(question)
    correctAnswer = correctAnswers[question-1]
    #print "correct answer is " + correctAnswer
    #loop over alternatives
    for alternative in alternatives:
        name = "vraag" + str(question) + alternative
        #print name
        #find column number in which question alternatives are given
        colNr = content_colNrs[content.index(name)]
        #get the answers for the participants (so skip for row with name of first row)
        columnAlternative=sheet.col_values(colNr,1,num_rows)
        #get the indices of the particpants that answered onmogelijk
        indicesOnmogelijkUpper = [x for x in indicesUpper if columnAlternative[x]=="onmogelijk"]
        indicesOnmogelijkMiddle = [x for x in indicesMiddle if columnAlternative[x]=="onmogelijk"]
        indicesOnmogelijkLower = [x for x in indicesLower if columnAlternative[x]=="onmogelijk"]
        #print len(indicesOnmogelijkUpper)
        #print len(indicesOnmogelijkMiddle)
        #print len(indicesOnmogelijkLower)
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


#number of mogelijk in upper group per question
#print len(indicesUpper)
#print len(indicesLower)
#print len(indicesMiddle)



#plot statistics for each question
# plot the histogram of the total score
plt.figure()
n, bins, patches = plt.hist(totalScore,bins=numpy.arange(0-0.5,maxTotalScore+1,1))
plt.title("histogram total score")
plt.xlabel("score (max " + str(maxTotalScore)+ ")")
plt.xlim([0-0.5,maxTotalScore+0.5])
plt.ylabel("number of students")
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()
plt.savefig('histogramGeheel.png', bbox_inches='tight')

# plot the quintiles of the total score
plt.figure()
n, bins, patches = plt.hist(totalScore,bins=numpy.arange(0-0.5,maxTotalScore+1,maxTotalScore/5.0))
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
    n, bins, patches = plt.hist(scoreQuestions[:,question-1],bins=numpy.arange(-1-1.0/6.0,1+1.0/3.0,1.0/3.0))
    plt.xticks(numpy.arange(-1, 1+1.0/3.0, 2*1.0/3.0))
    plt.title("vraag " + str(question))
    plt.xlabel("score (max " + str(maxTotalScore)+ ")")
    plt.xlim([-1-1.0/3.0,1+1.0/3.0])
    plt.ylabel("number of students")
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig('histogramVragen.png', bbox_inches='tight')



# write to excel_file
outputbook = Workbook()

style_header = easyxf("font: bold on; align: horiz center; border: bottom medium")
style_header_borderRight = easyxf("font: bold on; align: horiz center; border: right medium ")
style_correctAnswer = easyxf('font: color blue')
style_specialAttention  = easyxf('pattern: pattern solid, fore_colour red')
style_correctAnswerSpecialAttention = easyxf('font: color blue;''pattern: pattern solid, fore_colour red')

style_border_header = easyxf('border: left thick, top thick, bottom thick, right thick')


#First sheet: Score per Deelnemer
sheet1 = outputbook.add_sheet('ScorePerDeelnemer')

#column counter
columnCounter = 0;

#deelnemersnummers
deelnemersBarColNr= content_colNrs[content.index("Deelnemersnummer_bar")]
deelnemers=sheet.col_values(deelnemersBarColNr,columnCounter,num_rows)
#print deelnemers
sheet1.write(0, 0,"deelnemersnummer", style=style_header_borderRight) 
for i in xrange(1,len(deelnemers)):
    sheet1.write(i,columnCounter,deelnemers[i], style=style_header_borderRight)
columnCounter+=1;



#totale score
sheet1.write(0,columnCounter,"totale score",style=style_header)
offsetRow=1
for i in xrange(len(totalScore)):
    sheet1.write(i+offsetRow,columnCounter,totalScore[i])
columnCounter+=1;

#score per vraag
for question in xrange(1,numQuestions+1):
    sheet1.write(0,columnCounter,"vraag " + str(question),style=style_header)
    offsetRow=1
    for i in xrange(len(scoreQuestions)):
        sheet1.write(i+offsetRow,columnCounter,scoreQuestions[i,question-1])
    columnCounter+=1;

#Second sheet: Aantal keren onmogelijk per alternatief 
sheet = outputbook.add_sheet('StatistiekOnmogelijkAantal')
#column counter
columnCounter = 0;



#aantal keren onmogelijk per vraag en per antwoordalternatief
counter=0
#write alternative names on top
columnCounter = 1
for alternative in alternatives:
    sheet.write(0,columnCounter,alternative,style=style_header)
    columnCounter+=1
    counter+=1
counter=0
    
for question in xrange(1,numQuestions+1):
    columnCounter=0
    correctAnswer = correctAnswers[question-1]
    #sheet2.write_merge(counter,counter,columnCounter,columnCounter+numAlternatives,"vraag" + str(question))
    #loop over alternatives
    sheet.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
    columnCounter+=1
    for alternative in alternatives:
        if alternative == correctAnswer:
            sheet.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter],style_correctAnswer)
        else:
            # test of groep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
            if (numOnmogelijkQuestionsAlternatives[counter] < numOnmogelijkQuestionsAlternatives[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                sheet.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter],style_specialAttention)
            else:
                sheet.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter])
        columnCounter+=1
        counter+=1


    
#Third sheet: percentage onmogelijk per alternatief 
sheet = outputbook.add_sheet('StatistiekOnmogelijkPercentage')
#column counter
columnCounter = 0;


#percentage onmogelijk per vraag en per antwoordalternatief
counter=0
#write alternative names on top
columnCounter = 1
for alternative in alternatives:
    sheet.write(0,columnCounter,alternative,style=style_header)
    columnCounter+=1
    counter+=1
counter=0
    
for question in xrange(1,numQuestions+1):
    columnCounter=0
    correctAnswer = correctAnswers[question-1]
    #loop over alternatives
    sheet.write(question,columnCounter,"vraag"+str(question),style=style_header_borderRight)
    columnCounter+=1
    for alternative in alternatives:
        if alternative == correctAnswer:
            sheet.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter]/numParticipants,style_correctAnswer)
        else:
             # test of groep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
            if (numOnmogelijkQuestionsAlternatives[counter] < numOnmogelijkQuestionsAlternatives[(question-1)*len(alternatives)+alternatives.index(correctAnswer)] ):
                sheet.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter]/numParticipants,style_specialAttention)
            else:
                sheet.write(question,columnCounter,numOnmogelijkQuestionsAlternatives[counter]/numParticipants)
        columnCounter+=1
        counter+=1

#Fourth sheet: Aantal keren mogelijk per alternatief 
sheetC = outputbook.add_sheet('StatistiekMogelijkAantal')
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
    correctAnswer = correctAnswers[question-1]
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
        
#Fifth sheet: Percentage mogelijk per alternatief 
sheetC = outputbook.add_sheet('StatistiekMogelijkPercentage')
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
    correctAnswer = correctAnswers[question-1]
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

#Sixth sheet: aantal onmogelijk per alternatief voor upper/middle/lower group
sheetC = outputbook.add_sheet('StatistiekOnMogelijkAantalUML')
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
    correctAnswer = correctAnswers[question-1]
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



#Seventh sheet: Percentage mogelijk per alternatief voor upper/middle/lower group
sheetC = outputbook.add_sheet('StatistiekOnmogelijkPercUML')
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
    correctAnswer = correctAnswers[question-1]
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


#Eight sheet: aantal mogelijk per alternatief voor upper/middle/lower group
sheetC = outputbook.add_sheet('StatistiekMogelijkAantalUML')
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
    correctAnswer = correctAnswers[question-1]
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



#Ninth sheet: Percentage mogelijk per alternatief voor upper/middle/lower group
sheetC = outputbook.add_sheet('StatistiekMogelijkPercentageUML')
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
    correctAnswer = correctAnswers[question-1]
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
        
        
#Tenth sheet: gemiddelde score per vraag
sheetC = outputbook.add_sheet('GemScorePerVraag')
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
sheetC.write(numQuestions+1,columnCounter,averageScore)   
columnCounter+=1
sheetC.write(numQuestions+1,columnCounter,averageScoreUpper)
columnCounter+=1
sheetC.write(numQuestions+1,columnCounter,averageScoreMiddle)
columnCounter+=1
sheetC.write(numQuestions+1,columnCounter,averageScoreLower)  
columnCounter+=1



##11th sheet
sheetC = outputbook.add_sheet('HistScorePerVraag')
#column counter
columnCounter = 0;


#gemiddelde verdeling scores per vraag
counter=0
#write different scores

possibleScores=numpy.arange(-1.0,1+1.0/2.0,1.0/3.0)
print possibleScores

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


##12th sheet
sheetC = outputbook.add_sheet('globaleParameters')
#column counter
columnCounter = 0;


#gemiddelde verdeling scores per vraag

rowCounter = 0 

columnCounter = 0
sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=style_header)
columnCounter+=1  
sheetC.write(rowCounter,columnCounter,numParticipants)

rowCounter+=1
columnCounter = 0
sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=style_header)
columnCounter+=1  
sheetC.write(rowCounter,columnCounter,averageScore)

rowCounter+=1
columnCounter = 0
sheetC.write(rowCounter,columnCounter,"mediaan ",style=style_header)
columnCounter+=1  
sheetC.write(rowCounter,columnCounter,medianScore)

rowCounter+=1
columnCounter = 0
sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=style_header)
columnCounter+=1  
aantalGeslaagd = sum(score>= maxTotalScore/2.0 for score in totalScore)
sheetC.write(rowCounter,columnCounter,float(aantalGeslaagd)/float(numParticipants))


outputbook.save('output.xls') 