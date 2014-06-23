# -*- coding: utf-8 -*-
"""
Created on Mon Mar 31 13:58:48 2014

@author: tdelaet

Dit neemt aan dat de gebruikte sheet van excel file de volgende kolommen heeft (met eerste rij de naam van de kolom):
- studentennummer
- vragenreeksper
- Vraag1A, Vraag1B, ... (komt overeen met aantal alternatieven - numAlternatives)
 en dit voor alle vragen (komt overeen met numQuestions)
"""

from xlrd import open_workbook
import string
import numpy
import matplotlib.pyplot as plt
from xlwt import Workbook

import checkInputVariables
import supportFunctions
import writeResults


nameFile = "../OMR/review1_corr.xlsx" #name of excel file with scanned forms
nameSheet = "output" #sheet name of excel file with scanned forms
numQuestions = 25 # number of questions
numAlternatives = 4 #number of alternatives
maxTotalScore = 20 #maximum total score
numSeries=4 # number of series

correctAnswers = numpy.loadtxt("../sleutel.txt",delimiter=',',dtype=numpy.str)

#correctAnswers = ["B","B","C","B","A","A","B","D","A","C","B","B","D","B","A","A","B","D","A","C","D","D","C","C","A"] #correct answers for the first series

permutations = numpy.loadtxt("../permutatie.txt",delimiter=',',dtype=numpy.int32)
#permutations = [[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25],
#                [2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,1],
#                [3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,1,2],
#                [4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,1,2,3]] #permutations of the different series

weightsQuestions = numpy.loadtxt("../gewichten.txt",delimiter=',',dtype=numpy.int32)

twoOptions=["onmogelijk","mogelijk"] #elimination options should be first

plt.close("all")

#letters of answer alternatives
alternatives = list(string.ascii_uppercase)[0:numAlternatives]

############################
#create list of expected content of scan file
content = ["studentennummer","vragenreeks"]
for question in xrange(1,numQuestions+1):
    for alternative in alternatives:
        name = "Vraag" + str(question) + alternative
        content.append(name)
#print content
###########################
        
if not( checkInputVariables.checkInputVariables(nameFile,nameSheet,numQuestions,numAlternatives,numSeries,correctAnswers,permutations,weightsQuestions,twoOptions)):
     print "ERROR found in input variables"   

        
  
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

content_colNrs = supportFunctions.giveContentColNrs(content, sheet);

            
#print columnSeries
scoreQuestionsIndicatedSeries= numpy.zeros((numParticipants,numQuestions))


# write to excel_file
outputbook = Workbook(style_compression=2)
outputStudentbook = Workbook(style_compression=2)


name = "studentennummer"
studentenNrCol= content_colNrs[content.index(name)]
deelnemers=sheet.col_values(studentenNrCol,1,num_rows)

if not supportFunctions.checkForUniqueParticipants(deelnemers):
    print "ERROR: Duplicate participants found"


name = "vragenreeks"
#get the column in which the vragenreeks is stored
colNrSerie = content_colNrs[content.index(name)]
#get the series for the participants (so skip for row with name of first row)
columnSeries=sheet.col_values(colNrSerie,1,num_rows)


#get the score for all permutations for each of the questions
scoreQuestionsAllPermutations= supportFunctions.calculateScoreAllPermutations(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs,twoOptions)     

numOnmogelijkQuestionsAlternatives, numMogelijkQuestionsAlternatives = supportFunctions.getNumberMogelijkOnmogelijk(sheet,content,permutations,columnSeries,scoreQuestionsIndicatedSeries,alternatives,twoOptions,content_colNrs)
#print scoreQuestionsAllPermutations

matrixAnswers = supportFunctions.getMatrixAnswers(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs,twoOptions)     


#get the scores for the indicated series
scoreQuestionsIndicatedSeries, averageScoreQuestions =  supportFunctions.getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations,columnSeries)

#get the overall statistics
totalScore, averageScore, medianScore, standardDeviation, percentagePass = supportFunctions.getOverallStatistics(scoreQuestionsIndicatedSeries,maxTotalScore,weightsQuestions)
#print totalScore
#print averageScore
#print medianScore
#print percentagePass

totalScoreDifferentPermutations = supportFunctions.calculateTotalScoreDifferentPermutations(scoreQuestionsAllPermutations,maxTotalScore,weightsQuestions)
#print totalScoreDifferentPermutations

numParticipantsSeries, averageScoreSeries, medianScoreSeries, standardDeviationSeries, percentagePassSeries, averageScoreQuestionsDifferentSeries = supportFunctions.getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations,scoreQuestionsIndicatedSeries,columnSeries,maxTotalScore)

totalScoreUpper,totalScoreMiddle,totalScoreLower,averageScoreUpper, averageScoreMiddle, averageScoreLower, averageScoreQuestionsUpper, averageScoreQuestionsMiddle, averageScoreQuestionsLower, numOnmogelijkQuestionsAlternativesUpper, numOnmogelijkQuestionsAlternativesMiddle, numOnmogelijkQuestionsAlternativesLower, numMogelijkQuestionsAlternativesUpper, numMogelijkQuestionsAlternativesMiddle, numMogelijkQuestionsAlternativesLower, scoreQuestionsUpper, scoreQuestionsMiddle, scoreQuestionsLower, numUpper, numMiddle, numLower = supportFunctions.calculateUpperLowerStatistics(sheet,content,columnSeries,totalScore,scoreQuestionsIndicatedSeries,correctAnswers,alternatives,twoOptions,content_colNrs,permutations)
#totalScoreUpper,totalScoreMiddle,totalScoreLower,averageScoreUpper, averageScoreMiddle, averageScoreLower, averageScoreQuestionsUpper, averageScoreQuestionsMiddle, averageScoreQuestionsLower,numQuestionsAlternativesUpper,numQuestionsAlternativesMiddle,numQuestionsAlternativesLower, scoreQuestionsUpper, scoreQuestionsMiddle, scoreQuestionsLower,numUpper, numMiddle, numLower= supportFunctions.calculateUpperLowerStatistics(sheet,content,columnSeries,totalScore,scoreQuestionsIndicatedSeries,correctAnswers,alternatives,blankAnswer,content_colNrs,permutations)
 


## WRITING THE OUTPUT TO A FILE
writeResults.write_results(outputbook,weightsQuestions,numQuestions,correctAnswers,alternatives,maxTotalScore,content,content_colNrs,
                  columnSeries,deelnemers,
                  numParticipants,
                  totalScore,percentagePass,
                  scoreQuestionsIndicatedSeries,
                  totalScoreDifferentPermutations,
                  medianScore,
                  standardDeviation,
                  averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,
                  averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower,
                  averageScoreQuestionsDifferentSeries,
                  numUpper,numMiddle,numLower,
                  numParticipantsSeries,
                  averageScoreSeries,medianScoreSeries,standardDeviationSeries,percentagePassSeries,                  
                  numOnmogelijkQuestionsAlternatives, numMogelijkQuestionsAlternatives,
                  numOnmogelijkQuestionsAlternativesUpper,numOnmogelijkQuestionsAlternativesMiddle,numOnmogelijkQuestionsAlternativesLower,                  
                  numMogelijkQuestionsAlternativesUpper,numMogelijkQuestionsAlternativesMiddle,numMogelijkQuestionsAlternativesLower
                  )
                  
## WRITING A FILE TO UPLOAD TO TOLEDO WITH THE GRADES
writeResults.write_scoreStudents(outputStudentbook,"punten",permutations,weightsQuestions,numParticipants,deelnemers, numQuestions,numAlternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)           
                  

# plot the histogram of the total score
plt.figure()
n, bins, patches = plt.hist(totalScore,bins=numpy.arange(0-0.5,maxTotalScore+1,1))
plt.title("histogram total score")
plt.xlabel("score (max " + str(maxTotalScore)+ ")")
plt.xlim([0-0.5,maxTotalScore+0.5])
plt.xticks(numpy.arange(0,maxTotalScore+1,1))
plt.ylabel("number of students")
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()
plt.savefig('histogramGeheel.png', bbox_inches='tight')

#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) 
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

binsHist = numpy.arange(-1-1.0/6.0,1+1.0/3.0,1.0/3.0)

for question in xrange(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist(scoreQuestionsIndicatedSeries[:,question-1],bins=binsHist)
    plt.xticks([-1/(numAlternatives-1), 0,1])
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig('histogramVragen.png', bbox_inches='tight')

#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) 
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

for question in xrange(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist([scoreQuestionsUpper[:,question-1], scoreQuestionsMiddle[:,question-1], scoreQuestionsLower[:,question-1]],bins=binsHist, stacked=True,  label=['Upper', 'Middle', 'Lower'],color=['g','b','r'])
    plt.xticks([-1/(numAlternatives-1), 0,1])
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")
    plt.legend(loc=2,prop={'size':6})
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig('histogramVragenUML.png', bbox_inches='tight')


outputbook.save('output.xls') 
outputStudentbook.save('punten.xls')