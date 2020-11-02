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

#Create directories
import os
import checkInputVariables
import supportFunctions_reworked
import writeResults_reworked
import GUI_reworked


print ("Use the Graphical User Interface to continue")
nameTest = GUI_reworked.InputGUI()
#nameTest="TTT"

nameFile = "../"+ nameTest + "/OMR/OMRoutput.xlsx" #name of excel file with scanned forms
nameSheet = "outputScan" #sheet name of excel file with scanned forms

############################

output = GUI_reworked.VragenGUI()
numQuestions = output['questions']
numAlternatives = output['alternatives'][1]
maxTotalScore = output['totalscore']
numSeries= output['permutations'] # number of series
twoOptions=["onmogelijk","mogelijk"] #elimination options should be first

#numQuestions = 14
#numAlternatives = 4
#maxTotalScore = 10
#numSeries= 1 # number of series
#twoOptions=["onmogelijk","mogelijk"] #elimination options should be first



############################

correctAnswers = numpy.loadtxt("../"+ nameTest + "/sleutel.txt",delimiter=',',dtype=numpy.str)
permutations = numpy.loadtxt("../"+ nameTest + "/permutatie.txt",delimiter=',',dtype=numpy.int32)
weightsQuestions = numpy.loadtxt("../"+ nameTest + "/gewichten.txt",delimiter=',',dtype=numpy.int32)
badQuestions = numpy.loadtxt("../"+ nameTest + "/slechteVragen.txt",delimiter=',',dtype=numpy.int32)

if not os.path.exists("../"+ nameTest + "/Results/Sleutels"):
    os.makedirs("../"+ nameTest + "/Results/Sleutels")
if not os.path.exists("../"+ nameTest + "/Results/Histogram"):
    os.makedirs("../"+ nameTest + "/Results/Histogram")
    
correctAnswersDifferentPermutations = []
# maar 1 reeks
if numSeries ==1:
    numpy.savetxt("../"+ nameTest + "/Results/Sleutels" + "/sleutel"+ "_reeks1" +".txt",  [correctAnswers[x-1] for x in permutations], delimiter=" ", fmt="%s")
else:
    for i in range(len(permutations)):
        numpy.savetxt("../"+ nameTest + "/Results/Sleutels" + "/sleutel"+ "_reeks" + str(i+1)+".txt",  [correctAnswers[x-1] for x in permutations[i]], delimiter=" ", fmt="%s")

plt.close("all")

#letters of answer alternatives
alternatives = list(string.ascii_uppercase)[0:numAlternatives]

############################
#create list of expected content of scan file
content = ["studentennummer","vragenreeks"]
for question in range(1,numQuestions+1):
    for alternative in alternatives:
        name = "Vraag" + str(question) + alternative
        content.append(name)
#print content
###########################
        
if not( checkInputVariables.checkInputVariables(nameFile,nameSheet,numQuestions,numAlternatives,numSeries,correctAnswers,permutations,weightsQuestions,badQuestions, twoOptions)):
     print ("ERROR found in input variables"     )
  
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

content_colNrs = supportFunctions_reworked.giveContentColNrs(content, sheet);

scoreQuestionsIndicatedSeries= numpy.zeros((numParticipants,numQuestions))


# write to excel_file
outputbook = Workbook(style_compression=2)
outputStudentbook = Workbook(style_compression=2)


name = "studentennummer"
studentenNrCol= content_colNrs[content.index(name)]
deelnemers=sheet.col_values(studentenNrCol,1,num_rows)

# check for double participants
if not supportFunctions_reworked.checkForUniqueParticipants(deelnemers):
    print ("ERROR: Duplicate participants found")


name = "vragenreeks"
#get the column in which the vragenreeks is stored
colNrSerie = content_colNrs[content.index(name)]
#get the series for the participants (so skip for row with name of first row)
columnSeries=sheet.col_values(colNrSerie,1,num_rows)


#get the score for all permutations for each of the questions
scoreQuestionsAllPermutations= supportFunctions_reworked.calculateScoreAllPermutations(sheet,content,numSeries,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs,twoOptions,badQuestions)     

numOnmogelijkQuestionsAlternatives, numMogelijkQuestionsAlternatives = supportFunctions_reworked.getNumberMogelijkOnmogelijk(sheet,content,numSeries,permutations,columnSeries,scoreQuestionsIndicatedSeries,alternatives,twoOptions,content_colNrs)
#print scoreQuestionsAllPermutations

matrixAnswers = supportFunctions_reworked.getMatrixAnswers(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs,twoOptions)     


#get the scores for the indicated series
scoreQuestionsIndicatedSeries, averageScoreQuestions =  supportFunctions_reworked.getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations,columnSeries)

#get the overall statistics
totalScore, totalScore_nonRounded, averageScore, medianScore, standardDeviation, percentagePass = supportFunctions_reworked.getOverallStatistics(scoreQuestionsIndicatedSeries,maxTotalScore,weightsQuestions)
#print totalScore
#print averageScore
#print medianScore
#print percentagePass

totalScoreDifferentPermutations = supportFunctions_reworked.calculateTotalScoreDifferentPermutations(scoreQuestionsAllPermutations,maxTotalScore,weightsQuestions)
#print totalScoreDifferentPermutations

numParticipantsSeries, averageScoreSeries, medianScoreSeries, standardDeviationSeries, percentagePassSeries, averageScoreQuestionsDifferentSeries = supportFunctions_reworked.getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations,scoreQuestionsIndicatedSeries,columnSeries,maxTotalScore)

totalScoreUpper,totalScoreMiddle,totalScoreLower,averageScoreUpper, averageScoreMiddle, averageScoreLower, averageScoreQuestionsUpper, averageScoreQuestionsMiddle, averageScoreQuestionsLower, numOnmogelijkQuestionsAlternativesUpper, numOnmogelijkQuestionsAlternativesMiddle, numOnmogelijkQuestionsAlternativesLower, numMogelijkQuestionsAlternativesUpper, numMogelijkQuestionsAlternativesMiddle, numMogelijkQuestionsAlternativesLower, scoreQuestionsUpper, scoreQuestionsMiddle, scoreQuestionsLower, numUpper, numMiddle, numLower = supportFunctions_reworked.calculateUpperLowerStatistics(sheet,content,numSeries,columnSeries,totalScore,scoreQuestionsIndicatedSeries,correctAnswers,alternatives,twoOptions,content_colNrs,permutations)
#totalScoreUpper,totalScoreMiddle,totalScoreLower,averageScoreUpper, averageScoreMiddle, averageScoreLower, averageScoreQuestionsUpper, averageScoreQuestionsMiddle, averageScoreQuestionsLower,numQuestionsAlternativesUpper,numQuestionsAlternativesMiddle,numQuestionsAlternativesLower, scoreQuestionsUpper, scoreQuestionsMiddle, scoreQuestionsLower,numUpper, numMiddle, numLower= supportFunctions.calculateUpperLowerStatistics(sheet,content,columnSeries,totalScore,scoreQuestionsIndicatedSeries,correctAnswers,alternatives,blankAnswer,content_colNrs,permutations)
 
totalVariance, Variance = supportFunctions_reworked.calculateVariances(totalScore,scoreQuestionsIndicatedSeries,numQuestions)

itemToetsCorrelatie = supportFunctions_reworked.calculateItemToetsCorrelatie(totalScore,scoreQuestionsIndicatedSeries,numQuestions)
## WRITING THE OUTPUT TO A FILE
writeResults_reworked.write_results(outputbook,weightsQuestions,numQuestions,correctAnswers,alternatives,maxTotalScore,content,content_colNrs,
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
writeResults_reworked.write_scoreStudents(outputStudentbook,"punten",numSeries,permutations,weightsQuestions,numParticipants,deelnemers, numQuestions,numAlternatives,content,content_colNrs,totalScore,totalScore_nonRounded,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)          

writeResults_reworked.write_scoreStudentsNonPermutated(outputStudentbook,"verwerking",numSeries,permutations,weightsQuestions,numParticipants,deelnemers, numQuestions,numAlternatives,alternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)           
                  
writeResults_reworked.write_CronbachsAlpha(outputbook,"Cronbach's alpha", numQuestions, totalVariance, Variance)

writeResults_reworked.write_itemToetsCorrelatie(outputbook,"Item-toets correlatie", numQuestions, itemToetsCorrelatie)
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
plt.savefig("../"+ nameTest + "/Results/Histogram" + "/histogramGeheel.png", bbox_inches='tight')

#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))+1
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) +1
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

scoreWrongAnswer_ET_rummens = -1.0/(numAlternatives-1)
# scores for y correctly eliminated answers
scoreCorrectAnswer_ET_rummensyEL = [(1.0*y)/((numAlternatives-y)*(numAlternatives-1)) for y in numpy.arange(0,numAlternatives)]
possibleScores=numpy.concatenate(([scoreWrongAnswer_ET_rummens],scoreCorrectAnswer_ET_rummensyEL))
catLow = possibleScores - 1.0/10.0
catHigh = possibleScores + 1.0/10.0
binsHist = numpy.sort(numpy.concatenate((catLow,catHigh)))


for question in range(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist(scoreQuestionsIndicatedSeries[:,question-1],bins=binsHist)
    plt.xticks(possibleScores)
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig("../"+ nameTest + "/Results/Histogram" + "/histogramVragen.png", bbox_inches='tight')

#plot histogram for different questions
numColsPict = int(numpy.ceil(numpy.sqrt(numQuestions)))+1
#print numColsPict
numRowsPict = int(numpy.ceil(numQuestions/numColsPict)) +1
#print numRowsPict
fig, axes = plt.subplots(nrows=numRowsPict, ncols=numColsPict)
fig.tight_layout() # Or equivalently,  "plt.tight_layout()"

for question in range(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist([scoreQuestionsUpper[:,question-1], scoreQuestionsMiddle[:,question-1], scoreQuestionsLower[:,question-1]],bins=binsHist, stacked=True,  label=['Upper', 'Middle', 'Lower'],color=['g','b','r'])
    plt.xticks(possibleScores)
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives-1),1+1.0/(numAlternatives-1)])
    plt.ylabel("aantal studenten")
    plt.legend(loc=2,prop={'size':6})
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig("../"+ nameTest + "/Results/Histogram" + "/histogramVragenUML.png", bbox_inches='tight')


outputbook.save("../"+ nameTest + "/Results/output.xls") 
outputStudentbook.save("../"+ nameTest + "/Results/punten.xls")
