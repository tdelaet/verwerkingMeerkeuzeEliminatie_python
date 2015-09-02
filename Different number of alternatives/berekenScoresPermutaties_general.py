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
import sys
import checkInputVariables_general
import supportFunctions_general
import writeResults_general
import GUI_general

print "START: Use the Graphical User Interface to continue and input the data of your test"
nameTest = GUI_general.InputGUI()

nameFile = "../"+ nameTest + "/OMR/OMRoutput.xlsx" #name of excel file with scanned forms
nameSheet = "outputScan" #sheet name of excel file with scanned forms

############################

output = GUI_general.VragenGUI()

numQuestions = output['questions']
numAlternatives = output['alternatives']
maxTotalScore = output['totalscore']
numSeries= output['permutations'] # number of series
twoOptions=["onmogelijk","mogelijk"] #elimination options should be first

############################

correctAnswers = numpy.loadtxt("../"+ nameTest + "/sleutel.txt",delimiter=',',dtype=numpy.str)
permutations = numpy.loadtxt("../"+ nameTest + "/permutatie.txt",delimiter=',',dtype=numpy.int32)
weightsQuestions = numpy.loadtxt("../"+ nameTest + "/gewichten.txt",delimiter=',',dtype=numpy.int32)
############################
#All input is now defined, check if calculations can start        
if not( checkInputVariables_general.checkInputVariables(nameFile,nameSheet,numQuestions,numAlternatives,numSeries,correctAnswers,permutations,weightsQuestions,twoOptions)):
     sys.exit("ERROR found in input variables. Check higher for ERROR message.")
     
############################
     
#Make non-existent directories, generate keys for other permutations
if not os.path.exists("../"+ nameTest + "/Results/Sleutels"):
    os.makedirs("../"+ nameTest + "/Results/Sleutels")
if not os.path.exists("../"+ nameTest + "/Results/Histogram"):
    os.makedirs("../"+ nameTest + "/Results/Histogram")
correctAnswersDifferentPermutations = []
for i in xrange(len(permutations)):
    numpy.savetxt("../"+ nameTest + "/Results/Sleutels" + "/sleutel"+ "_reeks" + str(i+1)+".txt",  [correctAnswers[x-1] for x in permutations[i]], delimiter=" ", fmt="%s")

plt.close("all")

############################
#create list of expected content of scan file
content = ["studentennummer","vragenreeks"]
alternatives=dict()
for question in xrange(1,numQuestions+1):
    alternatives[question] = list(string.ascii_uppercase)[0:numAlternatives[question]]
    for alternative in alternatives[question]:
        name = "Vraag" + str(question) + alternative
        content.append(name)
###########################
#  
## read file and get sheet
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

(content_colNrs,errorIndicator) = supportFunctions_general.giveContentColNrs(content, sheet);
if errorIndicator == 1:
    sys.exit("ERROR: Some questions are not present in the input file. Check above for details.")
scoreQuestionsIndicatedSeries= numpy.zeros((numParticipants,numQuestions))


# write to excel_file
outputbook = Workbook(style_compression=2)
outputStudentbook = Workbook(style_compression=2)


name = "studentennummer"
studentenNrCol= content_colNrs[content.index(name)]
deelnemers=sheet.col_values(studentenNrCol,1,num_rows)

# check for double participants
if not supportFunctions_general.checkForUniqueParticipants(deelnemers):
    sys.exit("ERROR: Duplicate participants found or lack of participants. Check above for details.")

name = "vragenreeks"
#get the column in which the vragenreeks is stored
colNrSerie = content_colNrs[content.index(name)]
#get the series for the participants (so skip for row with name of first row)
columnSeries=sheet.col_values(colNrSerie,1,num_rows)

#Check if content in OMR can be used for calculations ("vragenreeks" between 0 and numSeries, answers one of twoOptions)
if not supportFunctions_general.checkForCorrectContent(deelnemers,columnSeries,numSeries,sheet,twoOptions):
    sys.exit("ERROR: The input file contains input values that are not allowed. Check above for details.")
#######################
#Ready to start the calculations

#get the score for all permutations for each of the questions
scoreQuestionsAllPermutations= supportFunctions_general.calculateScoreAllPermutations(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs,twoOptions)
numOnmogelijkQuestionsAlternatives, numMogelijkQuestionsAlternatives = supportFunctions_general.getNumberMogelijkOnmogelijk(sheet,content,permutations,columnSeries,scoreQuestionsIndicatedSeries,alternatives,twoOptions,content_colNrs)

matrixAnswers = supportFunctions_general.getMatrixAnswers(sheet,content,correctAnswers,permutations,alternatives,numParticipants,columnSeries,content_colNrs,twoOptions)     

#get the scores for the indicated series
scoreQuestionsIndicatedSeries, averageScoreQuestions =  supportFunctions_general.getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations,columnSeries)
#get the overall statistics
totalScore, averageScore, medianScore, standardDeviation, percentagePass = supportFunctions_general.getOverallStatistics(scoreQuestionsIndicatedSeries,maxTotalScore,weightsQuestions)

totalScoreDifferentPermutations = supportFunctions_general.calculateTotalScoreDifferentPermutations(scoreQuestionsAllPermutations,maxTotalScore,weightsQuestions)

numParticipantsSeries, averageScoreSeries, medianScoreSeries, standardDeviationSeries, percentagePassSeries, averageScoreQuestionsDifferentSeries = supportFunctions_general.getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations,scoreQuestionsIndicatedSeries,columnSeries,maxTotalScore)

totalScoreUpper,totalScoreMiddle,totalScoreLower,averageScoreUpper, averageScoreMiddle, averageScoreLower, averageScoreQuestionsUpper, averageScoreQuestionsMiddle, averageScoreQuestionsLower, numOnmogelijkQuestionsAlternativesUpper, numOnmogelijkQuestionsAlternativesMiddle, numOnmogelijkQuestionsAlternativesLower, numMogelijkQuestionsAlternativesUpper, numMogelijkQuestionsAlternativesMiddle, numMogelijkQuestionsAlternativesLower, scoreQuestionsUpper, scoreQuestionsMiddle, scoreQuestionsLower, numUpper, numMiddle, numLower = supportFunctions_general.calculateUpperLowerStatistics(sheet,content,columnSeries,totalScore,scoreQuestionsIndicatedSeries,correctAnswers,alternatives,twoOptions,content_colNrs,permutations)
 
totalVariance, Variance = supportFunctions_general.calculateVariances(totalScore,scoreQuestionsIndicatedSeries,numQuestions,weightsQuestions)

itemToetsCorrelatie = supportFunctions_general.calculateItemToetsCorrelatie(totalScore,scoreQuestionsIndicatedSeries,numQuestions,weightsQuestions)

########################

#Calculations are done, ready to write to output file

# WRITING THE OUTPUT TO A FILE
writeResults_general.write_results(outputbook,weightsQuestions,numQuestions,correctAnswers,alternatives,maxTotalScore,content,content_colNrs,
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
                 
# WRITING A FILE TO UPLOAD TO TOLEDO WITH THE GRADES
writeResults_general.write_scoreStudents(outputStudentbook,"punten",permutations,weightsQuestions,numParticipants,deelnemers, numQuestions,numAlternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)          

writeResults_general.write_scoreStudentsNonPermutated(outputStudentbook,"verwerking",permutations,weightsQuestions,numParticipants,deelnemers, numQuestions,numAlternatives,alternatives,content,content_colNrs,totalScore,scoreQuestionsIndicatedSeries,columnSeries,matrixAnswers)           
                  
writeResults_general.write_CronbachsAlpha(outputbook,"Cronbach's alpha", numQuestions, totalVariance, Variance, weightsQuestions)

writeResults_general.write_itemToetsCorrelatie(outputbook,"Item-toets correlatie", numQuestions, itemToetsCorrelatie,weightsQuestions)
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
for question in xrange(1,numQuestions+1):
    scoreWrongAnswer_ET_rummens = -1.0/(numAlternatives[question]-1)
    # scores for y correctly eliminated answers
    scoreCorrectAnswer_ET_rummensyEL = [(1.0*y)/((numAlternatives[question]-y)*(numAlternatives[question]-1)) for y in numpy.arange(0,numAlternatives[question])]
    possibleScores=numpy.concatenate(([scoreWrongAnswer_ET_rummens],scoreCorrectAnswer_ET_rummensyEL))
    catLow = possibleScores - 1.0/10.0
    catHigh = possibleScores + 1.0/10.0
    binsHist = numpy.sort(numpy.concatenate((catLow,catHigh)))
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist(scoreQuestionsIndicatedSeries[:,question-1],bins=binsHist)
    plt.xticks(possibleScores)
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives[question]-1),1+1.0/(numAlternatives[question]-1)])
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

for question in xrange(1,numQuestions+1):
    ax = plt.subplot(numRowsPict,numColsPict,question)
    n, bins, patches = plt.hist([scoreQuestionsUpper[:,question-1], scoreQuestionsMiddle[:,question-1], scoreQuestionsLower[:,question-1]],bins=binsHist, stacked=True,  label=['Upper', 'Middle', 'Lower'],color=['g','b','r'])
    plt.xticks(possibleScores)
    plt.title("vraag " + str(question))
    plt.xlabel("score")
    plt.xlim([-2.0/(numAlternatives[question]-1),1+1.0/(numAlternatives[question]-1)])
    plt.ylabel("aantal studenten")
    plt.legend(loc=2,prop={'size':6})
figManager = plt.get_current_fig_manager()
figManager.window.showMaximized()    
plt.savefig("../"+ nameTest + "/Results/Histogram" + "/histogramVragenUML.png", bbox_inches='tight')


outputbook.save("../"+ nameTest + "/Results/output.xls") 
outputStudentbook.save("../"+ nameTest + "/Results/punten.xls")
