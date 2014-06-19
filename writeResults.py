# -*- coding: utf-8 -*-
"""
Created on Wed May 21 14:58:24 2014

@author: tdelaet
"""

from xlwt import  easyxf
import numpy
import string

style_title = easyxf("font: bold on; align: horiz center; border: bottom medium, right medium, left medium, top medium")
style_header = easyxf("font: bold on; align: horiz center; border: bottom medium")
style_header_borderRight = easyxf("font: bold on; align: horiz center; border: right medium ")
style_correctAnswer = easyxf('pattern: pattern solid, fore_colour gray25; font: italic true')
style_specialAttention  = easyxf('font: color red')
style_correctAnswerSpecialAttention = easyxf('font: color red, italic true; pattern: pattern solid, fore_colour gray25')

style_border_header = easyxf('border: left thick, top thick, bottom thick, right thick')


def write_results(outputbook,numQuestions,correctAnswers,alternatives,maxTotalScore,content,content_colNrs,
                  columnSeries,deelnemers,
                  numParticipants,
                  totalScore,percentagePass,
                  scoreQuestionsIndicatedSeries,
                  totalScoreDifferentPermutations,
                  medianScore,
                  averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,
                  averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower,
                  numUpper,numMiddle,numLower,
                  numParticipantsSeries,
                  averageScoreSeries,medianScoreSeries,percentagePassSeries,
                  numOnmogelijkQuestionsAlternatives, numMogelijkQuestionsAlternatives,
                  numOnmogelijkQuestionsAlternativesUpper,numOnmogelijkQuestionsAlternativesMiddle,numOnmogelijkQuestionsAlternativesLower,                  
                  numMogelijkQuestionsAlternativesUpper,numMogelijkQuestionsAlternativesMiddle,numMogelijkQuestionsAlternativesLower
                  ):
                      
                      
    write_scoreAllPermutations(outputbook,'ScoreVerschillendeSeries',numParticipants,deelnemers,numQuestions,content,content_colNrs,totalScore,totalScoreDifferentPermutations,columnSeries)
    write_overallStatistics(outputbook,'GlobaleParameters',totalScore,averageScore,medianScore,percentagePass,maxTotalScore)
    write_overallStatisticsDifferentPermutations(outputbook,'GlobaleParametersSeries',numParticipantsSeries,averageScoreSeries,medianScoreSeries,percentagePassSeries,maxTotalScore)
    write_averageScoreQuestions(outputbook,'GemiddeldeScoreVraag',numQuestions,averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower)   
    write_percentageImpossibleQuestions(outputbook,"PercentageOnmogelijk",numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternatives,numParticipants)
    write_numberImpossibleQuestions(outputbook,"AantalOnmogelijk",numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternatives,numParticipants)
    write_percentagePossibleQuestions(outputbook,"PercentageMogelijk",numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternatives,numParticipants)
    write_numberPossibleQuestions(outputbook,"AantalMogelijk",numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternatives,numParticipants)
    write_percentageImpossibleQuestionsUML(outputbook,"PercentageOnmogelijkUML",numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternativesUpper,numOnmogelijkQuestionsAlternativesMiddle,numOnmogelijkQuestionsAlternativesLower,numUpper,numMiddle,numLower)
    write_numberImpossibleQuestionsUML(outputbook,"AantalOnmogelijkUML",numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternativesUpper,numOnmogelijkQuestionsAlternativesMiddle,numOnmogelijkQuestionsAlternativesLower)
    write_percentagePossibleQuestionsUML(outputbook,"PercentageMogelijkUML",numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternativesUpper,numMogelijkQuestionsAlternativesMiddle,numMogelijkQuestionsAlternativesLower,numUpper,numMiddle,numLower)
    write_numberPossibleQuestionsUML(outputbook,"AantalMogelijkUML",numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternativesUpper,numMogelijkQuestionsAlternativesMiddle,numMogelijkQuestionsAlternativesLower)
    write_histogramQuestions(outputbook,"HistogramVragen",numQuestions,scoreQuestionsIndicatedSeries,averageScoreQuestions)




##TODO: make different styles global parameters
def write_scoreAllPermutations(outputbook_loc,nameSheet_loc,numParticipants_loc,deelnemers_loc, numQuestion_loc,content_loc,content_colNrs_loc,totalScore_loc,totalScoreDifferentPermutations_loc,columnSeries_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)


    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Score deelnemers voor verschillende serie",style=style_title)
    rowCounter+=1
    
    numSeries_loc = len(totalScoreDifferentPermutations_loc[0])

    #deelnemersnummers
        #print deelnemers
    sheetC.write(rowCounter, 0,"studentennummer", style=style_header_borderRight) 
    rowCounter+=1
    for i in xrange(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=style_header_borderRight)
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 1;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score aangeduide serie ",style=style_header)
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    #total score for different series
    for serie in xrange(1,numSeries_loc+1):
        rowCounter = 1;
        sheetC.write(rowCounter,columnCounter,"totale score serie " + str(serie),style=style_header)
        rowCounter+=1
        totalScoreSerie = totalScoreDifferentPermutations_loc[:,serie-1]
        for i in xrange(len(totalScore_loc)):
            # if the series is the same as the one indicated 
            if (serie == columnSeries_loc[i]):
                sheetC.write(rowCounter,columnCounter,totalScoreSerie[i],style=style_correctAnswer)
            else:# the series is different fromthe one indicated 
                if (totalScoreSerie[i]>totalScore_loc[i]):  # if score on other serie than the one indicated is higer        
                    sheetC.write(rowCounter,columnCounter,totalScoreSerie[i],style=style_specialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,totalScoreSerie[i])
            rowCounter+=1                    
        columnCounter+=1;                    
    
def write_overallStatistics(outputbook_loc,nameSheet_loc,totalScore_loc,averageScore_loc,medianScore_loc,percentagePass_loc,maxTotalScore_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek",style=style_title)
    rowCounter+=1
    
    numParticipants_loc = len(totalScore_loc)
    #print numParticipants_loc
    #column counter
    columnCounter = 0;
    rowCounter = 1 
    
    sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=style_header)
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,numParticipants_loc)
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=style_header)
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(averageScore_loc,2))
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"mediaan ",style=style_header)
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(medianScore_loc,2))
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=style_header)
    columnCounter+=1 
    #print totalScore_loc
    sheetC.write(rowCounter,columnCounter,round(percentagePass_loc,2))
    
    
def write_overallStatisticsDifferentPermutations(outputbook_loc,nameSheet_loc,numParticipantsSeries_loc,averageScoreSeries_loc,medianScoreSeries_loc,percentagePassSeries_loc,maxTotalScore_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek verschillende reeksen",style=style_title)
    rowCounter+=1
    
    numSeries = len(numParticipantsSeries_loc)
    #print numParticipants_loc
    #column counter
    
    for serie in xrange(numSeries):
        columnCounter = 0;        
        sheetC.write(rowCounter,columnCounter,"serie " + str(serie+1),style=style_header)
        rowCounter+=1
        
        sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=style_header)
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,numParticipantsSeries_loc[serie]) 
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=style_header)
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(averageScoreSeries_loc[serie],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"mediaan ",style=style_header)
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(medianScoreSeries_loc[serie],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=style_header)
        columnCounter+=1 
        #print totalScore_loc
        sheetC.write(rowCounter,columnCounter,round(percentagePassSeries_loc[serie],2))
        
        rowCounter+=1
        rowCounter+=1
      
def write_averageScoreQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,averageScore_loc,averageScoreUpper_loc,averageScoreMiddle_loc,averageScoreLower_loc,averageScoreQuestions_loc,averageScoreQuestionsUpper_loc,averageScoreQuestionsMiddle_loc,averageScoreQuestionsLower_loc):
    sheetC = outputbook_loc.add_sheet('GemScorePerVraag')
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Gemiddelde score per vraag",style=style_title)
    rowCounter+=1
    
    #column counter
    columnCounter = 0; 

    #write all/upper/middle/lower on top
    columnCounter = 1
    sheetC.write(rowCounter,columnCounter,"all",style=style_header)
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"upper",style=style_header)
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"middle",style=style_header)
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"lower",style=style_header)
    columnCounter+=1
    rowCounter+=1
    columnCounter=0
    
    for question in xrange(1,numQuestions_loc+1):

        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1        
        if averageScoreQuestions_loc[question-1]<0:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3),style=style_specialAttention)        
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3))                
        columnCounter+=1 
        if averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsLower_loc[question-1] or averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsMiddle_loc[question-1]:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3),style=style_specialAttention)
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3),style=style_specialAttention)
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3),style=style_specialAttention)    
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3))
        rowCounter+=1
        columnCounter = 0;
        
    columnCounter=0
    sheetC.write(rowCounter,columnCounter,"total",style=style_header_borderRight)
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScore_loc,3))   
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreUpper_loc,3))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreMiddle_loc,3))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreLower_loc,3))  
    columnCounter+=1


def write_percentageImpossibleQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage onmogelijk per alternatief",style=style_title)
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=style_header)    
        columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc > 0.35):#TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style_correctAnswer)                    
            else:
                if (numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc < 0.35):#TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style_specialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2))
            columnCounter+=1
            counter+=1  
        rowCounter+=1

def write_numberImpossibleQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal onmogelijk per alternatief",style=style_title)
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=style_header)    
        columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numOnmogelijkQuestionsAlternatives_loc[counter] > 0.35 * numParticipants_loc): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter],style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter],style_correctAnswer)
            else:
                if (numOnmogelijkQuestionsAlternatives_loc[counter] < 0.35 * numParticipants_loc): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter],style_specialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter])
            columnCounter+=1
            counter+=1    
        rowCounter+=1

def write_percentagePossibleQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage mogelijk per alternatief",style=style_title)
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=style_header)    
        columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc < 0.65): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style_correctAnswer)
            else:
                if (numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc > 0.65): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style_specialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2))
            columnCounter+=1
            counter+=1    
        rowCounter+=1
        
def write_numberPossibleQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal mogelijk per alternatief",style=style_title)
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=style_header)    
        columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numMogelijkQuestionsAlternatives_loc[counter] < 0.65 * numParticipants_loc): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter],style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter],style_correctAnswer)
            else:
                if (numMogelijkQuestionsAlternatives_loc[counter] > 0.65 * numParticipants_loc):#TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter],style_specialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter])
            columnCounter+=1
            counter+=1    
        rowCounter+=1
            
def write_percentageImpossibleQuestionsUML(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternativesUpper_loc,numOnmogelijkQuestionsAlternativesMiddle_loc,numOnmogelijkQuestionsAlternativesLower_loc,numUpper_loc,numMiddle_loc,numLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage onmogelijk per alternatief voor upper, middle en lower",style=style_title)
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(rowCounter+1,columnCounter,"upper",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3   
    rowCounter+=2
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord meer onmogelijk aanduidt dan lower groep
                if ( (numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc > numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc) ):
                    #TODO make 20 parameter
                    sheetC.write(rowCounter,columnCounter  ,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style_correctAnswerSpecialAttention)
                else:
                    #TODO make 20 parameter
                    if(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc>0.25):                   
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_correctAnswerSpecialAttention)
                    else:
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord minder onmogelijk aanduidt dan lower groep or if upper group percentage is lower than fixed number
                if (numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc < numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc):
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                    if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] < numOnmogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_specialAttention)
                    else:
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2))
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2))
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2))
            columnCounter+=3
            counter+=1  
        rowCounter+=1
 
def write_percentagePossibleQuestionsUML(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternativesUpper_loc,numMogelijkQuestionsAlternativesMiddle_loc,numMogelijkQuestionsAlternativesLower_loc,numUpper_loc,numMiddle_loc,numLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage mogelijk per alternatief voor upper, middle en lower",style=style_title)
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(rowCounter+1,columnCounter,"upper",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3
    rowCounter+=2
    
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord minder mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc < numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc):
                    sheetC.write(rowCounter,columnCounter  ,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style_correctAnswerSpecialAttention)
                else:
                    if(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc<0.75):  
                        sheetC.write(rowCounter,columnCounter  ,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_correctAnswerSpecialAttention)
                    else:
                        sheetC.write(rowCounter,columnCounter  ,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord meer mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc > numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc):
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                    if (numMogelijkQuestionsAlternativesUpper_loc[counter] > numMogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style_specialAttention)
                    else:
                        sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2))
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2))
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2))
            columnCounter+=3
            counter+=1 
        rowCounter+=1
        
def write_numberImpossibleQuestionsUML(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternativesUpper_loc,numOnmogelijkQuestionsAlternativesMiddle_loc,numOnmogelijkQuestionsAlternativesLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    counter=0
    #write alternative names on top
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal onmogelijk per alternatief voor upper, middle en lower",style=style_title)
    rowCounter+=1
    
    columnCounter = 1
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(rowCounter+1,columnCounter,"upper",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3
    rowCounter+=2
    
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord meer onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] > numOnmogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter],style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter],style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter],style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord minder onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] < numOnmogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter],style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter],style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                    if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] < numOnmogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style_specialAttention)
                    else:
                        sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternativesUpper_loc[counter])
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter])
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter])
            columnCounter+=3
            counter+=1    
        rowCounter+=1
 
def write_numberPossibleQuestionsUML(outputbook_loc,nameSheet_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternativesUpper_loc,numMogelijkQuestionsAlternativesMiddle_loc,numMogelijkQuestionsAlternativesLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    counter=0
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal mogelijk per alternatief voor upper, middle en lower",style=style_title)
    rowCounter+=1
    
    columnCounter = 1
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=style_header)    
        sheetC.write(rowCounter+1,columnCounter,"upper",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=style_header)
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=style_header)    
        columnCounter+=3
    rowCounter+=2
    
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord minder mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter] < numMogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter  ,numMogelijkQuestionsAlternativesUpper_loc[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter],style_correctAnswerSpecialAttention)
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter],style_correctAnswerSpecialAttention)
                else:
                    sheetC.write(rowCounter,columnCounter  ,numMogelijkQuestionsAlternativesUpper_loc[counter],style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter],style_correctAnswer)
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter],style_correctAnswer)                
            else:
                # test of uppergroep een fout antwoord meer mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter] > numMogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternativesUpper_loc[counter],style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter],style_specialAttention)
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter],style_specialAttention)
                else:
                    # test of uppergroep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                    if (numMogelijkQuestionsAlternativesUpper_loc[counter] > numMogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternativesUpper_loc[counter],style_specialAttention)
                    else:
                        sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternativesUpper_loc[counter])
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter])
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter])
            columnCounter+=3
            counter+=1 
        rowCounter+=1
     
def write_histogramQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,scoreQuestionsIndicatedSeries_loc,averageScoreQuestions_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
        
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Histogram score vragen",style=style_title)
    rowCounter+=1
    
    #column counter
    columnCounter = 0;
    #gemiddelde verdeling scores per vraag
    #counter=0
    possibleScores=numpy.arange(-1.0,1+1.0/2.0,1.0/3.0)
    #print possibleScores
    columnCounter = 1
    for possibleScore in possibleScores[0:len(possibleScores)-1]:
        sheetC.write(rowCounter,columnCounter,possibleScore,style=style_header)
        columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"gemiddelde",style=style_header)
    rowCounter+=1    
    
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=style_header_borderRight)
        hist,bins = numpy.histogram(scoreQuestionsIndicatedSeries_loc[:,question-1],bins=possibleScores-1.0/6.0)
        columnCounter+=1    
        for n in hist:        
            if (hist[0]>hist[len(hist)-1] or hist[0]+hist[1]>hist[len(hist)-1]+hist[len(hist)-2]): #more confident in wrong answer than confident in correct answer
                sheetC.write(rowCounter,columnCounter,n,style=style_specialAttention)
            else:
                sheetC.write(rowCounter,columnCounter,n)
            columnCounter+=1
        if averageScoreQuestions_loc[question-1]<0:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2),style=style_specialAttention)        
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2)) 
        rowCounter+=1

def write_scoreStudents(outputbook_loc,nameSheet_loc,permutations_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,numAlternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    
    #deelnemersnummers
    sheetC.write(rowCounter, 0,"studentennummer", style=style_header_borderRight) 
    rowCounter+=1
    for i in xrange(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=style_header_borderRight)
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score",style=style_header)
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i])
        rowCounter+=1
    columnCounter+=1;
    

    rowCounter = 0;
    #indicated series
    sheetC.write(rowCounter,columnCounter,"reeks",style=style_header)
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,columnSeries_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    #score for different questions
    #beware scores are stored per question without the permutation; 
    #so for the student the scores have to be back-permutated to the order they got
    rowCounter = 0;
    #write heading    
    columnCounterHeader = columnCounter
    for question in xrange(1,numQuestions_loc+1):
        sheetC.write(rowCounter,columnCounterHeader,"score vraag " + str(question),style=style_header)
        columnCounterHeader+=1
        
    columnCounterScoreQuestions = columnCounter   
    rowCounter = 1;  
    for participant in xrange(len(totalScore_loc)): # loop over participants
        columnCounter = columnCounterScoreQuestions;
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        sorted_score = [score[i-1] for i in permutations_loc[int(columnSeries_loc[participant]-1)]]
        for question in xrange(1,numQuestions_loc+1):
            sheetC.write(rowCounter,columnCounter,sorted_score[question-1])
            columnCounter+=1
        rowCounter+=1            
 
    
    #print matrixAnswers
    alternatives = list(string.ascii_uppercase)[0:numAlternatives_loc]   
    #answer for different questions and alternatives
    questionAlternativeCounter = 0
    for question in xrange(1,numQuestions_loc+1):
        for alternative in alternatives:
            rowCounter=0
            sheetC.write(rowCounter,columnCounter,"antwoord vraag " + str(question) + str(alternative),style=style_header)
            #print "antwoord vraag " + str(question) + str(alternative)
            answer = matrixAnswers[:,questionAlternativeCounter]
            questionAlternativeCounter+=1
            rowCounter+=1
            #print 
            for i in xrange(len(totalScore_loc)):
                sheetC.write(rowCounter,columnCounter,answer[i])
                rowCounter+=1                    
            columnCounter+=1;    
            