# -*- coding: utf-8 -*-
"""
Created on Wed May 21 14:58:24 2014

@author: tdelaet
"""

from xlwt import  easyxf
import numpy
import string

font_bold = "font: bold on;"
font_red = "font: color red;"
font_italic = "font: italic true;"
align_horizcenter = "align: horiz center;"
align_vertcenter = "align: vert centre;"
align_horizvertcenter = "align: vert centre, horiz center;"
border_bottom_medium = "border: bottom medium;"
border_top_medium = "border: top medium;"
border_right_medium = "border: right medium;"
border_left_medium = "border: left medium;"
border_righttop_medium = "border: right medium, top medium;"
border_all_medium = "border: bottom medium, right medium, left medium, top medium;"
pattern_solid_grey = "pattern: pattern solid, fore_colour gray25;"

style_title = font_bold + border_all_medium + align_horizvertcenter
style_header = font_bold + border_bottom_medium + align_horizvertcenter
style_header_borderRight = style_header + border_right_medium
style_correctAnswer = pattern_solid_grey + font_italic
style_specialAttention = font_red

def write_results(outputbook,weightsQuestions,numQuestions,correctAnswers,alternatives,maxTotalScore,content,content_colNrs,
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
                  ):
                      
                      
    write_scoreAllPermutations(outputbook,'ScoreVerschillendeSeries',numParticipants,deelnemers,numQuestions,content,content_colNrs,totalScore,totalScoreDifferentPermutations,columnSeries)
    write_overallStatistics(outputbook,'GlobaleParameters',totalScore,averageScore,medianScore,standardDeviation,percentagePass,numParticipantsSeries,averageScoreSeries,medianScoreSeries,standardDeviationSeries,percentagePassSeries,maxTotalScore)
    #write_overallStatistics(outputbook,'GlobaleParameters',totalScore,averageScore,medianScore,percentagePass,maxTotalScore)
    #write_overallStatisticsDifferentPermutations(outputbook,'GlobaleParametersSeries',numParticipantsSeries,averageScoreSeries,medianScoreSeries,percentagePassSeries,maxTotalScore)
    #write_averageScoreQuestions(outputbook,'GemiddeldeScoreVraag',numQuestions,averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower,averageScoreQuestionsDifferentSeries)   
    write_averageScoreQuestions(outputbook,'GemiddeldeScoreVraag',weightsQuestions,numQuestions,averageScore,averageScoreUpper,averageScoreMiddle,averageScoreLower,averageScoreQuestions,averageScoreQuestionsUpper,averageScoreQuestionsMiddle,averageScoreQuestionsLower,averageScoreSeries,averageScoreQuestionsDifferentSeries)       
    write_percentageImpossibleQuestions(outputbook,"PercentageOnmogelijk",weightsQuestions,numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternatives,numParticipants)
    write_numberImpossibleQuestions(outputbook,"AantalOnmogelijk",weightsQuestions,numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternatives,numParticipants)
    write_percentagePossibleQuestions(outputbook,"PercentageMogelijk",weightsQuestions,numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternatives,numParticipants)
    write_numberPossibleQuestions(outputbook,"AantalMogelijk",weightsQuestions,numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternatives,numParticipants)
    write_percentageImpossibleQuestionsUML(outputbook,"PercentageOnmogelijkUML",weightsQuestions,numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternativesUpper,numOnmogelijkQuestionsAlternativesMiddle,numOnmogelijkQuestionsAlternativesLower,numUpper,numMiddle,numLower)
    write_numberImpossibleQuestionsUML(outputbook,"AantalOnmogelijkUML",weightsQuestions,numQuestions,correctAnswers,alternatives,numOnmogelijkQuestionsAlternativesUpper,numOnmogelijkQuestionsAlternativesMiddle,numOnmogelijkQuestionsAlternativesLower)
    write_percentagePossibleQuestionsUML(outputbook,"PercentageMogelijkUML",weightsQuestions,numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternativesUpper,numMogelijkQuestionsAlternativesMiddle,numMogelijkQuestionsAlternativesLower,numUpper,numMiddle,numLower)
    write_numberPossibleQuestionsUML(outputbook,"AantalMogelijkUML",weightsQuestions,numQuestions,correctAnswers,alternatives,numMogelijkQuestionsAlternativesUpper,numMogelijkQuestionsAlternativesMiddle,numMogelijkQuestionsAlternativesLower)
    write_histogramQuestions(outputbook,"HistogramVragen",weightsQuestions,numQuestions,scoreQuestionsIndicatedSeries,averageScoreQuestions)




##TODO: make different styles global parameters
def write_scoreAllPermutations(outputbook_loc,nameSheet_loc,numParticipants_loc,deelnemers_loc, numQuestion_loc,content_loc,content_colNrs_loc,totalScore_loc,totalScoreDifferentPermutations_loc,columnSeries_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)


    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Score deelnemers voor verschillende serie",style=easyxf(style_title))
    rowCounter+=1
    
    numSeries_loc = len(totalScoreDifferentPermutations_loc[0])

    #deelnemersnummers
        #print deelnemers
    sheetC.write(rowCounter, 0,"studentennummer", style=easyxf(style_header_borderRight))
    rowCounter+=1
    for i in xrange(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(font_bold+border_right_medium))
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 1;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score aangeduide serie ",style=easyxf(font_bold+border_all_medium))
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i],style=easyxf(border_right_medium))
        rowCounter+=1
    columnCounter+=1;
    
    #total score for different series
    for serie in xrange(1,numSeries_loc+1):
        rowCounter = 1;
        sheetC.write(rowCounter,columnCounter,"totale score serie " + str(serie),style=easyxf(style_header))
        rowCounter+=1
        totalScoreSerie = totalScoreDifferentPermutations_loc[:,serie-1]
        for i in xrange(len(totalScore_loc)):
            # if the series is the same as the one indicated 
            if (serie == columnSeries_loc[i]):
                sheetC.write(rowCounter,columnCounter,totalScoreSerie[i],style=easyxf(style_correctAnswer))
            else:# the series is different fromthe one indicated 
                if (totalScoreSerie[i]>totalScore_loc[i]):  # if score on other serie than the one indicated is higer        
                    sheetC.write(rowCounter,columnCounter,totalScoreSerie[i],style=easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,totalScoreSerie[i])
            rowCounter+=1                    
        columnCounter+=1;                    
    
def write_overallStatistics(outputbook_loc,nameSheet_loc,totalScore_loc,averageScore_loc,medianScore_loc,standardDeviation_loc,percentagePass_loc,numParticipantsSeries_loc,averageScoreSeries_loc,medianScoreSeries_loc,standardDeviationSeries_loc,percentagePassSeries_loc,maxTotalScore_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek",style=easyxf(style_title))
    rowCounter+=1
    
    numParticipants_loc = len(totalScore_loc)
    #print numParticipants_loc
    #column counter
    columnCounter = 0;
    rowCounter = 1 
    
    sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,numParticipants_loc)
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(averageScore_loc,2))
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"mediaan ",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(medianScore_loc,2))
    rowCounter+=1
    
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"standaard deviatie",style=easyxf(font_bold))
    columnCounter+=1  
    sheetC.write(rowCounter,columnCounter,round(standardDeviation_loc,2))
    rowCounter+=1
        
    columnCounter = 0
    sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=easyxf(font_bold))
    columnCounter+=1 
    #print totalScore_loc
    sheetC.write(rowCounter,columnCounter,round(percentagePass_loc,2))
    

    rowCounter+=5
    columnCounter = 0
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Globale statistiek verschillende reeksen",style=easyxf(style_title))
    rowCounter+=1
    
    numSeries = len(numParticipantsSeries_loc)
    #print numParticipants_loc
    #column counter
    
    for serie in xrange(numSeries):
        columnCounter = 0;        
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+1,"serie " + str(serie+1),style=easyxf(font_bold+border_bottom_medium))
        rowCounter+=1
        
        sheetC.write(rowCounter,columnCounter,"aantal deelnemers",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,numParticipantsSeries_loc[serie]) 
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"gemiddelde score ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(averageScoreSeries_loc[serie],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"mediaan ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(medianScoreSeries_loc[serie],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"standaard deviatie ",style=easyxf(font_bold))
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,round(standardDeviationSeries_loc[serie],2))
        rowCounter+=1
        
        columnCounter = 0
        sheetC.write(rowCounter,columnCounter,"% geslaagd ",style=easyxf(font_bold))
        columnCounter+=1 
        #print totalScore_loc
        sheetC.write(rowCounter,columnCounter,round(percentagePassSeries_loc[serie],2))
        
        rowCounter+=1
        rowCounter+=1
      
#def write_averageScoreQuestions(outputbook_loc,nameSheet_loc,numQuestions_loc,averageScore_loc,averageScoreUpper_loc,averageScoreMiddle_loc,averageScoreLower_loc,averageScoreQuestions_loc,averageScoreQuestionsUpper_loc,averageScoreQuestionsMiddle_loc,averageScoreQuestionsLower_loc):
#    sheetC = outputbook_loc.add_sheet('GemScorePerVraag')
#    
#    columnCounter = 0;
#    rowCounter = 0;
#    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Gemiddelde score per vraag",style=easyxf(style_title))
#    rowCounter+=1
#    
#    #column counter
#    columnCounter = 0; 
#
#    #write all/upper/middle/lower on top
#    columnCounter = 1
#    sheetC.write(rowCounter,columnCounter,"all",style=easyxf(style_header))
#    columnCounter+=1
#    sheetC.write(rowCounter,columnCounter,"upper",style=easyxf(style_header))
#    columnCounter+=1
#    sheetC.write(rowCounter,columnCounter,"middle",style=easyxf(style_header))
#    columnCounter+=1
#    sheetC.write(rowCounter,columnCounter,"lower",style=easyxf(style_header))
#    columnCounter+=1
#    rowCounter+=1
#    columnCounter=0
#    
#    for question in xrange(1,numQuestions_loc+1):
#
#        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(style_header_borderRight))
#        columnCounter+=1        
#        if averageScoreQuestions_loc[question-1]<0:
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3),style=easyxf(style_specialAttention))        
#        else:
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3))                
#        columnCounter+=1 
#        if averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsLower_loc[question-1] or averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsMiddle_loc[question-1]:
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3),style=easyxf(style_specialAttention))
#            columnCounter+=1
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3),style=easyxf(style_specialAttention))
#            columnCounter+=1
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3),style=easyxf(style_specialAttention))    
#        else:
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3))
#            columnCounter+=1
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3))
#            columnCounter+=1
#            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3))
#        rowCounter+=1
#        columnCounter = 0;
#        
#    columnCounter=0
#    sheetC.write(rowCounter,columnCounter,"total",style=easyxf(style_header_borderRight))
#    columnCounter+=1
#    sheetC.write(rowCounter,columnCounter,round(averageScore_loc,3))   
#    columnCounter+=1
#    sheetC.write(rowCounter,columnCounter,round(averageScoreUpper_loc,3))
#    columnCounter+=1
#    sheetC.write(rowCounter,columnCounter,round(averageScoreMiddle_loc,3))
#    columnCounter+=1
#    sheetC.write(rowCounter,columnCounter,round(averageScoreLower_loc,3))  
#    columnCounter+=1
      
def write_averageScoreQuestions(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,averageScore_loc,averageScoreUpper_loc,averageScoreMiddle_loc,averageScoreLower_loc,averageScoreQuestions_loc,averageScoreQuestionsUpper_loc,averageScoreQuestionsMiddle_loc,averageScoreQuestionsLower_loc,averageScoreSeries_loc,averageScoreQuestionsDifferentSeries_loc):
    numSeries = len(averageScoreQuestionsDifferentSeries_loc[0])  
    sheetC = outputbook_loc.add_sheet('GemScorePerVraag')
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8," Gemiddelde score per vraag",style=easyxf(style_title))
    rowCounter+=1
    
    #column counter
    columnCounter = 0; 


    #write all/upper/middle/lower on top
    columnCounter = 1
    sheetC.write(rowCounter+1,columnCounter,"all",style=easyxf(style_header + border_right_medium))
    columnCounter+=1
    
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2," UML",style=easyxf(style_title))
    sheetC.write_merge(rowCounter,rowCounter,columnCounter+3,columnCounter+6," reeksen",style=easyxf(style_title))
    rowCounter+=1

    sheetC.write(rowCounter,columnCounter,"upper",style=easyxf(style_header)) 
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"middle",style=easyxf(style_header)) 
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"lower",style=easyxf(style_header+ border_right_medium)) 
    columnCounter+=1

    for serie in xrange(0,numSeries): #TODO: numseries
        sheetC.write(rowCounter,columnCounter,"reeks " + str(serie+1) ,style=easyxf(style_header))
        columnCounter+=1
    
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    
    rowCounter+=1    
    columnCounter=0
    
    for question in xrange(1,numQuestions_loc+1):
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf( font_bold+border_right_medium))
        columnCounter+=1        
        if averageScoreQuestions_loc[question-1]<0:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3),style=easyxf(style_specialAttention + border_right_medium)        )
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],3),style=easyxf(border_right_medium))                
        columnCounter+=1 
        if averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsLower_loc[question-1] or averageScoreQuestionsUpper_loc[question-1]<=averageScoreQuestionsMiddle_loc[question-1]:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3),style=easyxf(style_specialAttention))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3),style=easyxf(style_specialAttention)) 
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3),style=easyxf(style_specialAttention+ border_right_medium))    
            columnCounter+=1
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsUpper_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsMiddle_loc[question-1],3))
            columnCounter+=1
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsLower_loc[question-1],3),style=easyxf(border_right_medium))
            columnCounter+=1
        for serie in xrange(1,numSeries+1):
            #print "rowCounter" + str(rowCounter)
            #print "columnCounter" + str(columnCounter)
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestionsDifferentSeries_loc[question-1,serie-1],3))
            columnCounter+=1
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1           
        rowCounter+=1
        columnCounter = 0;
        
    columnCounter=0
    sheetC.write(rowCounter,columnCounter,"totaal",style=easyxf(style_header+border_all_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScore_loc,3),style=easyxf(border_righttop_medium))   
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreUpper_loc,3),style=easyxf(border_top_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreMiddle_loc,3),style=easyxf(border_top_medium))
    columnCounter+=1
    sheetC.write(rowCounter,columnCounter,round(averageScoreLower_loc,3),style=easyxf(border_righttop_medium))  
    columnCounter+=1
    for serie in xrange(1,numSeries+1):
        sheetC.write(rowCounter,columnCounter,round(averageScoreSeries_loc[serie-1],3),style=easyxf(border_top_medium))
        columnCounter+=1
    
        
def write_percentageImpossibleQuestions(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage onmogelijk per alternatief",style=easyxf(style_title))
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=easyxf(style_header)  )  
        columnCounter+=1
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold + border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc > 0.35):#TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style=easyxf(style_correctAnswer))                    
            else:
                if (numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc < 0.35):#TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style=easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2))
            columnCounter+=1
            counter+=1  
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1     
        rowCounter+=1
        

def write_numberImpossibleQuestions(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal onmogelijk per alternatief",style=easyxf(style_title))
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=easyxf(style_header)  )  
        columnCounter+=1
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numOnmogelijkQuestionsAlternatives_loc[counter] > 0.35 * numParticipants_loc): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter],style=easyxf(style_correctAnswer))
            else:
                if (numOnmogelijkQuestionsAlternatives_loc[counter] < 0.35 * numParticipants_loc): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter],style=easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternatives_loc[counter])
            columnCounter+=1
            counter+=1  
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1         
        rowCounter+=1

def write_percentagePossibleQuestions(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage mogelijk per alternatief",style=easyxf(style_title))
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=easyxf(style_header)    )
        columnCounter+=1
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc < 0.65): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style=easyxf(style_correctAnswer))
            else:
                if (numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc > 0.65): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2),style=easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternatives_loc[counter]/numParticipants_loc,2))
            columnCounter+=1
            counter+=1    
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1             
        rowCounter+=1
        
def write_numberPossibleQuestions(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternatives_loc,numParticipants_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal mogelijk per alternatief",style=easyxf(style_title))
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1    
    for alternative in alternatives_loc:
        sheetC.write(rowCounter,columnCounter,alternative,style=easyxf(style_header)    )
        columnCounter+=1
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=1
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                if (numMogelijkQuestionsAlternatives_loc[counter] < 0.65 * numParticipants_loc): #TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter],style=easyxf(style_correctAnswer))
            else:
                if (numMogelijkQuestionsAlternatives_loc[counter] > 0.65 * numParticipants_loc):#TODO: make parameter
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter],style=easyxf(style_specialAttention))
                else:
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternatives_loc[counter])
            columnCounter+=1
            counter+=1    
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1                 
        rowCounter+=1
            
def write_percentageImpossibleQuestionsUML(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternativesUpper_loc,numOnmogelijkQuestionsAlternativesMiddle_loc,numOnmogelijkQuestionsAlternativesLower_loc,numUpper_loc,numMiddle_loc,numLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage onmogelijk per alternatief voor upper, middle en lower",style=easyxf(style_title))
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=easyxf(style_header)  )  
        sheetC.write(rowCounter+1,columnCounter,"upper",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=easyxf(style_header+border_right_medium) )   
        columnCounter+=3   
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=2
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord meer onmogelijk aanduidt dan lower groep
                if ( (numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc > numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc) ):
                    #TODO make 20 parameter
                    sheetC.write(rowCounter,columnCounter  ,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(style_correctAnswer+style_specialAttention+border_right_medium))
                else:
                    #TODO make 20 parameter
                    if(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc>0.25):                   
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(style_correctAnswer+border_right_medium))                
            else:
                # test of uppergroep een fout antwoord minder onmogelijk aanduidt dan lower groep or if upper group percentage is lower than fixed number
                if (numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc < numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc):
                    sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(style_specialAttention+border_right_medium))
                else:
                    # test of uppergroep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                    if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] < numOnmogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter,round(numOnmogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2))
                    sheetC.write(rowCounter,columnCounter+1,round(numOnmogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2))
                    sheetC.write(rowCounter,columnCounter+2,round(numOnmogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(border_right_medium))
            columnCounter+=3
            counter+=1  
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1             
        rowCounter+=1
 
def write_percentagePossibleQuestionsUML(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternativesUpper_loc,numMogelijkQuestionsAlternativesMiddle_loc,numMogelijkQuestionsAlternativesLower_loc,numUpper_loc,numMiddle_loc,numLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Percentage mogelijk per alternatief voor upper, middle en lower",style=easyxf(style_title))
    rowCounter+=1
    
    counter=0
    #write alternative names on top
    columnCounter = 1
    
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=easyxf(style_header)    )
        sheetC.write(rowCounter+1,columnCounter,"upper",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=easyxf(style_header+border_right_medium)    )
        columnCounter+=3
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=2
    
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord minder mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc < numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc):
                    sheetC.write(rowCounter,columnCounter  ,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(style_correctAnswer+style_specialAttention+border_right_medium))
                else:
                    if(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc<0.75):  
                        sheetC.write(rowCounter,columnCounter  ,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_correctAnswer+style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter  ,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(style_correctAnswer+border_right_medium))                
            else:
                # test of uppergroep een fout antwoord meer mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc > numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc):
                    sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2),style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(style_specialAttention+border_right_medium))
                else:
                    # test of uppergroep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                    if (numMogelijkQuestionsAlternativesUpper_loc[counter] > numMogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2),style=easyxf(style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter,round(numMogelijkQuestionsAlternativesUpper_loc[counter]/numUpper_loc,2))
                    sheetC.write(rowCounter,columnCounter+1,round(numMogelijkQuestionsAlternativesMiddle_loc[counter]/numMiddle_loc,2))
                    sheetC.write(rowCounter,columnCounter+2,round(numMogelijkQuestionsAlternativesLower_loc[counter]/numLower_loc,2),style=easyxf(border_right_medium))
            columnCounter+=3
            counter+=1 
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1            
        rowCounter+=1
        
def write_numberImpossibleQuestionsUML(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numOnmogelijkQuestionsAlternativesUpper_loc,numOnmogelijkQuestionsAlternativesMiddle_loc,numOnmogelijkQuestionsAlternativesLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    counter=0
    #write alternative names on top
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal onmogelijk per alternatief voor upper, middle en lower",style=easyxf(style_title))
    rowCounter+=1
    
    columnCounter = 1
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=easyxf(style_header)   ) 
        sheetC.write(rowCounter+1,columnCounter,"upper",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=easyxf(style_header+border_right_medium)  )  
        columnCounter+=3
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=2
    
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold + border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord meer onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] > numOnmogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention+border_right_medium))
                else:
                    sheetC.write(rowCounter,columnCounter  ,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter],style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(style_correctAnswer+border_right_medium))                
            else:
                # test of uppergroep een fout antwoord minder onmogelijk aanduidt dan lower groep
                if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] < numOnmogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter],style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(style_specialAttention+border_right_medium))
                else:
                    # test of uppergroep een fout antwoord minder ommogelijk aanduidt dan goed antwoord
                    if (numOnmogelijkQuestionsAlternativesUpper_loc[counter] < numOnmogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter,numOnmogelijkQuestionsAlternativesUpper_loc[counter])
                    sheetC.write(rowCounter,columnCounter+1,numOnmogelijkQuestionsAlternativesMiddle_loc[counter])
                    sheetC.write(rowCounter,columnCounter+2,numOnmogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(border_right_medium))
            columnCounter+=3
            counter+=1    
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1         
        rowCounter+=1
 
def write_numberPossibleQuestionsUML(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,correctAnswers_loc,alternatives_loc,numMogelijkQuestionsAlternativesUpper_loc,numMogelijkQuestionsAlternativesMiddle_loc,numMogelijkQuestionsAlternativesLower_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    counter=0
    
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Aantal mogelijk per alternatief voor upper, middle en lower",style=easyxf(style_title))
    rowCounter+=1
    
    columnCounter = 1
    for alternative in alternatives_loc:
        sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+2,alternative,style=easyxf(style_header)    )
        sheetC.write(rowCounter+1,columnCounter,"upper",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+1,"middle",style=easyxf(style_header))
        sheetC.write(rowCounter+1,columnCounter+2,"lower",style=easyxf(style_header+border_right_medium)    )
        columnCounter+=3
            
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=2
    
        
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        correctAnswer = correctAnswers_loc[question-1]
        #loop over alternatives
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        columnCounter+=1
        for alternative in alternatives_loc:
            if alternative == correctAnswer:
                # test of uppergroep het correcte antwoord minder mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter] < numMogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter  ,numMogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(style_correctAnswer+style_specialAttention+border_right_medium))
                else:
                    sheetC.write(rowCounter,columnCounter  ,numMogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter],style=easyxf(style_correctAnswer))
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(style_correctAnswer+border_right_medium))                
            else:
                # test of uppergroep een fout antwoord meer mogelijk aanduidt dan lower groep
                if (numMogelijkQuestionsAlternativesUpper_loc[counter] > numMogelijkQuestionsAlternativesLower_loc[counter]):
                    sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter],style=easyxf(style_specialAttention))
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(style_specialAttention+border_right_medium))
                else:
                    # test of uppergroep een fout antwoord meer mogelijk aanduidt dan goed antwoord
                    if (numMogelijkQuestionsAlternativesUpper_loc[counter] > numMogelijkQuestionsAlternativesUpper_loc[(question-1)*len(alternatives_loc)+alternatives_loc.index(correctAnswer)] ):
                        sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternativesUpper_loc[counter],style=easyxf(style_specialAttention))
                    else:
                        sheetC.write(rowCounter,columnCounter,numMogelijkQuestionsAlternativesUpper_loc[counter])
                    sheetC.write(rowCounter,columnCounter+1,numMogelijkQuestionsAlternativesMiddle_loc[counter])
                    sheetC.write(rowCounter,columnCounter+2,numMogelijkQuestionsAlternativesLower_loc[counter],style=easyxf(border_right_medium))
            columnCounter+=3
            counter+=1 
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1    
        rowCounter+=1
     
def write_histogramQuestions(outputbook_loc,nameSheet_loc,weightsQuestions_loc,numQuestions_loc,scoreQuestionsIndicatedSeries_loc,averageScoreQuestions_loc):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)
    
        
    columnCounter = 0;
    rowCounter = 0;
    sheetC.write_merge(rowCounter,rowCounter,columnCounter,columnCounter+8,"Histogram score vragen",style=easyxf(style_title))
    rowCounter+=1
    
    #column counter
    columnCounter = 0;
    #gemiddelde verdeling scores per vraag
    #counter=0
    possibleScores=numpy.arange(-1.0,1+1.0/2.0,1.0/3.0)
    #print possibleScores
    columnCounter = 1
    for possibleScore in possibleScores[0:len(possibleScores)-1]:
        sheetC.write(rowCounter,columnCounter,possibleScore,style=easyxf(style_header))
        columnCounter+=1
    sheetC.write(rowCounter,columnCounter,"gemiddelde",style=easyxf(style_header+border_left_medium))
    columnCounter+=1
    
    sheetC.write(rowCounter,columnCounter,"gewicht vraag",style=easyxf(style_header+ border_left_medium)) 
    columnCounter+=1
    rowCounter+=1    
    
    for question in xrange(1,numQuestions_loc+1):
        columnCounter=0
        sheetC.write(rowCounter,columnCounter,"vraag"+str(question),style=easyxf(font_bold+border_right_medium))
        hist,bins = numpy.histogram(scoreQuestionsIndicatedSeries_loc[:,question-1],bins=possibleScores-1.0/6.0)
        columnCounter+=1    
        for n in hist:        
            if (hist[0]>hist[len(hist)-1] or hist[0]+hist[1]>hist[len(hist)-1]+hist[len(hist)-2]): #more confident in wrong answer than confident in correct answer
                sheetC.write(rowCounter,columnCounter,n,style=easyxf(style_specialAttention))
            else:
                sheetC.write(rowCounter,columnCounter,n)
            columnCounter+=1
        if averageScoreQuestions_loc[question-1]<0:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2),style=easyxf(style_specialAttention+border_left_medium))        
        else:
            sheetC.write(rowCounter,columnCounter,round(averageScoreQuestions_loc[question-1],2),style=easyxf(border_left_medium)) 
        columnCounter+=1  
        sheetC.write(rowCounter,columnCounter,weightsQuestions_loc[question-1],style=easyxf(border_left_medium + align_horizcenter))  
        columnCounter+=1            
        rowCounter+=1

def write_scoreStudents(outputbook_loc,nameSheet_loc,permutations_loc,weightsQuestions_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,numAlternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    
    #deelnemersnummers
    sheetC.write(rowCounter, 0,"studentennummer", style=easyxf(style_header_borderRight)) 
    rowCounter+=1
    for i in xrange(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(style_header_borderRight))
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score",style=easyxf(style_header))
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i])
        rowCounter+=1
    columnCounter+=1;
    

    rowCounter = 0;
    #indicated series
    sheetC.write(rowCounter,columnCounter,"reeks",style=easyxf(style_header))
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
        sheetC.write(rowCounter,columnCounterHeader,"score vraag " + str(question),style=easyxf(style_header))
        columnCounterHeader+=1
        
    columnCounterScoreQuestions = columnCounter   
    rowCounter = 1;  
    for participant in xrange(len(totalScore_loc)): # loop over participants
        columnCounter = columnCounterScoreQuestions;
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        serie = int(columnSeries_loc[participant]-1)
        sorted_score = [score[i-1] for i in permutations_loc[serie]]
        #find questions with weight zero
        questionsZeroWeight = numpy.where(weightsQuestions_loc==0)
                #find questions
        #score[numpy.where(weightsQuestions_loc==0)[0]]=float('NaN')
        
        for question in xrange(1,numQuestions_loc+1):
            if (permutations_loc[serie][question-1]-1) in questionsZeroWeight:
                sheetC.write(rowCounter,columnCounter,"X")
            else:
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
            sheetC.write(rowCounter,columnCounter,"antwoord vraag " + str(question) + str(alternative),style=easyxf(style_header))
            #print "antwoord vraag " + str(question) + str(alternative)
            answer = matrixAnswers[:,questionAlternativeCounter]
            questionAlternativeCounter+=1
            rowCounter+=1
            #print 
            for i in xrange(len(totalScore_loc)):
                sheetC.write(rowCounter,columnCounter,answer[i])
                rowCounter+=1                    
            columnCounter+=1;    
            
def write_scoreStudentsNonPermutated(outputbook_loc,nameSheet_loc,permutations_loc,weightsQuestions_loc,numParticipants_loc,deelnemers_loc, numQuestions_loc,numAlternatives_loc,alternatives_loc,content_loc,content_colNrs_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,columnSeries_loc,matrixAnswers):
    sheetC = outputbook_loc.add_sheet(nameSheet_loc)

    columnCounter = 0;
    rowCounter = 0;
    
    #deelnemersnummers
    sheetC.write(rowCounter, 0,"studentennummer", style=easyxf(style_header_borderRight)) 
    rowCounter+=1
    for i in xrange(0,len(deelnemers_loc)):
        sheetC.write(rowCounter,columnCounter,deelnemers_loc[i], style=easyxf(style_header_borderRight))
        rowCounter+=1
    columnCounter+=1;
    
    rowCounter = 0;
    #total score for indicated series
    sheetC.write(rowCounter,columnCounter,"totale score",style=easyxf(style_header))
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,totalScore_loc[i])
        rowCounter+=1
    columnCounter+=1;
    

    rowCounter = 0;
    #indicated series
    sheetC.write(rowCounter,columnCounter,"reeks",style=easyxf(style_header))
    rowCounter+=1
    for i in xrange(len(totalScore_loc)):
        sheetC.write(rowCounter,columnCounter,columnSeries_loc[i])
        rowCounter+=1
    columnCounter+=1;
    
    #score for different questions
    rowCounter = 0;
    #write heading    
    columnCounterHeader = columnCounter
    for question in xrange(1,numQuestions_loc+1):
        sheetC.write(rowCounter,columnCounterHeader,"score vraag " + str(question),style=easyxf(style_header))
        columnCounterHeader+=1
        
    columnCounterScoreQuestions = columnCounter   
    rowCounter = 1;  
    for participant in xrange(len(totalScore_loc)): # loop over participants
        columnCounter = columnCounterScoreQuestions;
        score = scoreQuestionsIndicatedSeries_loc[participant,:]
        #serie = int(columnSeries_loc[participant]-1)
        #sorted_score = [score[i-1] for i in permutations_loc[serie]]
        #find questions with weight zero
        questionsZeroWeight = numpy.where(weightsQuestions_loc==0)
                #find questions
        #score[numpy.where(weightsQuestions_loc==0)[0]]=float('NaN')
        
        for question in xrange(1,numQuestions_loc+1):
            if (question-1) in questionsZeroWeight:
                sheetC.write(rowCounter,columnCounter,"X")
            else:
                sheetC.write(rowCounter,columnCounter,score[question-1])
            columnCounter+=1
        rowCounter+=1            
 

    #answers for different questions
    #beware answers are stored per question with the permutation; 
    #so for the us the answers have to be permutated 
    #write heading    
    for question in xrange(1,numQuestions_loc+1):
        for alternative in alternatives_loc:
            sheetC.write(0,columnCounterHeader,"antwoord vraag " + str(question) + alternative,style=easyxf(style_header))
            columnCounterHeader+=1
    rowCounter = 1;      
    columnCounterAnswers = columnCounter
    for participant in xrange(len(totalScore_loc)): # loop over participants
        columnCounter = columnCounterAnswers;
        serie = int(columnSeries_loc[participant]-1)
        answers =  matrixAnswers[participant]
        
        #find questions with weight zero
        questionsZeroWeight = numpy.where(weightsQuestions_loc==0)
                #find questions
        #score[numpy.where(weightsQuestions_loc==0)[0]]=float('NaN')
        
        for question in xrange(1,numQuestions_loc+1):
            questionInSerie = numpy.where(permutations_loc[serie]==question)[0][0]+1
            answersQuestion = answers[numAlternatives_loc*(questionInSerie-1):numAlternatives_loc*(questionInSerie-1)+numAlternatives_loc] 
            for counterAlternative in xrange(0,numAlternatives_loc):
                sheetC.write(rowCounter,columnCounter,answersQuestion[counterAlternative])
                columnCounter+=1
        rowCounter+=1                
        