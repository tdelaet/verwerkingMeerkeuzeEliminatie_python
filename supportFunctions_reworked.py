# -*- coding: utf-8 -*-
"""
Created on Wed May 21 14:54:46 2014

@author: tdelaet
"""

import numpy

def round2(x):
    """Numpy rounds x.5 to nearest even integer. To emulate SAS/SPSS, 
    which both round x.5 *up* to nearest integer, use this function.
    """
#    y = x - numpy.floor(x)
#    for i in numpy.arange(0,len(y)):
#        if (0 < y[i] < 0.5):
#            x[i] = numpy.floor(x[i])
#        else:
#            x[i] = numpy.ceil(x[i])
    x = numpy.floor(numpy.round(x,6)+0.5)    
    return x
    
def common_elements(list1, list2):
    return list(set(list1) & set(list2))
        
#def all_indices(value, qlist):
#    indices = []
#    idx = -1
#    while True:
#        try:
#            idx = qlist.index(value, idx+1)
#            indices.append(idx)
#        except ValueError:
#            break
#    return indices

def giveContentColNrs(content_loc, sheet_loc):
    content_colNrs_loc = [0] * len(content_loc);
    firstRowValues_loc = sheet_loc.row_values(0)
    indexC_loc =0
    # check if all content is present
    for x in content_loc:
        try:
            content_colNrs_loc[indexC_loc] = firstRowValues_loc.index(x)
            indexC_loc += 1
        except ValueError:
            print ("the expected content " + x + " is not present in the selected sheet")
            break;
    return content_colNrs_loc
    
def checkForUniqueParticipants(particpants):
    setd = set([x for x in particpants if particpants.count(x) > 1])
    if setd:
        print ("Duplicate particpants found: " + str(setd))
        return False
    else:
        return True
    
def getMatrixAnswers(sheet_loc,contentBook_loc,correctAnswers_loc,permutations_loc,alternatives_loc,numParticipants_loc,columnSeries_loc,content_colNrs_loc,twoOptions_loc):
    # Get the matrix of answers of the students
    numQuestions_loc = len(correctAnswers_loc)
    numAlternatives_loc = len(alternatives_loc)
    answers_loc= numpy.array(range(numParticipants_loc*numQuestions_loc*numAlternatives_loc),dtype='a10').reshape(numParticipants_loc,numQuestions_loc*numAlternatives_loc)
   
    counterColumn = 0
    for question_loc in range(1,numQuestions_loc+1):
        for alternative_loc in alternatives_loc:
            name_question_serie1 = "Vraag" + str(question_loc) + alternative_loc
            colNr_loc = content_colNrs_loc[contentBook_loc.index(name_question_serie1)]
            columnAlternative_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)
            answers_loc[:,counterColumn] = columnAlternative_loc;
            counterColumn+=1
    return answers_loc
            
    
def calculateScoreAllPermutations(sheet_loc,contentBook_loc,numSeries_loc,correctAnswers_loc,permutations_loc,alternatives_loc,numParticipants_loc,columnSeries_loc,content_colNrs_loc,twoOptions_loc,badQuestions_loc):
    # Calculate the score for each permutation and for each question 
    #numSeries_loc = len(permutations_loc)
    numQuestions_loc = len(correctAnswers_loc)
    numAlternatives_loc = len(alternatives_loc)
    scoreQuestionsAllPermutations_loc= numpy.zeros((numSeries_loc,numParticipants_loc,numQuestions_loc))
#    numOnmogelijkQuestionsAlternatives_loc= numpy.zeros((numQuestions_loc*numAlternatives_loc))    
#    numMogelijkQuestionsAlternatives_loc= numpy.zeros((numQuestions_loc*numAlternatives_loc) )   
    matrixAnswersQuestions_loc= numpy.zeros((numParticipants_loc,numAlternatives_loc) )
    #Calculate score for all permutations
                 
    for question_loc in range(1,numQuestions_loc+1):
        #print "----------------------"
        #print "question " + str(question_loc)
        for permutation in range(1,numSeries_loc+1):
            #print "permutatie " + str(permutation)
            #print "number of question in the first permutation"
            #print numSeries_loc
            if numSeries_loc ==1:
                numQuestionPermutations_loc = permutations_loc[question_loc-1]
            else:
                numQuestionPermutations_loc = permutations_loc[permutation-1][question_loc-1]
            #print numQuestionPermutations_loc
            correctAnswer = correctAnswers_loc[numQuestionPermutations_loc-1]
            # the correct answer =  the answer for the number of the question in the first permutation
            # get the answers or the different alternatives and put them in a matrix (entries in matrix are the index of the answer in twoOptions)
            counter_alternative = 0;
            for alternative_loc in alternatives_loc:
                #print "----------------------"
                name_question_serie1 = "Vraag" + str(question_loc) + alternative_loc
                #print name_question_serie1
                # get the column with the answers to this question
                colNr_loc = content_colNrs_loc[contentBook_loc.index(name_question_serie1)]
                columnAlternative_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)
                for participant_loc in range(numParticipants_loc):
                    #print (participant_loc)
                    answerIndex = twoOptions_loc.index(columnAlternative_loc[participant_loc])
                    matrixAnswersQuestions_loc[participant_loc,counter_alternative] = answerIndex
                if  correctAnswer == alternative_loc:   
                    correctAnswerIndex = counter_alternative
                counter_alternative+=1
            #print correctAnswer
            #print correctAnswerIndex
            #check if question is in list of BAD_Questions
            if badQuestions_loc[numQuestionPermutations_loc-1]==1:
                scoreQuestionsAllPermutations_loc[permutation-1,:,numQuestionPermutations_loc-1] = 1.0
            else:    
                scoreQuestionsAllPermutations_loc[permutation-1,:,numQuestionPermutations_loc-1] = calculateScoreQuestions(matrixAnswersQuestions_loc,correctAnswerIndex)
    return scoreQuestionsAllPermutations_loc
#    
#def calculateScoreQuestions(matrixAnswersQuestions_loc, correctAnswerIndex_loc):
#    numAlternatives_loc = len(matrixAnswersQuestions_loc[0])
#    numParticipants_loc = len(matrixAnswersQuestions_loc[:,0])
#    scoresQuestion = numpy.zeros(numParticipants_loc)    
#    
#    for counter_alternative in xrange(0,numAlternatives_loc):
#        indicesOnmogelijk_loc = [x for x in xrange(numParticipants_loc) if matrixAnswersQuestions_loc[x,counter_alternative]==0] # onmogelijk heeft index 0 in twoOptions
#        if  counter_alternative == correctAnswerIndex_loc:         # if the alternative is the correct answer => wrongly excluded so -1
#            scoresQuestion[indicesOnmogelijk_loc]-=1.0
#        else: # if the alternative is NOT the correct answer => correctly excluded so +1/(numAlternatives-1)
#            scoresQuestion[indicesOnmogelijk_loc]+= 1.0/(float(numAlternatives_loc)-1.0)  
#        counter_alternative+=1
#    return scoresQuestion

def calculateScoreQuestions(matrixAnswersQuestions_loc, correctAnswerIndex_loc):
    #print matrixAnswersQuestions_loc
    #print correctAnswerIndex_loc
    numAlternatives_loc = len(matrixAnswersQuestions_loc[0])
    numParticipants_loc = len(matrixAnswersQuestions_loc[:,0])
    scoresQuestion = numpy.zeros(numParticipants_loc)    
    
    scoreWrongAnswer_ET_rummens = -1.0/(numAlternatives_loc-1)
    
    # scores for y correctly eliminated answers
    scoreCorrectAnswer_ET_rummensyEL = [(1.0*y)/((numAlternatives_loc-y)*(numAlternatives_loc-1)) for y in numpy.arange(0,numAlternatives_loc)]
    
    # find students that indicated correct answer as impossible
    indicesCorrectAnswerEliminated_loc = [x for x in range(numParticipants_loc) if matrixAnswersQuestions_loc[x,correctAnswerIndex_loc]==0]
    scoresQuestion[indicesCorrectAnswerEliminated_loc] = scoreWrongAnswer_ET_rummens
    
    # the remaining questions where the correct answer is not indicated as impossible    
    remaining=  [x for x in range(numParticipants_loc) if matrixAnswersQuestions_loc[x,correctAnswerIndex_loc]!=0]
    # count number of correctly eliminated answers (y)
    numberCorrectlyEliminated = sum(numpy.concatenate((matrixAnswersQuestions_loc[:,0:correctAnswerIndex_loc],matrixAnswersQuestions_loc[:,correctAnswerIndex_loc+1:numAlternatives_loc]),axis=1).T ,0)*-1+numAlternatives_loc-1;
   
    for y in numpy.arange(0,numAlternatives_loc):
        indicesyCorrectlyEliminated = [x for x in remaining if numberCorrectlyEliminated[x]==y]
        scoresQuestion[indicesyCorrectlyEliminated] = scoreCorrectAnswer_ET_rummensyEL[y]
        
    return scoresQuestion
    
 
def getNumberMogelijkOnmogelijk(sheet_loc,content_loc,numSeries_loc,permutations_loc,columnSeries_loc,scoreQuestionsIndicatedSeries_loc,alternatives_loc,twoOptions_loc,content_colNrs_loc):
    numParticipants_loc = len(scoreQuestionsIndicatedSeries_loc)
    numQuestions_loc = len(scoreQuestionsIndicatedSeries_loc[0])
    numAlternatives_loc = len(alternatives_loc)
 
    numOnmogelijkQuestionsAlternatives_loc= numpy.zeros((numQuestions_loc*numAlternatives_loc))
    numMogelijkQuestionsAlternatives_loc= numpy.zeros((numQuestions_loc*numAlternatives_loc))
    
    #number of onmogelijk in upper group per question
    #loop over question
    for question_loc in range(1,numQuestions_loc+1):
#        print "----------------------"
#        print "question " + str(question_loc)
        counter_alternative = 0;
        for alternative_loc in alternatives_loc:
            name_question_serie1 = "Vraag" + str(question_loc) + alternative_loc
            #find column number in which question alternatives are given
            colNr_loc = content_colNrs_loc[content_loc.index(name_question_serie1)]
            #get the answers for the participants (so skip for row with name of first row)
            columnAlternative_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)            
            for permutation in range(1,numSeries_loc+1):
                indicesPermutation =  [x for x in range(len(columnSeries_loc)) if columnSeries_loc[x]==permutation]                
                if numSeries_loc ==1:
                    numQuestionPermutations_loc = permutations_loc[question_loc-1]              
                else:
                    numQuestionPermutations_loc = permutations_loc[permutation-1][question_loc-1]              
                counter_loc = (numQuestionPermutations_loc-1)*numAlternatives_loc+counter_alternative
                indicesOnmogelijk_loc = [x for x in indicesPermutation if columnAlternative_loc[x]==twoOptions_loc[0]]
                numOnmogelijkQuestionsAlternatives_loc[counter_loc]+=len(indicesOnmogelijk_loc)     
                indicesMogelijk_loc = [x for x in indicesPermutation if columnAlternative_loc[x]==twoOptions_loc[1]]                
                numMogelijkQuestionsAlternatives_loc[counter_loc]+=len(indicesMogelijk_loc)               
            counter_alternative+=1

    return numOnmogelijkQuestionsAlternatives_loc, numMogelijkQuestionsAlternatives_loc

def getScoreQuestionsIndicatedSeries(scoreQuestionsAllPermutations_loc,columnSeries_loc):
    #print "entered getScoreQuestionsIndicatedSeries"    
    numParticipants_loc = len(scoreQuestionsAllPermutations_loc[0])
    numQuestions_loc = len(scoreQuestionsAllPermutations_loc[0][0])           
    #print numParticipants_loc
    #print numQuestions_loc
    #print columnSeries_loc
    scoreQuestionsIndicatedSeries_loc= numpy.zeros((numParticipants_loc,numQuestions_loc))
    for participant in range(numParticipants_loc):
        serieIndicated = columnSeries_loc[participant]
        #print( serieIndicated)
        #print scoreQuestionsAllPermutations_loc[serieIndicated-1,participant,:] 
        scoreQuestionsIndicatedSeries_loc[participant,:] = scoreQuestionsAllPermutations_loc[int(serieIndicated)-1,participant,:] 
    averageScoreQuestions_loc = scoreQuestionsIndicatedSeries_loc.sum(axis=0)/float(numParticipants_loc)    
    return scoreQuestionsIndicatedSeries_loc, averageScoreQuestions_loc

def getOverallStatistics(scoreQuestionsIndicatedSeries_loc,maxTotalScore_loc,weightsQuestions_loc): 
    #print "entered getOverallStatistics" 
    numParticipants_loc = len(scoreQuestionsIndicatedSeries_loc)
    #numQuestions_loc = len(scoreQuestionsIndicatedSeries_loc[0])       
    #To calculate the total score only use the score for the series indicated by the student
    #totalScore_loc = scoreQuestionsIndicatedSeries_loc.sum(axis=1)/numQuestions_loc*maxTotalScore_loc
    totalScore_loc = (scoreQuestionsIndicatedSeries_loc*weightsQuestions_loc).sum(axis=1)/float(weightsQuestions_loc.sum(axis=0))*maxTotalScore_loc

    # set negative scores to 0
    totalScore_loc[totalScore_loc < 0]=0
    totalScore_nonRounded_loc = totalScore_loc
    #round the score     
    totalScore_loc = round2(totalScore_loc)
    #print totalScore
    averageScore_loc = sum(totalScore_loc)/float(numParticipants_loc)
    medianScore_loc = numpy.median(totalScore_loc)
    standardDeviation_loc = numpy.std(totalScore_loc)
    percentagePass_loc = 100*sum(score>= maxTotalScore_loc/2.0 for score in totalScore_loc)/float(numParticipants_loc)    
    return totalScore_loc, totalScore_nonRounded_loc, averageScore_loc, medianScore_loc, standardDeviation_loc, percentagePass_loc
    
def getOverallStatisticsDifferentSeries(totalScoreDifferentPermutations_loc,scoreQuestionsIndicatedSeries_loc, columnSeries_loc,maxTotalScore_loc):
    
    #print "entered getScoreQuestionsIndicatedSeries"    
    numParticipants_loc = len(totalScoreDifferentPermutations_loc)      
    numSeries_loc = len(totalScoreDifferentPermutations_loc[0]) 
    numParticipantsSeries_loc = numpy.zeros(numSeries_loc) 
    numQuestions_loc =  len(scoreQuestionsIndicatedSeries_loc[0])
    averageScore_loc = numpy.zeros(numSeries_loc)
    medianScore_loc = numpy.zeros(numSeries_loc)
    standardDeviation_loc = numpy.zeros(numSeries_loc)    
    percentagePass_loc = numpy.zeros(numSeries_loc)
    averageScoreQuestionsDifferentSeries_loc = numpy.zeros(numQuestions_loc* numSeries_loc)
    averageScoreQuestionsDifferentSeries_loc = averageScoreQuestionsDifferentSeries_loc.reshape(numQuestions_loc, numSeries_loc)
    
    for serie in range(1,numSeries_loc+1):
        indicesSerie_loc = [x for x in range(0,numParticipants_loc) if columnSeries_loc[x]==serie]
        totalScoreSerie_loc = [totalScoreDifferentPermutations_loc[i,serie-1] for i in indicesSerie_loc]
        numParticipantsSeries_loc[serie-1] = len(totalScoreSerie_loc)
        averageScore_loc[serie-1] = sum(totalScoreSerie_loc)/float(numParticipantsSeries_loc[serie-1])
        medianScore_loc[serie-1] = numpy.median(totalScoreSerie_loc)
        standardDeviation_loc[serie-1] = numpy.std(totalScoreSerie_loc)
        #print totalScoreSerie_loc
        percentagePass_loc[serie-1] = 100* sum(score>= maxTotalScore_loc/2.0 for score in totalScoreSerie_loc)/float(numParticipantsSeries_loc[serie-1]) 
        averageScoreQuestionsDifferentSeries_loc[:,serie-1] =  numpy.average(scoreQuestionsIndicatedSeries_loc[indicesSerie_loc,:],0)
    return numParticipantsSeries_loc, averageScore_loc, medianScore_loc, standardDeviation_loc, percentagePass_loc, averageScoreQuestionsDifferentSeries_loc

def calculateTotalScoreDifferentPermutations(scoreQuestionsAllPermutations_loc,maxTotalScore_loc,weightsQuestions_loc):
    numSeries_loc = len(scoreQuestionsAllPermutations_loc)
    numParticipants_loc = len(scoreQuestionsAllPermutations_loc[0])
    #numQuestions_loc = len(scoreQuestionsAllPermutations_loc[0][0])
    totalScorePermutations_loc = numpy.zeros((numParticipants_loc,numSeries_loc))    
    for serie in range(1,numSeries_loc+1):
        totalScore_temp = (scoreQuestionsAllPermutations_loc[serie-1]*weightsQuestions_loc).sum(axis=1)/float(weightsQuestions_loc.sum(axis=0) )*maxTotalScore_loc
        totalScore_temp[totalScore_temp < 0]=0
        totalScore_temp = round2(totalScore_temp)
        totalScorePermutations_loc[:,serie-1] = totalScore_temp
    return totalScorePermutations_loc

 
def calculateUpperLowerStatistics(sheet_loc,content_loc,numSeries_loc,columnSeries_loc,totalScore_loc,scoreQuestionsIndicatedSeries_loc,correctAnswers_loc,alternatives_loc,twoOptions_loc,content_colNrs_loc,permutations_loc):
    orderedDeelnemers_loc = sorted(range(len(totalScore_loc)),key=totalScore_loc.__getitem__) 
    numParticipants_loc = len(orderedDeelnemers_loc)
    numQuestions_loc = len(scoreQuestionsIndicatedSeries_loc[0])
    #print numQuestions_loc
    numAlternatives_loc = len(alternatives_loc)
    third_loc = int(numpy.ceil(numParticipants_loc/3.0))
    indicesUpper_loc = orderedDeelnemers_loc[numParticipants_loc-third_loc:numParticipants_loc]
    indicesLower_loc = orderedDeelnemers_loc[0:third_loc]
    indicesMiddle_loc= orderedDeelnemers_loc[third_loc:numParticipants_loc-third_loc]
    numUpper_loc = len(indicesUpper_loc)
    numLower_loc = len(indicesLower_loc)
    numMiddle_loc = len(indicesMiddle_loc)
    
    totalScoreUpper_loc = totalScore_loc[indicesUpper_loc]
    totalScoreMiddle_loc = totalScore_loc[indicesMiddle_loc]
    totalScoreLower_loc = totalScore_loc[indicesLower_loc]
    averageScoreUpper_loc = sum(totalScoreUpper_loc)/float(numUpper_loc)
    averageScoreMiddle_loc = sum(totalScoreMiddle_loc)/float(numMiddle_loc)
    averageScoreLower_loc = sum(totalScoreLower_loc)/float(numLower_loc)
    scoreQuestionsUpper_loc =  scoreQuestionsIndicatedSeries_loc[indicesUpper_loc,:]
    scoreQuestionsMiddle_loc =  scoreQuestionsIndicatedSeries_loc[indicesMiddle_loc,:]
    scoreQuestionsLower_loc =  scoreQuestionsIndicatedSeries_loc[indicesLower_loc,:]
    averageScoreQuestionsUpper_loc = scoreQuestionsUpper_loc.sum(axis=0)/float(numUpper_loc)
    averageScoreQuestionsMiddle_loc = scoreQuestionsMiddle_loc.sum(axis=0)/float(numMiddle_loc)
    averageScoreQuestionsLower_loc = scoreQuestionsLower_loc.sum(axis=0)/float(numLower_loc)
    
    numOnmogelijkQuestionsAlternativesUpper_loc= numpy.zeros(numQuestions_loc*numAlternatives_loc)
    numMogelijkQuestionsAlternativesUpper_loc= numpy.zeros(numQuestions_loc*numAlternatives_loc)
    numOnmogelijkQuestionsAlternativesMiddle_loc= numpy.zeros(numQuestions_loc*numAlternatives_loc)
    numMogelijkQuestionsAlternativesMiddle_loc= numpy.zeros(numQuestions_loc*numAlternatives_loc)
    numOnmogelijkQuestionsAlternativesLower_loc= numpy.zeros(numQuestions_loc*numAlternatives_loc)
    numMogelijkQuestionsAlternativesLower_loc= numpy.zeros(numQuestions_loc*numAlternatives_loc)
    
    
    #number of onmogelijk in upper group per question
    #loop over question
    for question_loc in range(1,numQuestions_loc+1):
#        print "----------------------"
#        print "question " + str(question_loc)
        counter_alternative = 0;
        for alternative_loc in alternatives_loc:
            name_question_serie1 = "Vraag" + str(question_loc) + alternative_loc
            #print "--------------------------------"
            #print name_question_serie1
            #find column number in which question alternatives are given
            colNr_loc = content_colNrs_loc[content_loc.index(name_question_serie1)]
            #get the answers for the participants (so skip for row with name of first row)
            columnAlternative_loc=sheet_loc.col_values(colNr_loc,1,numParticipants_loc+1)
                            
            for permutation in range(1,numSeries_loc+1):
                indicesPermutation =  [x for x in range(len(columnSeries_loc)) if columnSeries_loc[x]==permutation]                
                if numSeries_loc ==1:
                    numQuestionPermutations_loc = permutations_loc[question_loc-1]
                else:
                    numQuestionPermutations_loc = permutations_loc[permutation-1][question_loc-1]
                
                counter_loc = (numQuestionPermutations_loc-1)*numAlternatives_loc+counter_alternative

                indicesOnmogelijk = [x for x in indicesPermutation if columnAlternative_loc[x]==twoOptions_loc[0]]
                indicesOnmogelijkUpper_loc = [x for x in common_elements(indicesOnmogelijk,indicesUpper_loc)]           
                indicesOnmogelijkMiddle_loc = [x for x in common_elements(indicesOnmogelijk,indicesMiddle_loc)]
                indicesOnmogelijkLower_loc = [x for x in common_elements(indicesOnmogelijk,indicesLower_loc)]

                numOnmogelijkQuestionsAlternativesUpper_loc[counter_loc]+=len(indicesOnmogelijkUpper_loc)
                numOnmogelijkQuestionsAlternativesMiddle_loc[counter_loc]+=len(indicesOnmogelijkMiddle_loc)
                numOnmogelijkQuestionsAlternativesLower_loc[counter_loc]+=len(indicesOnmogelijkLower_loc)
            
                indicesMogelijk = [x for x in indicesPermutation if columnAlternative_loc[x]==twoOptions_loc[1]]
                indicesMogelijkUpper_loc = [x for x in common_elements(indicesMogelijk,indicesUpper_loc)]  
                indicesMogelijkMiddle_loc = [x for x in common_elements(indicesMogelijk,indicesMiddle_loc) ]  
                indicesMogelijkLower_loc = [x for x in common_elements(indicesMogelijk,indicesLower_loc)] 
                
                numMogelijkQuestionsAlternativesUpper_loc[counter_loc]+=len(indicesMogelijkUpper_loc)        
                numMogelijkQuestionsAlternativesMiddle_loc[counter_loc]+=len(indicesMogelijkMiddle_loc)        
                numMogelijkQuestionsAlternativesLower_loc[counter_loc]+=len(indicesMogelijkLower_loc)        
            counter_alternative+=1

    return totalScoreUpper_loc,totalScoreMiddle_loc,totalScoreLower_loc,averageScoreUpper_loc, averageScoreMiddle_loc, averageScoreLower_loc, averageScoreQuestionsUpper_loc, averageScoreQuestionsMiddle_loc, averageScoreQuestionsLower_loc, numOnmogelijkQuestionsAlternativesUpper_loc, numOnmogelijkQuestionsAlternativesMiddle_loc, numOnmogelijkQuestionsAlternativesLower_loc, numMogelijkQuestionsAlternativesUpper_loc, numMogelijkQuestionsAlternativesMiddle_loc, numMogelijkQuestionsAlternativesLower_loc, scoreQuestionsUpper_loc, scoreQuestionsMiddle_loc, scoreQuestionsLower_loc,numUpper_loc, numMiddle_loc, numLower_loc
    
def calculateVariances(totalScore_loc, scoreQuestionsIndicatedSeries_loc,numQuestions_loc):
    totalVariance = numpy.var(totalScore_loc)
    Variance = []
    for x in range(numQuestions_loc):
        nextVariance=numpy.var(scoreQuestionsIndicatedSeries_loc[:,x])
        Variance.append(nextVariance)
    return totalVariance, Variance

def calculateItemToetsCorrelatie(totalScore_loc, scoreQuestionsIndicatedSeries_loc,numQuestions_loc):
    Correlatie = []
    for x in range(numQuestions_loc):
        nextCovariance=numpy.cov(scoreQuestionsIndicatedSeries_loc[:,x],totalScore_loc)
        nextCorrelatie=nextCovariance[0][1]/(numpy.std(scoreQuestionsIndicatedSeries_loc[:,x])*numpy.std(totalScore_loc))
        Correlatie.append(nextCorrelatie)
    return Correlatie