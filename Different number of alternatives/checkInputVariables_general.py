# -*- coding: utf-8 -*-
"""
Created on Wed May 21 14:37:18 2014

@author: tdelaet
"""

from xlrd import open_workbook, biffh


def checkInputVariables(nameFile_loc,nameSheet_loc,numQuestions_loc,numAlternatives_loc,numSeries_loc,correctAnswers_loc,permutations_loc,weighsQuestions_loc,twoOptions_loc):
    return (
    checkFileAndSheet(nameFile_loc,nameSheet_loc) &
    checkCorrectAnswers(numQuestions_loc, numAlternatives_loc, correctAnswers_loc) & 
    checkPermutations(numSeries_loc,numQuestions_loc, permutations_loc, numAlternatives_loc) &
    checkWeightsQuestions(numQuestions_loc, weighsQuestions_loc) &
    checkTwoOptions(twoOptions_loc) 
    )
            
def checkFileAndSheet(nameFile_loc,nameSheet_loc):
    try:
        book = open_workbook(nameFile_loc)
        book.sheet_by_name(nameSheet_loc)
    except IOError:
        print "The selected file " + nameFile_loc +  " can not be opened as a workbook"
        return False
    except biffh.XLRDError:
        print "The selected sheet " + nameSheet_loc +  " can not be opened"
        return False
    return True;    
        
            
def checkCorrectAnswers(numQuestions_loc, numAlternatives_loc, correctAnswers_loc):
    check_OK=0#Check if there are errors with the correct answer input by setting to 1 if at least one error
    if (len(correctAnswers_loc) == numQuestions_loc):
        for x in range(1,len(numAlternatives_loc)+1):
            if not correctAnswers_loc[x-1] in (set(map(chr, range(65,65+numAlternatives_loc[x])))):
                print "ERROR: The correct answer for question " + str(x) + " (" + str(correctAnswers_loc[x-1]) + ") is outside the range " + str(map(chr, range(65,65+numAlternatives_loc[x])))#The correct answer for question 1, 2, 3, ... (\answer) is outside the range A,B,C, .. up to number of alternatives:
                check_OK=1
        if check_OK == 1:
            return False
    else:
        print "ERROR: The number of indicated questions (" + str(numQuestions_loc) +  ") is not equal to the number of correct answers listed (" + str(len(correctAnswers_loc)) + "): " + str(correctAnswers_loc)
        return False
    return True             
        
def checkPermutations(numSeries_loc,numQuestions_loc, permutations_loc, numAlternatives_loc):             
    check_OK=0#Check if there are errors in permutation list input by setting to 1 if at least one error
    # check that for all the permutations all questions are present  
    if (len(permutations_loc) == numSeries_loc):
        # check if all questions are present
        for permutationNumber_loc in xrange(1,numSeries_loc+1):
            if len(permutations_loc[permutationNumber_loc-1]) > numQuestions_loc:
                print "ERROR: The number of questions (" + str(numQuestions_loc) +  ") is less than the number of questions present in permutation " + str(permutationNumber_loc) + " ("+str(len(permutations_loc[permutationNumber_loc-1]))+") : " + str(permutations_loc[permutationNumber_loc-1])
                check_OK=1
            elif (set(xrange(1,numQuestions_loc+1)) != set(permutations_loc[permutationNumber_loc-1])):
                print "ERROR: Not all " + str(numQuestions_loc) +  " questions are present in permutation " + str(permutationNumber_loc) + ": " + str(permutations_loc[permutationNumber_loc-1])
                check_OK=1
            else:
                permutation_series_1=permutations_loc[0]
                for x in range(0,numQuestions_loc):
                    if numAlternatives_loc[permutations_loc[permutationNumber_loc-1][x]]!=numAlternatives_loc[permutation_series_1[x]]:
                        check_OK=1
                        print "ERROR: Permutation " + str(permutationNumber_loc) + " is not allowed, because it puts question " + str(permutations_loc[permutationNumber_loc-1][x]) + " in the same spot as question " + str(permutation_series_1[x])  + " in permutation 1 and they have a different number of alternatives."
        if check_OK==1:
            return False
    else:
        print "ERROR: The number of indicated series (" + str(numSeries_loc) +  ") is not equal to the number of permutations listed in the permutation list " + str(permutations_loc)
        return False
    return True     
 
def checkWeightsQuestions(numQuestions_loc, weightsQuestions_loc):
    if (len(weightsQuestions_loc) != numQuestions_loc):
        print "ERROR: The length of the questions weights (" + str(len(weightsQuestions_loc)) +  ") is not equal to the number of questions (" + str(numQuestions_loc)+")"
        return False
    return True     
    
def checkTwoOptions(twoOptions_loc ):
    if (len(twoOptions_loc) != 2):
        print "ERROR: The list of options"  + str(twoOptions_loc) + " does not contain two items"
        return False
    return True
    
