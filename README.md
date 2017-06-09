# powerPointWordCloud
Creates WordCloud in PowerPoint using Text from Excel

This is a VBA-Script that should be used inside an Excel-File (.xlsm). It dynamically creates a word cloud on a new PowerPoint-Slid. The size of each word is calculated based on it's number of occurence.

This script uses the function QuickSortMultiDim by Dieter Otter (http://www.vbarchiv.net/tipps/tipp_1881.html)

# Pre-Requisites

Create a new excel file and assign the VBA-code to the first sheet or just download the complete Excel-File. If you are using your own file, your worksheet neeeds the following named ranges: 

Parameters
sourceText - Your wordlist, unsorted, just a text, nothing else
ppWidth - Width of the PowerPoint-Slide 
ppHeight - Height of the PowerPoint-Slide
fontSize - Initial size of the first (and therefore most frequent) word
factorHeight - A decimal number that will change the height of every word container 
factorPerformance - A decimal number that will speed up the process, the lower it is
topWords - Only process this amount of words

Feedback
speed - shows how fast the script was
wordList - a range with all the words and counting information
wordCount - column inside the Range "wordList" that shows the counter number of each word

Controller
startEngine - you can use this as a "button" to start the process (or just run it as you want), there is a trigger-sub called "Worksheet_SelectionChange" that starts the process when this range is selected

