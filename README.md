# Text_Analysing_Using_VBA-
## Descriptive quantitative free text analysis approach using Visual Basic for Applications

### Introduction
This VBA-Routine for Excel helps you to analyse text string and assigne them to categories.
Many questionaires use closed questions. But there are often also open-end questions. To analyse those open-end question is very important becaus of the new informations you will get. One possibility to analyse the content of an open-end text field ist to look for keywords and assign the content to category. So you can count assigned categories to get a feeling for impotance. Our tool will do this job for you. I will demonstrate the function of our analysing tool by a little example:

You are the owner of a hotel. With an online questionaire you ask your guests about the sprended time in your hotel. 
Figure 1 shows a part of this questionaire.
![Figure 1: Online Questionaire](https://myoctocat.com/assets/images/base-octocat.svg)

Preparing:
1. Create a new table.
2. Insert three spreadsheets
  a. spreadsheet named "free text"
  b. spreadsheet named "category"
  c. spreadsheet named "result"
3. In spreadsheet "free text" fill in the first column ("A") with text fields that will be analysed, in the second column ("B") you can insert a factor. Maybe you have text fields out of an questionaire and the content of two or more textfields are the same, you only have to fill in this text once an enter a factor. If you don't have a factor you have ton insert 1 (as number)
4. in spreadsheet "Category" fill in the first column ("A") the name of a category (e.g. "color"). In the columns right of column "A" you can insert contexts (e.g. "red", "blue", "yellow", "green"). Please only one context in a column. If one of the context was found in the textfield of spreasheet "free text", this text field will be assigned to this category.
5. You als can declare examinations. For example: If one text field contains the string "The employee was fired", the VBA-routine will find "red" in the word "fired". You can declare an examination by using an "!". Put after the context "red" the examination "!fired" (in one cell! "red !fired"). Of Course you can declare more than one examinations.
6. Copy and paste the VBA-routine in your Excel-table ("Developer Tools --> Visual basic") and start the routine
7. After finishing the routine, you will find the results of counting the assigned categories in the spreadshead "results".
