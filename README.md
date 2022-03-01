# Text_Analysing_Using_VBA-
## Descriptive quantitative free text analysis approach using Visual Basic for Applications

### Introduction
This VBA-Routine for Excel helps you to analyse text string and assigne them to categories.
Many questionaires use closed questions. But there are often also open-end questions. To analyse those open-end question is very important becaus of the new informations you will get. One possibility to analyse the content of an open-end text field ist to look for keywords and assign the content to category. So you can count assigned categories to get a feeling for impotance. Our tool will do this job for you. I will demonstrate the function of our analysing tool by a little example:

You are the owner of a hotel. With an online questionaire you ask your guests about the sprended time in your hotel. 
This Figure shows a part of this questionaire.

![Figure 1: Online Questionaire](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/online%20questinaire.png)

As you see, question 4 is an open-end question.
You can transfer the answers of this questionaire - and also the contents of the open-end text field of question number 4 - into an excel sheet.
Maybe this will look like this:

![Figure 2: Excel Transfer](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%201.png)

Now you can use our little tool to assigne the content of the text fields to diffenernt categories.
Please download the Excel file "TextAnalysis_v_1_0_with example". Open the file. Maybe you have to allow editing because of the VBA-routine inside the file.
Inside the file you will see four spreadsheats:
1. "START": Here you will find the button to start the analyse 

![Figure 3: Start Button](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%20button.png)

3. "free text": This is the spreadsheet you have to input your textfields (as you see in the example)

![Figure 4: free text spreadsheet](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%20category.png)

3. "category": Here you can declare your categories and assign contexts to the categories. You also can declare exception.
As you see in the spreadsheet there are categories like "cheese", "bread", "fruit" and so on. In the columns right of them you will the contexts. 
In the category "fruits" you will find following contexts: "kiwi", "banana" and so on. You will also find "raspberry !juice". The exclamation mark declared an exception. When the text field contains the word "raspberry" the text field will assigned the text field to the category "fruits". When the text field contains "raspberry juice" it will also be assigned to the category "fruits". With the exception "!juice" it will not. Maybe the content of the text filed contains "orange, orange juice" and you declared "orange !juce" in the category "juice" and "orange" in the category "fruit" the tool will detect both categories: "fruits" and "juice".  

![Figure 5: Category](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%20category.png)

4. "result": in this spreadsheet you will find the results af analysing. While analysing our tool counts the assignments of each category. The result of counting you will find here.

Information: You don't have to clear the spreadsheets before analysing. Our tool will clear the entries and the results every time you startet a new analasing.

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
