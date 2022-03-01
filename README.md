# Text_Analysing_Using_VBA-
## Descriptive quantitative free text analysis approach using Visual Basic for Applications

### Introduction
This VBA-Routine for Excel helps you to analyse text string and assigne them to categories.
Many questionaires use closed questions. But there are often also open-end questions. To analyse those open-end question is very important becaus of the new informations you will get. One possibility to analyse the content of an open-end text field ist to look for keywords and assign the content to category. So you can count assigned categories to get a feeling for impotance. Our tool will do this job for you. I will demonstrate the function of our analysing tool by a little example:


### Example

You are the owner of a hotel. With an online questionaire you ask your guests about the sprended time in your hotel. 
This Figure shows a part of this questionaire.

![Figure 1: Online Questionaire](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/online%20questinaire.png)

As you see, question 4 is an open-end question.
You can transfer the answers of this questionaire - and also the contents of the open-end text field of question number 4 - into an excel sheet.
Maybe this will look like this:

![Figure 2: Excel Transfer](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%201.png)

Now you can use our little tool to assigne the content of the text fields to diffenernt categories.
Please download the Excel file "TextAnalysis_using_VBA_v_1_0_with example.xlsm". Open the file. Maybe you have to allow editing because of the VBA-routine inside the file.
Inside the file you will see four spreadsheats:
1. "START": Here you will find the button to start the analyse 

![Figure 3: Start Button](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%20button.png)

2. "free text": This is the spreadsheet you have to input your textfields (as you see in the example). Some tools for questionaire will count free text field with the same content. For example, "lactose free yogurt" was entered by three differnet persons. Some tools will give you following output: "lactose free yogurt - 3". You can transfer these number into the column "counts". Our tool will note this. IMPORTANT: in all other cases you have to enter "1"!

![Figure 4: free text spreadsheet](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%20free%20text.png)

You will also find the VBA-routine in the files.


3. "category": Here you can declare your categories and assign contexts to the categories. You also can declare exception.
As you see in the spreadsheet there are categories like "cheese", "bread", "fruit" and so on. In the columns right of them you will the contexts. 
In the category "fruits" you will find following contexts: "kiwi", "banana" and so on. You will also find "raspberry !juice". The exclamation mark declared an exception. When the text field contains the word "raspberry" the text field will assigned the text field to the category "fruits". When the text field contains "raspberry juice" it would also be assigned to the category "fruits". With the exception "!juice" it will not. Maybe the content of the text filed contains "orange, orange juice" and you declared "orange !juce" in the category "juice" and "orange" in the category "fruit" the tool will detect both categories: "fruits" and "juice".  

![Figure 5: Category](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%20category.png)

4. "result": in this spreadsheet you will find the results af analysing. While analysing our tool counts the assignments of each category. The result of counting you will find here.

![Figure 6: Result](https://github.com/AndreasKlausen/Text_Analysing_Using_VBA-/blob/main/excel%20result.png)

Information: You don't have to clear the spreadsheets before analysing. Our tool will clear the entries and the results every time you startet a new analasing.

### Try Out!

You can download the file "TextAnalysis_using_VBA_v_1_0_with example.xlsm". Go to the first spreadsheet "START" and press the button. After a few seconds analysing is finished and the result will be shown. Go to the spreadsheet "free text". You will see new entries right of the colomn "counts". These are the assigned categories. Maybe this is very helpful for your own projects, because you can see if you have to add contexts or include exceptions.

For your own projects you can use the file with the example of course. Please overwrite our entries. But there is also a clear file without entries you can use: "TextAnalysis_using_VBA_v_1_0.xlsm". 

**Please note that they do not change the names of the spreadsheets and do not overwrite the headings in the spreadsheets!**
