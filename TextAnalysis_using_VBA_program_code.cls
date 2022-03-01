VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    
    Dim Row, Col, ctxRow, ctxCol, delCol, delRow, numRows, t  As Integer
    Dim searchStr As String
    Dim searchCat As String
    Dim numCat, countCat As Integer
    Dim assigned As Boolean
    Dim searchRes As Integer
    Dim numHint As Integer
    Dim numExcepHint As Integer
    Dim numExcep As Integer
    Dim excepStr As String
    Dim posExcep(1 To 20) As Integer
    Dim pos As Integer
    
    
    '*** Searching the free text fields ***
    
    numRows = Worksheets("free text").UsedRange.Rows.Count   'Number of entries in the "Free text" table
    numCat = Worksheets("category").UsedRange.Rows.Count  'number of categorys
    
    '*** delating the table with the results ***
    
    For delCol = 1 To 25                'if there are more than 25 colums in the result table, please increase the number
        For delRow = 1 To 25
            Worksheets("result").Cells(delRow, delCol).Value = ""
        Next delRow
    Next delCol
    
    '*** deleting the entries of the categories in the free text field table
    
    For delCol = 3 To (numCat + 3) 'in the table, entries from column 3 onwards are deleted, max number of columns corresponds to the number of categories
        For delRow = 1 To numRows
            Worksheets("free text").Cells(delRow, delCol).Value = ""
        Next delRow
    Next delCol
    
        
        
    For countCat = 2 To numCat  'Start loop until all categories are processed
        searchRes = 0           'Number of search results
        searchCat = Worksheets("category").Cells(countCat, 1).Value                    'Current category to be edited
        numKeyw = (Worksheets("category").Cells(countCat, 256).End(xlToLeft).Column)   'Number of keywords in the category currently being edited
        For countKeyw = 2 To numKeyw    'Start loop until all keywords are processed
            searchStr = UCase(Worksheets("category").Cells(countCat, countKeyw).Value) 'current keyword is loaded in upper case into variable searchStr
            numExcep = 0
            If (InStr(searchStr, "!")) > 0 Then      'if the current keyword contains an exception (!) then
                searchStrExcep = searchStr
                For pos = 1 To Len(searchStrExcep)  'Check how often an exception occurs in the search string
                    If (InStr(pos, searchStrExcep, "!")) > 0 Then
                        numExcep = numExcep + 1
                        posExcep(numExcep) = (InStr(pos, searchStrExcep, "!"))   'Save the position of the exception
                        pos = InStr(pos, searchStrExcep, "!")                    'continue the loop with the found position
                    End If
                Next pos
                posExcep(numExcep + 1) = Len(searchStr) + 2      'Save as last position the next position to be assumed
                searchStr = Mid(searchStr, 1, posExcep(1) - 2)   'only take the part of the string up to the first exception as the search string
            End If
            For Row = 2 To numRows   'Start of loop in which table is searched for current keyword
                If numExcep > 0 Then 'If exceptions are found, first see which ones occur in the search string and add them up.
                    numExcepHint = 0
                    For t = 1 To numExcep
                        excepStr = Mid(searchStrExcep, posExcep(t) + 1, posExcep(t + 1) - posExcep(t) - 2)
                        If ((InStr(UCase(Worksheets("free text").Cells(Row, 1).Value), excepStr))) > 0 Then
                            numExcepHint = numExcepHint + 1 'if the exception string occurs in the search term increase the number of hits by 1
                        End If
                    Next t
                End If
                If ((InStr(UCase(Worksheets("free text").Cells(Row, 1).Value), searchStr))) > 0 Then 'if current table content contains the search string, then
                    numHint = 0
                    For pos = 1 To Len(Worksheets("free text").Cells(Row, 1).Value)  'Check how often the search string occurs in the current table content.
                        If ((InStr(pos, UCase(Worksheets("free text").Cells(Row, 1).Value), searchStr))) > 0 Then
                            numHint = numHint + 1
                            pos = InStr(pos, UCase(Worksheets("free text").Cells(Row, 1).Value), searchStr)
                        End If
                    Next pos
                    assigned = False
                    For Col = 3 To (Worksheets("free text").Cells(Row, 256).End(xlToLeft).Column)         'See if the detected category has already been assigned
                        If Worksheets("free text").Cells(Row, Col).Value = searchCat Then assigned = True 'Category has already been assigned
                    Next Col
                    If assigned = False And numHint > numExcepHint Then 'No entry of the current category so far and search word occurs next to the exceptions in the search string
                            Worksheets("free text").Cells(Row, Col).Value = searchCat  'if category is not assigned, then enter category in the first free column
                            searchRes = searchRes + (Worksheets("free text").Cells(Row, 2).Value) 'Increase number of search results by the number of frequencies of the entries
                    End If
                End If
            Next Row
        Next countKeyw
        Worksheets("result").Cells(1, 1).Value = "free text"
        Worksheets("result").Cells(2, countCat - 1).Value = searchCat 'The current category is added to the Overview table
        Worksheets("result").Cells(3, countCat - 1).Value = searchRes 'and the current number of entries found is entered
    Next countCat
    
    
    
  '*** Displaying the "Result" worksheet ***

  Worksheets("result").Activate
End Sub
