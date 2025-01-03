Attribute VB_Name = "Module1"
Sub GenerateDocuments()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim TemplatePath As String
    Dim SavePath As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim rowData As Range
    Dim field As String
    Dim docName As String
    Dim i As Integer
    Dim cell As Range
    
    ' Define paths
    TemplatePath = "C:\Users\gerlie\Documents\VBA Projects\SampleTemplate.docx" ' Adjust path to your template
    SavePath = "C:\Users\gerlie\Documents\VBA Projects\" ' Path to save generated documents
    
    ' Open Word Application
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True ' For debugging; change to False when done
    
    ' Get data from Excel sheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the range for data
    Set rng = ws.Range("A2:D" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
    
    ' Loop through each row of data
    For i = 1 To rng.Rows.Count
        Set rowData = rng.Rows(i)
        Set WordDoc = WordApp.Documents.Open(TemplatePath) ' Open template document
        
        ' Replace placeholders with actual data
        For Each cell In rowData.Cells
            field = ws.Cells(1, cell.Column).Value ' Get the header value, i.e., field name
            
            ' Find and replace the placeholder in the Word document
            If Len(Trim(cell.Value)) > 0 Then ' If cell is not empty
                WordDoc.Content.Find.Execute _
                    FindText:="<" & field & ">", _
                    ReplaceWith:=cell.Value, _
                    Replace:=2 ' wdReplaceAll
            End If
        Next cell
        
        ' Generate and save the document with a unique name (using Name column)
        docName = rowData.Cells(1, 1).Value & "_Document.docx" ' Name based on the Name column (column A)
        WordDoc.SaveAs SavePath & docName
        WordDoc.Close False ' Close the document, not saving again
        
    Next i
    
    ' Quit Word
    WordApp.Quit
    Set WordApp = Nothing
    
    MsgBox "Documents generated successfully!", vbInformation
End Sub


