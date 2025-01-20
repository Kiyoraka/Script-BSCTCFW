Sub ChangeNonWhiteNonRedShadedTableColorInBatch()
    Dim folderPath As String
    Dim file As String
    Dim doc As Document
    Dim tbl As Table
    Dim cell As cell
    
    ' Specify the folder containing your Word documents
    folderPath = "C:\FOLDER PATH 1\FOLDER PATH 2" ' Change this to your folder path
    
    ' Ensure the folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Get the first file in the folder
    file = Dir(folderPath & "*.docx")
    
    ' Loop through all Word documents in the folder
    While file <> ""
        ' Open the document
        Set doc = Documents.Open(folderPath & file)
        
        ' Loop through all tables in the document
        For Each tbl In doc.Tables
            ' Loop through each cell in the table
            For Each cell In tbl.Range.Cells
                ' Check the shading color of the cell
                If cell.Shading.BackgroundPatternColor <> wdColorAutomatic And _
                   cell.Shading.BackgroundPatternColor <> RGB(255, 255, 255) Then
                    ' Change the shading color to red
                    cell.Shading.BackgroundPatternColor = wdColorBlue
                End If
            Next cell
        Next tbl
        
        ' Save and close the document
        doc.Save
        doc.Close
        
        ' Get the next file
        file = Dir
    Wend
    
    MsgBox "Shading color changed for all non-white shaded cells in the folder!", vbInformation
End Sub


