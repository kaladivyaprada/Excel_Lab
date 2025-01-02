Attribute VB_Name = "Module1"
Sub siddharameshwara()
    Dim FileName As String
    ' Update the path to a valid path on your PC
    FileName = VBA.FileSystem.Dir("C:\Users\PRACTICE\Documents\output.csv")
    
    If FileName = VBA.Constants.vbNullString Then
        MsgBox "File does not exist."
    Else
        ' Update the path to a valid path on your PC
        Workbooks.Open "C:\Users\PRACTICE\Documents\output.csv" & output
    End If
End Sub

Sub PathExists()
    Dim Path As String
    Dim Folder As String
    Dim Answer As VbMsgBoxResult
    ' Update the path to a valid path on your PC
    Path = "C:\Users\PRACTICE\Documents\output.csv"
    Folder = Dir(Path, vbDirectory)

    If Folder = vbNullString Then
        Answer = MsgBox("Path does not exist. Would you like to create it?", vbYesNo, "Create Path?")
        Select Case Answer
            Case vbYes
                VBA.FileSystem.MkDir ("C:\Users\PRACTICE\Documents\output.csv")
            Case Else
                Exit Sub
        End Select
    Else
        MsgBox "Folder exists."
    End If
End Sub

Sub Get_Data_From_File()
    ' Note: In the Regional Project that's coming up we learn how to import data from multiple Excel workbooks
    ' Also see BONUS sub procedure below (Bonus_Get_Data_From_File_InputBox()) that expands on this by including an input box
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
    Application.ScreenUpdating = False
    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.csv*),*.csv*")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        ' Copy data from A1 to E20 from first sheet
        OpenBook.Sheets(1).Range("A1:E20").Copy
        ThisWorkbook.Worksheets("output").Range("A10").PasteSpecial xlPasteValues
        OpenBook.Close False
    End If
    Application.ScreenUpdating = True
End Sub



