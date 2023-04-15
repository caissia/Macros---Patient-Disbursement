Attribute VB_Name = "NewBook"
Option Explicit

Sub NwBook()
'Sum, creates a new workbook
'based on the current workbook

    Dim FilePath, Folder, Title, Year

    'exit if backup
    If Left(ThisWorkbook.Name, 6) = "Backup" Then Exit Sub

    'confirm new worksave
    a = MsgBox(Space(8) & "This will create a new workbook based on the current one." & Chr(10) & _
                Space(13) & "Note: this will not change or erase the old workbook." & Chr(10) & Chr(10) & _
                Space(50) & "Proceed?", 4, "Confirm")
    If a <> 6 Then Exit Sub

    'acquire new year for new workbook folder
    Year = InputBox("Please enter the new year for the new workbook.", "New Year", 2010)
    If Year = "" Then Exit Sub
    If Not IsNumeric(Year) Or Len(Year) <> 4 Or Year <= 2018 Then GoTo err

    'acquire new directory
    a = ThisWorkbook.Name
    Title = Left(a, InStr(a, ".") - 5) & Year & Right(a, Len(a) - InStr(a, ".") + 1)

    a = ActiveWorkbook.Path
    Folder = Left(a, InStrRev(a, "\") - 1) & "\" & Year

    FilePath = Folder & "\" & Title

    'create folder if needed
    If Len(Dir(Folder, vbDirectory)) = 0 Then MkDir Folder

    'save the new workbook
    On Error GoTo err
    ThisWorkbook.Save
    ThisWorkbook.SaveAs Filename:=FilePath
    On Error GoTo 0

    'exit if no data present
    If Range("T5") = "" Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'erase data if it exists
    For a = 12 To 1 Step -1
        Sheets(a).Range("C5:N104").ClearContents
        Application.Goto Sheets(a).Range("C5")
    Next

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    ActiveSheet.Protect

    Exit Sub

err:

    MsgBox Space(14) & "There is an error with the year entered." & Chr(10) & _
           Space(6) & "Creating a new workbook has been terminated.", 0, "Error"

End Sub
