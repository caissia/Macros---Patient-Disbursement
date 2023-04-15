Attribute VB_Name = "Setups"
Option Explicit
Global Patients()
'Contains 2 macros: Backup, Setup

Sub Backup()
'sx, saves a backup
'before closing workbook

    Dim FilePath, Folder, Year

    'exit if backup
    If Left(ThisWorkbook.Name, 6) = "Backup" Then Exit Sub

    'confirm save
    a = MsgBox(Space(6) & "This workbook will be saved and a backup created." & Chr(10) & Chr(10) & _
               Space(41) & "Proceed?", 4, "Confirm")
    If a <> 6 Then Exit Sub

    'acquire desktop directory
    a = ActiveWorkbook.Name
    Year = Mid(a, InStr(a, ".") - 4, 4)
    Folder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    Folder = Folder & FilePath & "\Subsidy - Backup\" & Year
    FilePath = Folder & "\Backup - " & ThisWorkbook.Name

    'create folder if needed
    If Len(Dir(Folder, vbDirectory)) = 0 Then MkDir Folder

    'save backup and original
    On Error GoTo err
    ThisWorkbook.SaveCopyAs FilePath
    ThisWorkbook.Save
    On Error GoTo 0

    Exit Sub

err:

    MsgBox Space(12) & "There is an unexpected error with the backup." & Chr(10) & Chr(10) & _
           Space(6) & "Check the path: Documents\Subsidy - Backup\" & Year & Chr(10) & Chr(10) & _
           Space(36) & "Then try again.", 0, "Backup Error"

End Sub

Sub Setup()
'sx, setups up workbook on open

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'empty patient array
    ReDim Patients(2000, 5)

    'allows macros without disabling protection
    For a = 13 To 1 Step -1
        With Sheets(a)
            .Activate
            .Protect , UserInterfaceOnly:=True
             Application.DisplayFormulaBar = False
             ActiveWindow.DisplayGridlines = False
             ActiveWindow.DisplayHeadings = False
             Application.DisplayStatusBar = False
             Application.Goto Sheets(a).Range("A1"), True

             If a < 13 Then
                Call Entry
                For Each b In Range("C5:C104")
                    If b = "" Then Exit For
                    If b <> "" Then
                        c = c + 1
                        'setup patient array
                        Patients(c, 1) = b.Offset(0, 0)
                        Patients(c, 2) = b.Offset(0, 1)
                        Patients(c, 3) = b.Offset(0, 2)
                        Patients(c, 4) = b.Offset(0, 3)
                        Patients(c, 5) = b.Offset(0, 4)
                    End If
                Next
             End If

             If a = 13 Then Range("W5").Select

        End With

    Next

    'select current month sheet
    Sheets(Month(Date)).Activate

    'maximize excel window
    Application.WindowState = xlMaximized
    ActiveWindow.WindowState = xlMaximized

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
