Attribute VB_Name = "ClearData"
Option Explicit
Global a, b, c
'Contains 2 macros: Clear, Reset

Sub Clear()
's1-s12, delete sheet data

    'check if data exists else exit
    If Range("C5") = "" Then
        Call Entry
        Exit Sub
    End If

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    a = MsgBox(Space(6) & "Do you want to delete all the data on this sheet?" & Chr(10) _
                & Chr(9) & Space(8) & "This cannot be undone.", 4, "Reset")

    If a = 6 Then Range("C5:N104").ClearContents

    Call Entry

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    ActiveSheet.Protect

End Sub

Sub Reset()
'Sum, delete all data

    'check if data exists else exit
    If Range("T5") = "" Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'confirm deletion
    a = MsgBox(Space(6) & "Do you want to delete ALL of the data for the ENTIRE year?" & Chr(10) _
                & Chr(9) & Space(18) & "This cannot be undone.", 4, "Complete Reset")
    'clear data
    If a = 6 Then
        For a = 12 To 1 Step -1
            Sheets(a).Range("C5:N104").ClearContents
            Application.Goto Sheets(a).Range("C5")
        Next
    End If

    'return to summary sheet
    Sheets("Sum").Activate

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    ActiveSheet.Protect

End Sub
