Attribute VB_Name = "SaveWb"
Option Explicit

Sub SaveBook()
'Sum, saves wb data
'as images in each sheet

    Dim a, b, c, d
    Dim NewSh, OrigWb
    Dim count, first, last
    Dim FilePath, Named, Title
    Dim ws As Worksheet, wb As Workbook

    'check if data exists else exit
    If Range("T5") = "" Then Exit Sub

    'select year or monthly range
    a = MsgBox(Space(6) & "Do you want to save a copy of a range of months or the entire year?" & Chr(10) & Chr(10) & _
           Space(42) & "[  yes = month  |  no = year  ]", 3, "Select Monthly Range or Year")
    If a = 2 Then Exit Sub

    'confirm if copy required, acquire monthly range to copy
    If a = 6 Then

        a = InputBox(Chr(10) & Space(16) & "Enter a number for the start month:" & Chr(10) & Chr(10) & _
                 Space(44) & "1     =     Jan" & Chr(10) & _
                 Space(44) & "2     =     Feb" & Chr(10) & _
                 Space(44) & "3     =     Mar . . . etc." & Chr(10), "Start Month", "Enter Here")
        If a = "" Or a > 12 Or Len(a) > 2 Or Not IsNumeric(a) Then Call Entry: Exit Sub

        b = InputBox(Chr(10) & Space(16) & "Enter a number for the last month:" & Chr(10) & Chr(10) & _
                 Space(44) & "1     =     Jan" & Chr(10) & _
                 Space(44) & "2     =     Feb" & Chr(10) & _
                 Space(44) & "3     =     Mar . . . etc." & Chr(10), "End Month", "Enter Here")
        If b = "" Or b > 12 Or Len(b) > 2 Or Not IsNumeric(b) Then Call Entry: Exit Sub

        first = a: first = first * 1
        last = b: last = last * 1
        
        If first > last Then Call Entry: Exit Sub

    End If

    If a = 7 Then

        a = MsgBox(Space(6) & "A copy of the entire year will be created on the desktop." & Chr(10) & Chr(10) & _
                   Space(48) & "Proceed?" & Chr(10) & Chr(10) & Space(12) & "[  Note:  Only months with data will be copied.  ]", 4, "Create Image Copy")
        If a <> 6 Then Call Entry: Exit Sub

    End If

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'setup var
    d = True

    'count sheets with data
    If first = "" Then
        For a = 1 To 12
            If Sheets(MonthName(a, True)).Range("C5") <> "" Then
                count = count + 1
                b = b & a & ","
                If count > 1 And d = True Then
                    If Split(b, ",")(count - 1) - Split(b, ",")(count - 2) = 1 Then _
                     d = True Else d = False
                End If
            End If
        Next
    End If

    'acquire original workbook name
    OrigWb = ThisWorkbook.Name

    'create name of new workbook
    If first = "" Then
        If count = 12 Then
            Named = "Annual " & Left(OrigWb, InStr(OrigWb, ".") - 1)

        ElseIf count > 1 And count < 12 And Split(b, ",")(0) = 1 And d Then
            Named = "Jan-" & MonthName(count, 1)

        ElseIf count = 1 Then
            Named = MonthName(Split(b, ",")(0), 1)

        ElseIf Not d Then
            For c = 1 To count
                Named = Named & MonthName(Split(b, ",")(c - 1), 1) & ","
            Next
            Named = Left(Named, Len(Named) - 1)

        Else
            Named = MonthName(Split(b, ",")(0), 1) & "-" & MonthName(Split(b, ",")(count - 1), 1)

        End If
    Else
        If first = last Then
            Named = MonthName(first, 1)

        Else
            Named = MonthName(first, 1) & "-" & MonthName(last, 1)

        End If
    End If

    'finalize title of new workbook
    GoSub NameWb
    Title = Named & ".xlsx"

    'acquire desktop directory
    FilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    FilePath = FilePath & "\" & Title

    'create copy of workbook
    Set wb = Workbooks.Add

    'copy all sheets to new wb
    For Each ws In ThisWorkbook.Sheets
        ws.Copy After:=wb.Sheets(wb.Sheets.count)
    Next ws

    'delete unneeded sheets from the new workbook
    If first = "" Then
        If count < 12 Then
            For a = 1 To 12
                If wb.Sheets(MonthName(a, 1)).Range("C5") = "" Then
                    wb.Sheets(MonthName(a, 1)).Delete
                End If
            Next
        End If
    Else
        For a = 1 To 12
            If a < first Or a > last Then
                wb.Sheets(MonthName(a, 1)).Delete
            End If
        Next
    End If
    wb.Sheets("Sheet1").Delete
    wb.Sheets("Sheet2").Delete
    wb.Sheets("Sheet3").Delete

    'save, close, reopen new wb - sum sheet can re-calculate
    wb.SaveAs FilePath

    'copy sheets as image
    For Each ws In wb.Sheets

        'acquire last row of patients
        For Each c In Sheets(ws.Name).Range("B5:B104")
            If c = "" Then b = c.Offset(-1).Row: Exit For
        Next

        'if range is above cut-off for chart then extend to chart
        If b < 24 Then b = 24

        'acquire copy range
        If ws.Name = "Sum" Then c = "A1:V50" Else c = "A1:R" & b

        'create temp name of new sheet
        NewSh = ws.Name & "-"

        'add, format, & name sheet
        With wb
            .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = NewSh
        End With

        'copy/paste data as image
        ws.Range(c).Copy
        Sheets(NewSh).Select
        ActiveSheet.Pictures.Paste.Select
        Selection.ShapeRange.ScaleWidth 1.004, msoFalse, msoScaleFromTopLeft
        Selection.ShapeRange.ScaleHeight 1.02, msoFalse, msoScaleFromTopLeft

    Next

    'format, rename, setup new sheets
    For Each ws In wb.Sheets
        If Right(ws.Name, 1) <> "-" Then
            wb.Sheets(ws.Name).Delete
        Else
            With Sheets(ws.Name)
                .Select
                .Cells.Select
                 Selection.Interior.ThemeColor = xlThemeColorDark2
                 Application.DisplayFormulaBar = False
                 ActiveWindow.DisplayGridlines = False
                 ActiveWindow.DisplayHeadings = False
                 Application.DisplayStatusBar = False
                 Application.Goto Range("A1")
                .Name = Left(ws.Name, 3)
                .Protect
            End With
        End If

    Next
    Sheets(1).Activate

    'save and close
    wb.Save
    wb.Close

    'alert of successful completion
    MsgBox Space(6) & "A copy of the workbook was created on the desktop." & Chr(10) _
        & Space(11) & "Workbook Title:  " & Named & ".", 0, "Success"

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

Exit Sub

NameWb:

    If count <> 12 Then
        a = InStr(OrigWb, ".")
        b = Mid(OrigWb, 1, a - 6)
        c = Mid(OrigWb, a - 4, 4)
        Named = b & " - " & Named & " " & c
    End If
    Return

End Sub
