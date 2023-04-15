Attribute VB_Name = "SaveMth"
Option Explicit

Sub SaveCopy()
'sx, saves current sheet as
'an image in new workbook

    Dim a, b, c, d
    Dim OrigWb, OrigSh, NewBk
    Dim FilePath, Named, Title

    'check if data exists else exit
    If Range("C5") = "" Then Call Entry: Exit Sub
    If ActiveSheet.Index = 13 And Range("T5") = "" Then Exit Sub

    'if summary sheet called modify macro
    If ActiveSheet.Index = 13 Then
        a = "Annual Summary"
        c = True
        d = "summary"
    Else
        a = "month of " & MonthName(ActiveSheet.Index)
        d = "month"
    End If

    'confirm if copy required
    a = MsgBox(Space(6) & "A copy of the " & a & " will be created on the desktop." & Chr(10) & Chr(10) & Space(56) & "Proceed?", 4, "Create Image Copy")
    If a <> 6 Then
        If Not c Then Call Entry
        Exit Sub
    End If

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'acquire last row of patients
    If Not c Then b = Application.Index(Range("B5:B104"), _
                      Application.Match(Application.Max(Range("B5:B104")), _
                      Range("B5:B104"), 0)).Row

    'if range is above cut-off for chart then extend to chart
    If b < 24 Then b = 24
    
    'acquire copy range
    a = "A1:R" & b
    If c Then a = "A1:V50"

    'acquire this workbook data
    OrigWb = ThisWorkbook.Name
    OrigSh = ActiveSheet.Name

    'create title of newbook
    Named = "Subsidy Report - " & OrigSh & " " & Left(Right(OrigWb, 9), 4)
    If c Then Named = "Subsidy Report - " & "Summary" & " " & Left(Right(OrigWb, 9), 4)
    Title = Named & ".xlsx"

    'acquire desktop directory
    FilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    FilePath = FilePath & "\" & Title

    'create the new workbook and save
    On Error Resume Next
    Set NewBk = Workbooks.Add
    With NewBk
        .Title = Title
        .Subject = Named
        .SaveAs Filename:=FilePath
    End With

    'format & rename sheet
    Sheets(1).Select
    Cells.Select
    Selection.Interior.ThemeColor = xlThemeColorDark2
    ActiveSheet.Name = OrigSh
    If c Then ActiveSheet.Name = "Summary"

    'copy/paste data as a picture
    Workbooks(OrigWb).Sheets(OrigSh).Range(a).Copy
    Workbooks(Title).Sheets(1).Select
    ActiveSheet.Pictures.Paste.Select
    Selection.ShapeRange.ScaleWidth 1.004, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 1.02, msoFalse, msoScaleFromTopLeft

    'setup new workbook
    With Workbooks(Title).Sheets(1)
        Application.DisplayFormulaBar = False
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        Application.DisplayStatusBar = False
        .Range("A1").Select
        .Protect
    End With

    'delete extra sheets, save and close
    Workbooks(Title).Sheets("Sheet2").Delete
    Workbooks(Title).Sheets("Sheet3").Delete
    Workbooks(Title).Save
    Workbooks(Title).Close

    'alert of successful completion
    MsgBox Space(6) & "A copy of this " & d & " was created on the desktop." & Chr(10) _
        & Space(11) & "Workbook Title:  " & Named & ".", 0, "Success"

    'select next entry
    If Not c Then Call Entry

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
