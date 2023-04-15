Attribute VB_Name = "DataCopy"
Option Explicit
Global cTarget, cAddress
'Contains 4 macros: Duplicate, Entry, Scroll, Update

Sub Duplicate()
's1-s12, copies patient data if previous listed

    If Range(cAddress).Columns.count > 1 Then Exit Sub
    If Range(cAddress).Rows.count > 1 Then Exit Sub
    If cTarget = "" Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'if patient name entered
    If Range(cAddress).Column = 3 Then
    
        a = cTarget
        Set b = Range(cAddress)

        For Each c In Range("C5:C104")
            If c = a And c.Address <> cAddress Then
                Range(b.Offset(0, 1).Address & ":" & b.Offset(0, 4).Address).Value = Range(c.Offset(0, 1).Address & ":" & c.Offset(0, 4).Address).Value
                Range(cAddress).Offset(0, 5).Select
                c = True
                Exit For
            End If
        Next c
    
        If Not c Then
            For a = 1 To UBound(Patients)
                If LCase(cTarget) = LCase(Patients(a, 1)) Then
                    Range(b.Offset(0, 1).Address).Value = Patients(a, 2)
                    Range(b.Offset(0, 2).Address).Value = Patients(a, 3)
                    Range(b.Offset(0, 3).Address).Value = Patients(a, 4)
                    Range(b.Offset(0, 4).Address).Value = Patients(a, 5)
                    Range(cAddress).Offset(0, 5).Select
                    Exit For
                End If
            Next
        End If

    End If

    'if patient MR number entered
    If Range(cAddress).Column = 4 Then
    
        a = cTarget
        Set b = Range(cAddress)

        For Each c In Range("D5:D104")
            If c = a And c.Address <> cAddress Then
                Range(b.Offset(0, 1).Address & ":" & b.Offset(0, 3).Address).Value = Range(c.Offset(0, 1).Address & ":" & c.Offset(0, 3).Address).Value
                Range(b.Offset(0, -1).Address).Value = Range(c.Offset(0, -1).Address).Value
                Range(cAddress).Offset(0, 4).Select
                c = True
                Exit For
            End If
        Next c
    
        If Not c Then
            For a = 1 To UBound(Patients)
                If LCase(cTarget) = LCase(Patients(a, 2)) Then
                    Range(b.Offset(0, -1).Address).Value = Patients(a, 1)
                    Range(b.Offset(0, 1).Address).Value = Patients(a, 3)
                    Range(b.Offset(0, 2).Address).Value = Patients(a, 4)
                    Range(b.Offset(0, 3).Address).Value = Patients(a, 5)
                    Range(cAddress).Offset(0, 4).Select
                    Exit For
                End If
            Next
        End If

    End If

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    ActiveSheet.Protect

End Sub

Sub Entry()
's1-s12, goes to next entry in list

    For Each c In Range("C5:C104")
        If c = "" Then
            Range(c.Address).Select
            Exit For
        End If
    Next

End Sub

Sub Scroll()

    Application.ScreenUpdating = False
    Application.Goto Range("A1"), True

    If ActiveSheet.Index <= 12 Then
        c = Range("C5:C" & Rows.count).Cells.SpecialCells(xlCellTypeBlanks).Row
        Application.Goto Range("C" & c)
    Else
        Application.Goto Range("W5")
    End If

    Application.ScreenUpdating = True

End Sub

Sub Update()
's1-s12, populates previous
'patient data if match found

    Dim a, b, c, i

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'empty patient array
    ReDim Patients(2000, 5)
    
    'populate patient array
    For i = 1 To 12
        For Each c In Sheets(i).Range("C5:C104")
            If c = "" Then Exit For
            If c <> "" Then
                a = a + 1
                Patients(a, 1) = c.Offset(0, 0)
                Patients(a, 2) = c.Offset(0, 1)
                Patients(a, 3) = c.Offset(0, 2)
                Patients(a, 4) = c.Offset(0, 3)
                Patients(a, 5) = c.Offset(0, 4)
            End If
        Next
    Next

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
