Attribute VB_Name = "Swtch"
Option Explicit

Sub Switch()

    'check if data exists else exit
    If Range("T5") = "" Then Exit Sub

    a = MsgBox(Space(8) & "Do you want to save a copy of the workbook or only the Annual Summary?" & Chr(10) & _
                Space(31) & "The copy is an image format saved to the desktop." & Chr(10) & Chr(10) & _
                Space(43) & "[  yes = workbook   |   no = summary  ]", 3, "Save a Copy")

    If a = 6 Then Call SaveBook
    If a = 7 Then Call SaveCopy

End Sub
