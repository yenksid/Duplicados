Attribute VB_Name = "Module4"
Sub ShowForm()
Attribute ShowForm.VB_Description = "Show User Form"
Attribute ShowForm.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' ShowForm Macro
' Show User Form
'
' Keyboard Shortcut: Ctrl+Shift+R
'
DupForm.Show

End Sub

Sub FindDuplicates()
    Dim ws As Worksheet, wsDupes As Worksheet
    Dim dataRng As Range, row As Range, cell As Range
    Dim dict As Object, key As String
    Dim count As Long

    Set ws = ThisWorkbook.Worksheets("CATS_FILE")

    ' Check if MLP_DUPES sheet exists; if not, create it
    On Error Resume Next
    Set wsDupes = ThisWorkbook.Worksheets("MLP_DUPES")
    If wsDupes Is Nothing Then
        Set wsDupes = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsDupes.name = "MLP_DUPES"
    End If
    On Error GoTo 0

    Set dataRng = ws.Range("A2", ws.Cells(ws.Rows.count, "AC").End(xlUp))
    Set dict = CreateObject("Scripting.Dictionary")

    ' Iterate through rows and create a dictionary to count duplicates
    For Each row In dataRng.Rows
        If row.Cells(20) <> "True" Then
            key = row.Cells(17) & "|" & row.Cells(29) & "|" & row.Cells(28) & "|" & row.Cells(27) & "|" & row.Cells(23) & "|" & row.Cells(20)
            If dict.exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict(key) = 1
            End If
        End If
    Next row

    ' Clear MLP_DUPES sheet
    wsDupes.Cells.ClearContents

    ' Add headers to MLP_DUPES sheet
    ws.Rows(1).Copy Destination:=wsDupes.Rows(1)

    ' Copy rows with 10 or more duplicates to MLP_DUPES sheet
    count = 2
    For Each row In dataRng.Rows
        If row.Cells(20) <> "True" Then
            key = row.Cells(17) & "|" & row.Cells(29) & "|" & row.Cells(28) & "|" & row.Cells(27) & "|" & row.Cells(23) & "|" & row.Cells(20)
            If dict(key) >= 10 Then
                row.Copy Destination:=wsDupes.Cells(count, 1)
                count = count + 1
            End If
        End If
    Next row

'    MsgBox "Duplicates with 10 or more occurrences have been copied to the MLP_DUPES sheet.", vbInformation
End Sub


