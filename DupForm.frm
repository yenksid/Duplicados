VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DupForm 
   Caption         =   "Duplicates Report"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6435
   OleObjectBlob   =   "DupForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Delete_SaveB_Click()
Delete_Save
Unload Me
End Sub

Private Sub GenerateReportB_Click()

Dim RS As Workbook, NF As Workbook, ws As Worksheet

Set RS = ThisWorkbook

Application.ScreenUpdating = False
Application.DisplayAlerts = False


If FileBox.Value = vbNullString Then
    MsgBox "No File Selected"
    Exit Sub
End If

If IsFileOpen(FileBox.Value) Then
    MsgBox FileBox.Value & Chr(10) & Chr(10) & "File is Open, Please close the file and try again"
    Exit Sub
End If

Workbooks.Open filename:=FileBox.Value

Set NF = Workbooks(Workbooks.count)

NF.Activate

''Modded MsgBox in v2.2 to make the tool reject anything different than the VF file or the Duplicates Report
'If NF.Worksheets(1).Range("J1").Value2 <> "Data_Feed_Type__c" Then
'        NF.Close savechanges = False
'        'MsgBox "Invalid File Format, Please select a VivialForce format file"
'        MsgBox "Invalid File Format, Please select a VivialForce format file or a Duplicates Report"
'        Exit Sub
'End If

For Each ws In RS.Worksheets
    ws.Visible = xlSheetVisible
Next


DupForm.Hide
fname = NF.name

filtype = Left(fname, 18)

'Added validation to make the tool choose between a newly VF file or an already existing Duplicates Report
If filtype <> "Duplicates_Report_" Then
    
'    NF.Worksheets(1).Copy before:=RS.Worksheets(1)
          
    NF.Close savechanges = False
    
    RS.Activate
    
    For Each ws In Application.ActiveWorkbook.Worksheets
        If ws.name <> RS.Worksheets(1).name Then
            ws.Delete
        End If
    Next
    
    Set ws = RS.Worksheets(1)
    ws.name = "CATS_FILE"
    
    
    'fvs: To separate some columns from the thryv CAT format to simulate vivial force format to work this file.
    If RS.Worksheets("CATS_FILE").Range("AR2").Value <> vbNullString Then
        Thryv_CATS_Treatment
    End If


    'Where the magic happens
    Report_Build (fname)
    
    
    
    
    'added in v2.2 to accelerate the process when there's no data in the Near Dups and Caption Header Dups sheets, it will go straight fwd to the end.
    neardupscount = Worksheets("Near Dups").Range("B" & Rows.count).End(xlUp).row
    captionhcount = Worksheets("Caption Header Dups").Range("B" & Rows.count).End(xlUp).row
    
    'added in v2.2 to accelerate the process when there's no data in the Near Dups and Caption Header Dups sheets, it will go straight fwd to the end.
    If neardupscount = 1 And captionhcount = 1 Then
        
        Delete_Save
        RS.Worksheets("Dataset").Visible = xlSheetVeryHidden
        Unload Me
        
    Else
    'If the Near Dups and Caption Header Dups have data in them, this will create and save a Duplicates Report file.
        RS.Worksheets("CATS_FILE").Visible = -1
        
        RS.Sheets(Array("Dataset", "Automatic Removal Dups", "Near Dups", "Caption Header Dups", "CATS_FILE")).Copy
        
        Set NF = Workbooks(Workbooks.count)
        
        NF.Activate
        
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        
        fname = ThisWorkbook.Path & "\" & "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2
        
        NF.Worksheets("Dataset").Visible = xlSheetVeryHidden
        NF.Worksheets("CATS_FILE").Visible = xlSheetVeryHidden
        'fvs: Replace added to remove thryv files extention .txt in order to convert the file to the .xlsx format.
        NF.SaveAs filename:=Replace(fname, ".txt", ""), FileFormat:=51
        MsgBox fname
        NF.Close savechanges = False
        
        RS.Worksheets("Dataset").Visible = 2
        RS.Worksheets("CATS_FILE").Visible = 2
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        
        MsgBox "Duplicates Report Save Completed"
        
        Unload Me
        
    End If

Else
    'added to load an existing Duplicates Report to the workbook
    Set NF = Workbooks(Workbooks.count)
    
    NF.Activate
    
    NF.Worksheets("Dataset").Visible = -1
    NF.Worksheets("CATS_FILE").Visible = -1
    
    RS.Activate
    
    'Replacing the existing Sheets with the ones in the report to be loaded.
    RS.Worksheets("CATS_FILE").Delete
    NF.Worksheets("CATS_FILE").Copy before:=RS.Worksheets("Dataset")
    
    RS.Worksheets("Dataset").Delete
    NF.Worksheets("Dataset").Copy After:=RS.Worksheets("CATS_FILE")
    
    RS.Worksheets("Automatic Removal Dups").Delete
    NF.Worksheets("Automatic Removal Dups").Copy After:=RS.Worksheets("Dataset")
    
    RS.Worksheets("Near Dups").Delete
    NF.Worksheets("Near Dups").Copy After:=RS.Worksheets("Automatic Removal Dups")
    
    RS.Worksheets("Caption Header Dups").Delete
    NF.Worksheets("Caption Header Dups").Copy After:=RS.Worksheets("Near Dups")
    
    NF.Close savechanges = False
    
    'Delete_Save
    
    Unload Me
    
    MsgBox "Duplicates Report Load Complete"
    
    RS.Worksheets("Dataset").Visible = 2
    RS.Worksheets("CATS_FILE").Visible = 2
    
End If
'RS.Worksheets("Dataset").Visible = xlSheetVeryHidden
'MsgBox "Dup Report Complete"

End Sub

Private Sub MultiPage1_Change()

End Sub

'Private Sub OpenButton_Click()
'Dim filename As Variant
'
'file_Name = Application.GetOpenFilename(Title:="Choose VivialForce File to Generate Duplicates Report", MultiSelect:=False)
'
'If file_Name <> False Then FileBox.Value = file_Name
'End Sub

'Get VF files
Public Sub OpenButton_Click()
    Dim ws As Worksheet, PR As Workbook, WK As Workbook, files As Variant
    Set WK = ThisWorkbook
    
    'Improve macro performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Get the files
    file_Name = Application.GetOpenFilename(Title:="Choose a file to add", MultiSelect:=False)

    If file_Name <> False Then
        
        'Update the counter
        Controls("count_filesLabel").Caption = count_filesLabel + 1
        
        'Files names
        Dim search_array_files_names As Range
        Set search_array_files_names = WK.Worksheets(1).Range("AS:AS")
        
        'Add file to the form
        DupForm.FileBox.Value = extract_file_name(file_Name)
        
        'Open the files
        Workbooks.Open filename:=file_Name
        Set PR = Workbooks(Workbooks.count)
        
'        WK.Activate
'
'        For Each WS In Application.ActiveWorkbook.Worksheets
'            If WS.name <> WK.Worksheets(1).name Then
'                WS.Delete
'            End If
'        Next
'
'        Set WS = WK.Worksheets(1)
'        WS.name = "CATS_FILE"
        
        'Validating I'ts the correct Thryv file
        If Mid(PR.Worksheets(1).Range("A1").Value2, 513, 1) <> vbNullString Then
            'thryv_file = True
'            SectionTextBox.Enabled = True
            
            'Take the last row from the two files
            lastrow = WK.Worksheets(1).Range("AR" & Rows.count).End(xlUp).row
            lastrow2 = PR.Worksheets(1).Range("A" & Rows.count).End(xlUp).row
            'MsgBox lastrow
            If lastrow > 2 Then

                'Match the name of the file before add it to the sheet
                matched_file = match_name_phone(DupForm.FileBox.Value, search_array_files_names)
                
                If matched_file <> 0 Then
                    MsgBox "This file was added previously:" & " " & DupForm.FileBox.Value & vbCrLf & "Please try with a new file..."
                    PR.Close savechanges = False
                    Exit Sub
                End If
        
                'Save file name
                lastrow = WK.Worksheets(1).Range("AR" & Rows.count).End(xlUp).row
                WK.Worksheets(1).Range("AS" & lastrow + 1).Value = DupForm.FileBox.Value
'                'Take section
'                section = Replace(Replace(Replace(Right(file_Name, Len(file_Name) - InStrRev(file_Name, "_")), ".", ""), ".txt", ""), ".xlsx", "")
'
'                If Len(section) = 8 Then
'                    WK.Worksheets("CATS_FILE").Range("N" & lastrow + 1).Value = section
'                Else
'                    'RS.Worksheets("CATS_FILE").Cells.Clear
'                    MsgBox "No Section found, add section to the name of the file and try again."
'                    PR.Close savechanges = False
'                    Exit Sub
'                End If
                
                'copy the others files to the working workbook
                PR.Worksheets(1).Range("A1:A" & lastrow2).Copy Destination:=WK.Worksheets(1).Range("AR" & lastrow + 2)
                MsgBox "File added : " & DupForm.FileBox.Value
            Else
'                'Take section
'                section = Replace(Replace(Replace(Right(file_Name, Len(file_Name) - InStrRev(file_Name, "_")), ".", ""), ".txt", ""), ".xlsx", "")
'
'                If Len(section) = 8 Then
'                    WK.Worksheets("CATS_FILE").Range("N1").Value = section
'                Else
'                    'RS.Worksheets("CATS_FILE").Cells.Clear
'                    MsgBox "No Section found, add section to the name of the file and try again."
'                    PR.Close savechanges = False
'                    Exit Sub
'                End If
                
                'copy the first file to the working workbook
                PR.Worksheets(1).Range("A:A").Copy Destination:=WK.Worksheets("CATS_FILE").Range("AR:AR")
                
                'Take the names
                WK.Worksheets(1).Range("AS1").Value = DupForm.FileBox.Value
                
                css_file = True
                
            End If
            
        ElseIf PR.Worksheets(1).Range("AD1").Value2 = "Name" Then
        
'            'Gray out DataFeedType fields untill a corret file is added
'            DupForm.ProviderTextBox.Enabled = False
'            DupForm.ProviderTextBox.BackColor = &H80000016
'
'            DupForm.DataFeedTypeComboBox.Enabled = False
'            DupForm.DataFeedTypeComboBox.BackColor = &H80000016
    
            'Add file to the form
            DupForm.FileBox.Value = file_Name
            
            'Get file names
            file_Name = extract_file_name(file_Name)
            
            'Open the files
            Workbooks.Open filename:=file_Name
            Set PR = Workbooks(Workbooks.count)
            
            'copy the others files to the working workbook
             PR.Worksheets(1).Range("A:AP").Copy Destination:=WK.Worksheets(1).Range("A:AP")
             
             css_file = False
        Else
            MsgBox "You Choose a wrong VF file. Please try again:"
            PR.Close savechanges = False
            Exit Sub
        End If
        
        'Close the VF File
        PR.Close savechanges = False
        
        'Show image
        'Form.Image1.Visible = True
    End If
    
    'Insert telco provider
    Call Telco_Provider
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

'Matching name and phone number between Vivial Force and Galley
Function match_name_phone(search_value_vivialforce As Variant, search_array_galley_all As Range) As Variant
    'If not error run the match
    If Not IsError(Application.Match(search_value_vivialforce, search_array_galley_all, 0)) Then
        match_name_phone = Application.Match(search_value_vivialforce, search_array_galley_all, 0)
    Else
        match_name_phone = 0
    End If
End Function

'Extract the name of the file from the path.
Function extract_file_name(file As Variant) As String
    If file <> False Then
        While InStr(1, file, "\") <> 0
            file = Right(file, Len(file) - InStr(1, file, "\"))
        Wend
        extract_file_name = file
    End If
End Function

Sub UserForm_Initialize()
    
    With DupForm.DataFeedTypeComboBox
    .AddItem "Annual (LOCAL)"
    .AddItem "Annual (CLEC)"
    .AddItem "EAS"
    End With
End Sub

