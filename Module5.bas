Attribute VB_Name = "Module5"
'Extract the data of frontier EAS files in columns
Public Sub Thryv_CATS_Treatment()
    Dim WK As Workbook, arr As Variant, class_of_service, phone, section As String
    Set WK = ThisWorkbook
    
    Dim caption_header As Boolean
    Dim i As Long
    
    Dim header_name As Variant
    header_name = Array("Action__c", "BEX__c", "BOID__c", "Bus_Res_Gov_Indicator__c", "Caption_Display_Text__c", "Caption_Header__c", "Caption_Member__c", "CLEC_Provider__c", "Cross_Reference_Text__c", "Data_Feed_Type__c", "Designation__c", "Directory__c", "Directory_Heading__c", "Directory_Section__c", "Disconnect_Reason__c", "Effective_Date__c", "First_Name__c", "Honorary_Title__c", "Indent_Level__c", "Indent_Order__c", "Left_Telephone_Phrase__c", "Lineage_Title__c", "Listing_City__c", "Listing_Country__c", "Listing_PO_Box__c", "Listing_Postal_Code__c", "Listing_State__c", "Listing_Street__c", "Listing_Street_Number__c", "Name", "Phone__c", "Phone_Override__c", "Phone_Type__c", "Right_Aligned_Phrase__c", "Secondary_Surname__c", "Service_Order__c", "Telco_Provider__c", "Title__c", "Under_Caption__c", "Under_Sub_Caption__c", "Year__c", "Manual_Sort_As_override")
    
    'Asigning the data to an array
    arr = WK.Worksheets("CATS_FILE").Range("AR1").CurrentRegion

    'From the first item in the array to the last
    For i = LBound(arr, 1) To UBound(arr, 1)
        
        'Indent level
        indent_level = Mid(arr(i, 1), 55, 1)
        WK.Worksheets("CATS_FILE").Range("S" & i).Value = indent_level
        'MsgBox indent_level
        On Error Resume Next
        'Captions Header
        'NEWSTART It's not flagging caption headers, DLS does.
        If Mid(arr(i, 1), 55, 1) = 0 And Mid(arr(i + 1, 1), 55, 1) <> 0 Then
            caption_header = True
            
            'To know if its a caption we need to review the line below each line
            'At the end of the file the next line is empty, so we need to identify
            'this line so that the macro can work properly
            If Mid(arr(i + 1, 1), 55, 1) = vbnllstring Then
                caption_header = False
            End If
        Else
            caption_header = False
        End If
        'Finish the program
        On Error GoTo 0
        
        'Indent Order
        If indent_level = 0 And caption_header = False Then
            WK.Worksheets("CATS_FILE").Range("T" & i).Value = 0
            
        ElseIf caption_header = True Then
            WK.Worksheets("CATS_FILE").Range("T" & i).Value = 10
            
        ElseIf indent_level <> 0 And caption_header = False Then
            WK.Worksheets("CATS_FILE").Range("T" & i).Value = WK.Worksheets("CATS_FILE").Range("T" & i - 1).Value + 10
        End If
        
        'Class Of Service
        'captions headers do not have the type so we take the next type and added it to the caption header.
        If Mid(arr(i, 1), 249, 1) <> " " Then
            class_of_service = Mid(arr(i, 1), 249, 1)
            WK.Worksheets("CATS_FILE").Range("D" & i).Value = class_of_service
        Else
            'Take next type
            class_of_service = Mid(arr(i + 1, 1), 249, 1)
            WK.Worksheets("CATS_FILE").Range("D" & i).Value = class_of_service
        End If
        
        '//If Caption Header Is Falso Do Things Below//
        '//Missing toll free//
        '//3 digits number//
        '//Telco code, EAS, CLEC, LOCAL//
                
        'Street number
        Streets_Numbers = Mid(arr(i, 1), 260, 32)
        WK.Worksheets("CATS_FILE").Range("AC" & i).Value = Trim(Streets_Numbers)
        
        'Cardinals
        'DLS & NEWSTART have it in different positions
        cardinals = Mid(arr(i, 1), 362, 15)
        cardinals = Trim(Replace(cardinals, " ", ""))
        
        'Street name
        Streets_Names = Mid(arr(i, 1), 292, 70)
        WK.Worksheets("CATS_FILE").Range("AB" & i).Value = Trim(cardinals & " " & Streets_Names)
        
        'Comunity
        Communitys = Mid(arr(i, 1), 377, 45)
        WK.Worksheets("CATS_FILE").Range("W" & i).Value = Trim(Communitys)
        
        'State
        state_code = Mid(arr(i, 1), 422, 18)
        WK.Worksheets("CATS_FILE").Range("AA" & i).Value = Trim(state_code)
        
        'Postal Code
        postal_code = Mid(arr(i, 1), 440, 13)
        WK.Worksheets("CATS_FILE").Range("AA" & i).Value = Trim(state_code)
        
        'Phone
        phone = Mid(arr(i, 1), 453, 20)
        WK.Worksheets("CATS_FILE").Range("AE" & i).Value = Trim(Replace(phone, " ", ""))
        
        'Name
        caption_name = Mid(arr(i, 1), 513, 100)
        
'        If caption_header = True Then
'            caption_name = Trim(Replace(caption_name, "|", ""))
'            WK.Worksheets("CATS").Range("AD" & i).Value = Trim(caption_name)
'        ElseIf indent_level = 0 And class_of_service = "R" Then
'            last_name = Left(caption_name, InStr(1, caption_name, "|") - 1)
'            WK.Worksheets("CATS").Range("AD" & i).Value = Trim(last_name)
'            'MsgBox last_name
'        End If
        
        '//Verify if we have differents class service or listings types//
        
        'Residential listings
        If indent_level = 0 And class_of_service = "R" And caption_header = False Then
        
            'Last name
            last_name = Left(caption_name, InStr(1, caption_name, "|") - 1)
            WK.Worksheets("CATS_FILE").Range("AD" & i).Value = Trim(last_name)
            
            'First name
            first_name = Right(caption_name, Len(caption_name) - Len(last_name) - 2)
            WK.Worksheets("CATS_FILE").Range("Q" & i).Value = Trim(first_name)
            'Business listings
        ElseIf indent_level = 0 And class_of_service = "B" And caption_header = False Then
            'Manage cross reference listings
            '//The macro is leaving an empty line were it find a cross refence//
            If Left(caption_name, 3) = "See" Then
                WK.Worksheets("CATS_FILE").Range("I" & i - 1).Value = Trim(caption_name)
            Else
                caption_name = Trim(Replace(caption_name, "|", ""))
                WK.Worksheets("CATS_FILE").Range("AD" & i).Value = Trim(caption_name)
            End If
            'Caption header business
        ElseIf caption_header = True And class_of_service = "B" Then
            WK.Worksheets("CATS_FILE").Range("AD" & i).Value = Trim(caption_name)
            'Caption header residentials
        ElseIf caption_header = True And class_of_service = "R" Then
            
            'Last name
            last_name = Left(caption_name, InStr(1, caption_name, " ") - 1)
            WK.Worksheets("CATS_FILE").Range("AD" & i).Value = Trim(last_name)
            
            'First name
            first_name = Right(caption_name, Len(caption_name) - Len(last_name) - 1)
            WK.Worksheets("CATS_FILE").Range("Q" & i).Value = Trim(first_name)
            'Caption display text
        ElseIf indent_level <> 0 Then
             WK.Worksheets("CATS_FILE").Range("E" & i).Value = Trim(caption_name)
        End If
  
        'Toll free indicator
'        If Toll_Free_Indicator(i) = True Then WK.Worksheets("CATS").Range("AG" & i).Value = "Toll Free"
        
        'Cross reference listings
'        If indent_level = 0 And Left(caption_name, 3) = "See" Then
'            WK.Worksheets("CATS").Range("I" & i).Value = Trim(caption_name)
'        End If
        
        'Section
'        section = DupForm.SectionTextBox.Value
'        WK.Worksheets("CATS_FILE").Range("N" & i).Value = section
        
        'Section
        directory_number = Left(section, 6)
        WK.Worksheets("CATS_FILE").Range("L" & i).Value = directory_number
        
'        'If its empty the dup report remove all lines
'
'        'Clec Provider
'        If DupForm.DataFeedTypeComboBox = "Annual (CLEC)" Then
'            WK.Worksheets("CATS_FILE").Range("H" & i).Value = DupForm.ProviderTextBox
'        End If
'
'        'Data Feed Type
'        If DupForm.DataFeedTypeComboBox = "Annual (LOCAL)" Then
'            WK.Worksheets("CATS_FILE").Range("J" & i).Value = "Annual"
'        ElseIf DupForm.DataFeedTypeComboBox = "Annual (CLEC)" Then
'            WK.Worksheets("CATS_FILE").Range("J" & i).Value = "Annual"
'        Else
'            WK.Worksheets("CATS_FILE").Range("J" & i).Value = "EAS"
'        End If
'
'        'Telco Provider
'        If DupForm.DataFeedTypeComboBox <> "Annual (CLEC)" Then
'            WK.Worksheets("CATS_FILE").Range("AK" & i).Value = DupForm.ProviderTextBox
'        End If
        
    Next i
    'Insert row to provide the header
    WK.Worksheets("CATS_FILE").Rows(1).Insert
    WK.Worksheets("CATS_FILE").Range("A1:AP1").Value = header_name
    
    'This is where we are saving the name of the file:
    WK.Worksheets(1).Columns("BA").Delete
End Sub

'Extract the data of frontier EAS files in columns
Public Sub Telco_Provider()
    Dim WK As Workbook, arr As Variant, class_of_service, phone, section As String
    Set WK = ThisWorkbook
    
    'Asigning the data to an array
    lastrow = WK.Worksheets("CATS_FILE").Range("AR" & Rows.count).End(xlUp).row
    
    If WK.Worksheets("CATS_FILE").Range("J1").Value = vbNullString Then
        For i = 1 To lastrow Step 1
        
            'If its empty the dup report remove all lines
            'Clec Provider
            If DupForm.DataFeedTypeComboBox = "Annual (CLEC)" Then
                WK.Worksheets("CATS_FILE").Range("H" & i).Value = DupForm.ProviderTextBox
            End If
            
            'Data Feed Type
            If DupForm.DataFeedTypeComboBox = "Annual (LOCAL)" Then
                WK.Worksheets("CATS_FILE").Range("J" & i).Value = "Annual"
            ElseIf DupForm.DataFeedTypeComboBox = "Annual (CLEC)" Then
                WK.Worksheets("CATS_FILE").Range("J" & i).Value = "Annual"
            Else
                WK.Worksheets("CATS_FILE").Range("J" & i).Value = "EAS"
            End If
    
            'Telco Provider
            If DupForm.DataFeedTypeComboBox <> "Annual (CLEC)" Then
                WK.Worksheets("CATS_FILE").Range("AK" & i).Value = DupForm.ProviderTextBox
            End If
            
            'Add section
            WK.Worksheets("CATS_FILE").Range("N" & i).Value = DupForm.SectionTextBox.Value
        Next i
    Else
        'Asigning the data to an array
        lastrow2 = WK.Worksheets("CATS_FILE").Range("J" & Rows.count).End(xlUp).row

        For i = lastrow2 + 1 To lastrow Step 1
            'If its empty the dup report remove all lines
            'Clec Provider
            If DupForm.DataFeedTypeComboBox = "Annual (CLEC)" Then
                WK.Worksheets("CATS_FILE").Range("H" & i).Value = DupForm.ProviderTextBox
            End If
            
            'Data Feed Type
            If DupForm.DataFeedTypeComboBox = "Annual (LOCAL)" Then
                WK.Worksheets("CATS_FILE").Range("J" & i).Value = "Annual"
            ElseIf DupForm.DataFeedTypeComboBox = "Annual (CLEC)" Then
                WK.Worksheets("CATS_FILE").Range("J" & i).Value = "Annual"
            Else
                WK.Worksheets("CATS_FILE").Range("J" & i).Value = "EAS"
            End If
    
            'Telco Provider
            If DupForm.DataFeedTypeComboBox <> "Annual (CLEC)" Then
                WK.Worksheets("CATS_FILE").Range("AK" & i).Value = DupForm.ProviderTextBox
            End If
            
            'Add section
            WK.Worksheets("CATS_FILE").Range("N" & i).Value = DupForm.SectionTextBox.Value
        Next i
    End If
    
End Sub

'Add section to the files
Public Sub Add_Section()
    Dim WK As Workbook, arr As Variant, section As String
    Set WK = ThisWorkbook
    
    'Take the last row from the file
    lastrow = WK.Worksheets("CATS_FILE").Range("AR" & Rows.count).End(xlUp).row

    For i = 1 To lastrow Step 1
        'If its empty the dup report remove all lines
        section = DupForm.SectionTextBox.Value
        WK.Worksheets("CATS_FILE").Range("N" & i).Value = section
    Next i
    
End Sub

Public Sub UploadFiles()
    Dim WK As Workbook, arr As Variant, class_of_service, phone, section As String
    Set WK = Workbooks(Workbooks.count)
    
    'Remove empty cell
    WK.Worksheets("CATS_FILE").Range("A1").EntireRow.Delete
    
    'cells counts
    lastrow = WK.Worksheets("CATS_FILE").Range("A" & Rows.count).End(xlUp).row
    
    WK.Worksheets("CATS_FILE").Range("B" & lastrow).Value = "LAST_LINE"
    
    counts = 0
    
    For i = 1 To lastrow Step 1
    
        counts = counts + 1
    
        'Phone number in the name field
        If WK.Worksheets("CATS_FILE").Range("B" & i).Value <> vbNullString Then
            
            'To add the last listing to the file it wasnt passing to the file.
            If WK.Worksheets("CATS_FILE").Range("B" & i).Value = "LAST_LINE" Then
                WK.Worksheets(1).Range("A" & WK.Worksheets(1).Range("A" & Rows.count).End(xlUp).row + 1) = WK.Worksheets("CATS_FILE").Range("A" & i).Value
            End If
            
            If counts > 1 Then

'                For Each r In Selection.Rows
'                    sTemp = ""
'                    For Each c In r.Cells
'                        sTemp = sTemp & c.Text & Chr(9)
'                    Next c
'
'                    'Get rid of trailing tabs
'                    While Right(sTemp, 1) = Chr(9)
'                        sTemp = Left(sTemp, Len(sTemp) - 1)
'                    Wend
'                    Print #1, sTemp
'                Next r
                            
                fname = "https://thryv.sharepoint.com/sites/DR-ServiceOrders/FODBMAnnual Load Files/--@Annual Load Duplicate Cleanup--/" & _
                WK.Worksheets("CATS_FILE").Range("B" & name_cell_number).Value & "."
                WK.Worksheets(1).SaveAs filename:=fname, FileFormat:=20
                MsgBox fname
                'NF.Close savechanges = False
                'RS.Worksheets(1).Visible = xlSheetVeryHidden
            End If
            
            'Create a sheet
            If WK.Worksheets("CATS_FILE").Range("B" & i).Value <> "LAST_LINE" Then
                'Create sheets + counts to have a different name
                Sheets.Add.name = Left(WK.Worksheets("CATS_FILE").Range("B" & i).Value, 20) & counts
            End If
                        
            'Move data line by line to the new sheet
            WK.Worksheets(1).Range("A" & WK.Worksheets(1).Range("A" & Rows.count).End(xlUp).row + 1) = WK.Worksheets("CATS_FILE").Range("A" & i).Value
            
            'Name cell address
            name_cell_number = Replace(WK.Worksheets("CATS_FILE").Range("B" & i).address, "$B$", "")
        Else
            WK.Worksheets(1).Range("A" & WK.Worksheets(1).Range("A" & Rows.count).End(xlUp).row + 1) = WK.Worksheets("CATS_FILE").Range("A" & i).Value
            
        End If
        
    Next i
    
    'Clear the column B
    WK.Worksheets("CATS_FILE").Columns("B").ClearContents
    Set WK = Workbooks(Workbooks.count)
        
End Sub

'Function Toll_Free_Indicator(line As Long) As Boolean
'Dim WK As Workbook, tempnumber As String
'Set WK = ThisWorkbook
'WK.Activate
'
'If Len(WK.Worksheets("CATS").Range("AE" & i).Value) >= 10 Or Len(WK.Worksheets("CATS").Range("AF" & i).Value) >= 10 Then
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 800 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 800 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 888 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 888 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 877 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 877 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 866 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 866 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 855 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 855 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 844 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 844 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 833 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 833 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 822 Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = 822 Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'    If Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = "1-8" Or Left(WK.Worksheets("CATS").Range("AE" & i).Value, 3) = "1-8" Then
'        Toll_Free_Indicator = True
'        GoTo tollfreeidentified
'    End If
'End If
'Toll_Free_Indicator = False
'
'tollfreeidentified:
'
'End Function
'
'
'
''Phone Override Original
'
'    If Len(RS.Worksheets("VFfile").Cells(i, 31).Value2) <= 6 Then
'        RS.Worksheets("VFfile").Cells(i, 32).Value2 = Right(RS.Worksheets("VFfile").Cells(i, 31).Value2, 3)
'    Else
'        RS.Worksheets("VFfile").Cells(i, 32).Value2 = RS.Worksheets("CATS").Cells(i, 20).Value2
'    End If


