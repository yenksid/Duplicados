VBA Code
Attribute VB_Name = "Module3"
Option Private Module
Dim fname As Variant
Public css_file As Boolean

Sub Duplicates_Mark_Automatic_Removals()

Dim ws As Worksheet, lastrow As Long, RS As Workbook, i As Long, j As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

Set RS = ThisWorkbook
lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row

For i = 2 To lastrow
j = i + 1
        
Repeat:


    If Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" Then
        If (FName_Exact_Dup(i, j) = True Or FName_Switch_Dup(i, j) = True) And LName_Exact_Dup(i, j) = True Then


            'CrossReference Dups
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Cross_Reference_Lines(i, j) = True Then
                
                If Empty_Address_Match(i, j) = True Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Cross Reference Dup, bottom discard"
                    j = j + 1
                    GoTo Repeat
                End If
                
            End If
            'Community Bottom Discard
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Street_Number(i, j) = "Keep_Either" And _
                Street_Name(i, j) = "Keep_Either" And _
                (Community(i, j) = "Keep_Either" Or Community(i, j) = "YesTop_NoBottom") Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Name and Address match, Community discard"
                    j = j + 1
                    GoTo Repeat
            End If

            'Community Top Discard
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Street_Number(i, j) = "Keep_Either" And _
                Street_Name(i, j) = "Keep_Either" And _
                Community(i, j) = "NoTop_YesBottom" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Name and Address match, Community discard"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                    GoTo NextIteration
            End If

            'Street Residential Bottom Discard
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Bus_Res_Gov(i, j) = "Res_vs_Res" And _
                Street_Number(i, j) = "NoTop_YesBottom" And _
                Street_Name(i, j) = "NoTop_YesBottom" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Name Match, Residential Street discard"
                    j = j + 1
                    GoTo Repeat
            End If

            'Street Residential Top Discard
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Bus_Res_Gov(i, j) = "Res_vs_Res" And _
                Street_Number(i, j) = "YesTop_NoBottom" And _
                Street_Name(i, j) = "YesTop_NoBottom" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Name Match, Residential Street discard"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                    GoTo NextIteration
            End If

            'Street Business Bottom Discard
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                (Street_Number(i, j) = "Keep_Either" Or Street_Number(i, j) = "YesTop_NoBottom") And _
                (Street_Name(i, j) = "Keep_Either" Or Street_Name(i, j) = "YesTop_NoBottom") Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Name Match, missing address discard"
                    j = j + 1
                    GoTo Repeat
            End If

            'Street Business Top Discard
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                (Street_Number(i, j) = "Keep_Either" Or Street_Number(i, j) = "NoTop_YesBottom") And _
                (Street_Name(i, j) = "Keep_Either" Or Street_Name(i, j) = "NoTop_YesBottom") Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Name Match, missing address discard"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                    GoTo NextIteration
            End If


            'Street number Bottom Discard NENA
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Street_Number(i, j) = "YesTop_NoBottom" And _
                Street_Name(i, j) = "NENA" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Name Match, NENA match, missing addr number discard"
                    j = j + 1
                    GoTo Repeat
            End If

            'Street number Top Discard NENA
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Street_Number(i, j) = "NoTop_YesBottom" And _
                Street_Name(i, j) = "NENA" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Name Match, NENA match, missing addr number discard"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                    GoTo NextIteration
            End If

            'Same Street number NENA Discard
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Street_Number(i, j) = "Keep_Either" And _
                Street_Name(i, j) = "NENA" Then
                    If Len(RS.Worksheets("Dataset").Cells(i, "J").Value2) >= Len(RS.Worksheets("Dataset").Cells(j, "J").Value2) Then
                        RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                        RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Name Match, Street discard NENA"
                        j = j + 1
                        GoTo Repeat
                    Else
                        RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Name Match, Street discard NENA"
                        RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                        GoTo NextIteration
                    End If
            End If

'Straights vs Straights, Exact Name and Phone, Different Type

DifferentType:

            'LOCAL vs CLEC Bottom Discard
            If DataFeedType_Telco_Clec(i, j) = "LOCAL_vs_CLEC" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "LOCAL vs CLEC, Discard CLEC"
                    j = j + 1
                    GoTo Repeat
            End If

            'CLEC vs LOCAL Top Discard
            If DataFeedType_Telco_Clec(i, j) = "CLEC_vs_LOCAL" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "CLEC vs LOCAL, Discard CLEC"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                    GoTo NextIteration
            End If

            'CLEC vs EAS Bottom Discard
            If DataFeedType_Telco_Clec(i, j) = "CLEC_vs_EAS" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "CLEC vs EAS, Discard EAS"
                    j = j + 1
                    GoTo Repeat
            End If

            'EAS vs CLEC Top Discard
            If DataFeedType_Telco_Clec(i, j) = "EAS_vs_CLEC" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "EAS vs CLEC, Discard EAS"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                    GoTo NextIteration
            End If

            'LOCAL vs EAS Bottom Discard
            If DataFeedType_Telco_Clec(i, j) = "LOCAL_vs_EAS" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "LOCAL vs EAS, Discard EAS"
                    j = j + 1
                    GoTo Repeat
            End If

            'EAS vs LOCAL Top Discard
            If DataFeedType_Telco_Clec(i, j) = "EAS_vs_LOCAL" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "EAS vs LOCAL, Discard EAS"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                    GoTo NextIteration
            End If
        Else
        
        
        'validated logic matches documentation
           'First Name Partial matching
            If LName_Exact_Dup(i, j) = True And Partial_FName_Match(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" Then
                If DataFeedType_Telco_Clec(i, j) = "Same_Type" Then
                    If Street_Number(i, j) = "Keep_Either" And (Street_Name(i, j) = "Keep_Either" Or Street_Name(i, j) = "NENA") Then
                        If Len(RS.Worksheets("Dataset").Cells(i, "G").Value2) < Len(RS.Worksheets("Dataset").Cells(j, "G").Value2) Then
                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Partial Name Match, Shorter Name Discard"
                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                            GoTo NextIteration
                        End If
                        If Len(RS.Worksheets("Dataset").Cells(i, "G").Value2) >= Len(RS.Worksheets("Dataset").Cells(j, "G").Value2) Then
                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Partial Name Match, Shorter Name Discard"
                            j = j + 1
                            GoTo Repeat
                        End If
                    End If
                    If Street_Number(i, j) = "NoTop_YesBottom" And Street_Name(i, j) = "NoTop_YesBottom" Then
                        RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                        RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Partial Name Match, Residential Street discard"
                        j = j + 1
                        GoTo Repeat
                    End If
                    If Street_Number(i, j) = "YesTop_NoBottom" And Street_Name(i, j) = "YesTop_NoBottom" Then
                        RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Partial Name Match, Residential Street discard"
                        RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                        GoTo NextIteration
                    End If
                Else
                    GoTo DifferentType
                End If

            End If

            'Special Character Removal Match in Business
            If Special_Char_Rem_Match(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" And Bus_Res_Gov(i, j) = "Bus_vs_Bus" Then
               If DataFeedType_Telco_Clec(i, j) = "Same_Type" Then
                    If Street_Number(i, j) = "Keep_Either" And (Street_Name(i, j) = "Keep_Either" Or Street_Name(i, j) = "NENA") Then
                        If Len(RS.Worksheets("Dataset").Cells(i, "H").Value2) < Len(RS.Worksheets("Dataset").Cells(j, "H").Value2) Then
                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Special Char Match, Shorter Name Discard"
                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                            GoTo NextIteration
                        End If
                        If Len(RS.Worksheets("Dataset").Cells(i, "H").Value2) >= Len(RS.Worksheets("Dataset").Cells(j, "H").Value2) Then
                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Special Char Match, Shorter Name Discard"
                            j = j + 1
                            GoTo Repeat
                        End If
                    End If
                Else
                    GoTo DifferentType
                End If
            End If



            'Dup Bus vs Res
            If Dup_BusVsRes(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" Then
                If DataFeedType_Telco_Clec(i, j) = "Same_Type" Then
                    If Street_Number(i, j) <> "Keep_Both" And Street_Name(i, j) <> "Keep_Both" And Bus_Res_Gov(i, j) = "Bus_vs_Res" Then
                        RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
                        RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Bus vs Res Dup, Residential discard"
                        j = j + 1
                        GoTo Repeat
                    End If
                    If Street_Number(i, j) <> "Keep_Both" And Street_Name(i, j) <> "Keep_Both" And Bus_Res_Gov(i, j) = "Res_vs_Bus" Then
                        RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Bus vs Res Dup, Residential discard"
                        RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
                        GoTo NextIteration
                    End If
                Else
                    GoTo DifferentType
                End If
            End If


            If LName_Exact_Dup(i, j) = True And DataFeedType_Telco_Clec(i, j) <> "Same_Type" And _
                Partial_FName_Match(i, j) = False And _
                Caption_Indicator(i, j) = "Str_vs_Str" Then
                    GoTo DifferentType
            End If


        End If
    End If
    
    
    
    
    'Straight vs Caption Criteria
    
    If Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Cap" And _
        FName_Exact_Dup(i, j) = True And LName_Exact_Dup(i, j) = True And _
        Street_Number(i, j) = "Keep_Either" And Street_Name(i, j) = "Keep_Either" Then
            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Straight vs Cap Dup, Straight discard"
            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
            GoTo NextIteration
    End If
    If DataFeedType_Telco_Clec(i, j) = "Same_Type" And Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Cap" And _
        FName_Exact_Dup(i, j) = True And LName_Exact_Dup(i, j) = True And _
        Street_Number(i, j) = "Keep_Either" And Street_Name(i, j) = "NENA" Then
            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Straight vs Cap Dup, Straight discard"
            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
            GoTo NextIteration
    End If
    If (DataFeedType_Telco_Clec(i, j) = "CLEC_vs_LOCAL" Or DataFeedType_Telco_Clec(i, j) = "EAS_vs_LOCAL" Or DataFeedType_Telco_Clec(i, j) = "EAS_vs_CLEC") And _
        Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Cap" And _
        FName_Exact_Dup(i, j) = True And LName_Exact_Dup(i, j) = True Then
            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Straight vs Cap Dup, Straight discard"
            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Keep"
            GoTo NextIteration
    End If
    
    If Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Cap_vs_Str" And _
        FName_Exact_Dup(i, j) = True And LName_Exact_Dup(i, j) = True And _
        Street_Number(i, j) = "Keep_Either" And Street_Name(i, j) = "Keep_Either" Then
            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Straight vs Cap Dup, Straight discard"
            GoTo NextIteration
    End If
    If DataFeedType_Telco_Clec(i, j) = "Same_Type" And Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Cap_vs_Str" And _
        FName_Exact_Dup(i, j) = True And LName_Exact_Dup(i, j) = True And _
        Street_Number(i, j) = "Keep_Either" And Street_Name(i, j) = "NENA" Then
            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Straight vs Cap Dup, Straight discard"
            GoTo NextIteration
    End If
    If (DataFeedType_Telco_Clec(i, j) = "LOCAL_vs_CLEC" Or DataFeedType_Telco_Clec(i, j) = "LOCAL_vs_EAS" Or DataFeedType_Telco_Clec(i, j) = "CLEC_vs_EAS") And _
        Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Cap_vs_Str" And _
        FName_Exact_Dup(i, j) = True And LName_Exact_Dup(i, j) = True Then
            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Keep"
            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Straight vs Cap Dup, Straight discard"
            GoTo NextIteration
    End If
    
'change log 1.3 - added condition below
            If Section_And_Phone_Dup(i, j) = True Then
                j = j + 1
                GoTo Repeat
            Else
                While RS.Worksheets("Dataset").Cells(i + 1, "P").Value2 <> vbNullString
                    i = i + 1
                Wend
                GoTo NextForI
            End If
 
NextIteration:
            If j <> i + 1 Then
                i = j - 1
            End If

NextForI:

Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub
Sub Duplicates_Mark_Near()

Dim ws As Worksheet, lastrow As Long, RS As Workbook, i As Long, j As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

Set RS = ThisWorkbook
lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row

For i = 2 To lastrow
j = i + 1
NearRepeat:
    If Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" Then
        If (FName_Exact_Dup(i, j) = True Or FName_Switch_Dup(i, j) = True) And LName_Exact_Dup(i, j) = True Then

'            Street Business to Report
            If DataFeedType_Telco_Clec(i, j) = "Same_Type" And _
                Bus_Res_Gov(i, j) = "Bus_vs_Bus" And _
                Street_Number(i, j) <> "Keep_Both" And Street_Name(i, j) <> "Keep_Both" Then
                'need to review commented below, doesnt make sense why we are choosing this approach
'                ((Street_Number(i, j) = "YesTop_NoBottom" And Street_Name(i, j) = "YesTop_NoBottom") Or _
'                (Street_Number(i, j) = "NoTop_YesBottom" And Street_Name(i, j) = "NoTop_YesBottom") Or _
'                (Street_Number(i, j) = "Keep_Either" And Street_Name(i, j) = "Address_Partial_Match")) Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Action Required"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Action Required"
'
'                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Street Business"
'                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Street Business"
                    GoTo NearNextIteration
            End If

        Else

            'Special Character Removal Match in Business
            If Special_Char_Rem_Match(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" And _
                Bus_Res_Gov(i, j) = "Bus_vs_Bus" And DataFeedType_Telco_Clec(i, j) = "Same_Type" Then
                'need to review commented below, doesnt make sense why we are choosing this approach
'                    If ((Street_Number(i, j) = "YesTop_NoBottom" And Street_Name(i, j) = "YesTop_NoBottom") Or _
'                        (Street_Number(i, j) = "NoTop_YesBottom" And Street_Name(i, j) = "NoTop_YesBottom") Or _
'                        (Street_Number(i, j) = "Keep_Either" And Street_Name(i, j) = "Address_Partial_Match")) Then
                    If Street_Number(i, j) <> "Keep_Both" And Street_Name(i, j) <> "Keep_Both" Then
                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Action Required"
                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Action Required"
                            
'
'                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Special Character Removal"
'                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Special Character Removal"
                            GoTo NearNextIteration
                    End If
                End If

'            'Space Partial First Name Matching
            If LName_Exact_Dup(i, j) = True And Space_Partial_FName_Match(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" And _
                Street_Number(i, j) <> "Keep_Both" And Street_Name(i, j) <> "Keep_Both" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Action Required"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Action Required"
                    
                    
                            
'                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Space Partial First Name"
'                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Space Partial First Name"
                    GoTo NearNextIteration
            End If

'            'Business Name Partial matching
            If Partial_LName_Match(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" And Bus_Res_Gov(i, j) = "Bus_vs_Bus" And _
                Street_Number(i, j) <> "Keep_Both" And Street_Name(i, j) <> "Keep_Both" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Action Required"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Action Required"
'
                    
'                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Business Name Partial"
'                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Business Name Partial"
                    GoTo NearNextIteration
            End If

            'different data feed partial name match in bus
            If Partial_LName_Match(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" And Bus_Res_Gov(i, j) = "Bus_vs_Bus" And _
                DataFeedType_Telco_Clec(i, j) <> "Same_Type" Then
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Action Required"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Action Required"
'
                    
'                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "different data feed partial name match in bus"
'                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "different data feed partial name match in bus"
                    GoTo NearNextIteration
            End If
            
            
            'change log 1.4 matching of names where only one word is different
            If one_word_mismatch(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" And Bus_Res_Gov(i, j) = "Bus_vs_Bus" And _
                    DataFeedType_Telco_Clec(i, j) <> "Same_Type" Then
                    
                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Action Required"
                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Action Required"
                    
                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "one word mismatch different type"
                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "one word mismatch different type"
                    GoTo NearNextIteration
            End If
            
            
            'change log 1.4 logic below not used, only matching when datafeedtype not equal (above)
'            If one_word_mismatch(i, j) = True And Caption_Indicator(i, j) = "Str_vs_Str" And Bus_Res_Gov(i, j) = "Bus_vs_Bus" And _
'                Street_Number(i, j) <> "Keep_Both" And Street_Name(i, j) <> "Keep_Both" Then
''                    RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Action Required"
''                    RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Action Required"
''
'
'                            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "one word mismatch same type"
'                            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "one word mismatch same type"
'                    GoTo NearNextIteration
'            End If

        End If
    End If
    
    'change log 1.3 - added condition below
            If Section_And_Phone_Dup(i, j) = True Then
                j = j + 1
                GoTo NearRepeat
            Else
                While RS.Worksheets("Dataset").Cells(i + 1, "P").Value2 <> vbNullString
                    i = i + 1
                Wend
                GoTo NextNearForI
            End If
    
NearNextIteration:
            If j <> i + 1 Then
                i = j - 1
            End If

NextNearForI:

Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub
Sub Duplicates_Mark_Captions()

Dim ws As Worksheet, lastrow As Long, RS As Workbook, i As Long, j As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


Set RS = ThisWorkbook

lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row

RS.Worksheets("Dataset").Range("P2") = "=E2&M2"
If RS.Worksheets("Dataset").Range("M3") <> vbNullString Then
    RS.Worksheets("Dataset").Range("P2").AutoFill Destination:=Worksheets("Dataset").Range("P2:P" & lastrow)
End If
For i = lastrow To 2 Step -1

'Deleting all unique phones in captions
If Application.WorksheetFunction.CountIf(RS.Worksheets("Dataset").Range("P:P"), RS.Worksheets("Dataset").Range("P" & i)) = 1 Then
    RS.Worksheets("Dataset").Rows(i).Delete
End If

Next i



lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row
RS.Worksheets("Dataset").Range("P2:P" & lastrow).Clear

RS.Worksheets("Dataset").Range("P2") = "=E2&M2&""-""&O2"
If RS.Worksheets("Dataset").Range("M3") <> vbNullString Then
    RS.Worksheets("Dataset").Range("P2").AutoFill Destination:=Worksheets("Dataset").Range("P2:P" & lastrow)
End If
For i = lastrow To 2 Step -1

'Deleting all duplicate phones within the same caption
If Application.WorksheetFunction.CountIf(RS.Worksheets("Dataset").Range("P:P"), RS.Worksheets("Dataset").Range("P" & i)) <> 1 Then
    RS.Worksheets("Dataset").Rows(i).Delete
End If

Next i


lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row
RS.Worksheets("Dataset").Range("P2:P" & lastrow).Clear

'sorting by section, phone and caption header
RS.Worksheets("Dataset").Range("P2") = "=E2&M2&O2"
If RS.Worksheets("Dataset").Range("M3") <> vbNullString Then
    RS.Worksheets("Dataset").Range("P2").AutoFill Destination:=Worksheets("Dataset").Range("P2:P" & lastrow)
End If

RS.Worksheets("Dataset").Sort.SortFields.Add2 key:=Range("P:P"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With RS.Worksheets("Dataset").Sort
    .SetRange Range("A:P")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

For i = lastrow To 2 Step -1
'deleting all duplicate phone numbers under the same caption
If RS.Worksheets("Dataset").Range("P" & i).Value2 = RS.Worksheets("Dataset").Range("P" & i - 1).Value2 Then
    RS.Worksheets("Dataset").Rows(i).Delete
End If

Next i

lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row
RS.Worksheets("Dataset").Range("P2:P" & lastrow).Clear

RS.Worksheets("Dataset").Range("P2") = "=E2&SUBSTITUTE(H2,"" "","""")&M2"
If RS.Worksheets("Dataset").Range("M3") <> vbNullString Then
    RS.Worksheets("Dataset").Range("P2").AutoFill Destination:=Worksheets("Dataset").Range("P2:P" & lastrow)
End If
With RS.Worksheets("Dataset").UsedRange
.Value = .Value
End With

'RS.Worksheets("Dataset").Columns("A:P").Sort key1:=Range("P2"), order1:=xlAscending, Header:=xlYes

RS.Worksheets("Dataset").Sort.SortFields.Clear
RS.Worksheets("Dataset").Sort.SortFields.Add2 key:=Range("P:P"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
With RS.Worksheets("Dataset").Sort
    .SetRange Range("A:P")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'
'
lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row
RS.Worksheets("Dataset").Range("P2:P" & lastrow).Clear


For i = 2 To lastrow Step 1
j = i + 1


    If RS.Worksheets("Dataset").Cells(i, "M").Value2 = vbNullString Then Exit For

    If Section_And_Phone_Dup(i, j) = True And Caption_Indicator(i, j) = "Cap_vs_Cap" And _
        FName_Exact_Dup(i, j) = True And Special_Char_Rem_Match(i, j) = True Then
            RS.Worksheets("Dataset").Cells(i, "P").Value2 = "Cap Dup"
            RS.Worksheets("Dataset").Cells(j, "P").Value2 = "Cap Dup"
    End If


    If RS.Worksheets("Dataset").Cells(i, "P").Value2 = vbNullString Then
        RS.Worksheets("Dataset").Rows(i).Delete
        i = i - 1
    End If
Next i

RS.Worksheets("Dataset").Range("A:P").RemoveDuplicates Columns:=15, Header:=xlYes

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Sub Report_Automatic_Removal()

Dim ws As Worksheet, lastrow As Long, RS As Workbook, i As Long, j As Long, x As Long, y As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


Set RS = ThisWorkbook

        
RS.Worksheets("Dataset").Copy After:=Worksheets(Sheets.count)
RS.Worksheets(Sheets.count).name = "Automatic Removal Dups"

RS.Worksheets("Automatic Removal Dups").Range("Z1").Clear

RS.Worksheets("Automatic Removal Dups").Range("E1") = "Section"
RS.Worksheets("Automatic Removal Dups").Range("R1") = "Source"
RS.Worksheets("Automatic Removal Dups").Range("S1") = "Listing"
RS.Worksheets("Automatic Removal Dups").Range("T1") = "Phone"
RS.Worksheets("Automatic Removal Dups").Range("U1") = "ACTION"
RS.Worksheets("Automatic Removal Dups").Range("Q1") = "Sort"
RS.Worksheets("Automatic Removal Dups").Range("Q2") = RS.Worksheets("Automatic Removal Dups").Range("E2") & RS.Worksheets("Automatic Removal Dups").Range("H2") & RS.Worksheets("Automatic Removal Dups").Range("M2")
lastrow = RS.Worksheets("Automatic Removal Dups").Range("M" & Rows.count).End(xlUp).row

For i = 2 To lastrow Step 1

    If RS.Worksheets("Automatic Removal Dups").Range("M" & i).Value2 = vbNullString Then Exit For
    If RS.Worksheets("Automatic Removal Dups").Range("P" & i).Value2 = vbNullString Then

        RS.Worksheets("Automatic Removal Dups").Rows(i).Delete
        i = i - 1
        GoTo DeletedLine
    End If
    If RS.Worksheets("Automatic Removal Dups").Range("P" & i).Value2 = "Keep" Then RS.Worksheets("Automatic Removal Dups").Range("P" & i).Clear
    If i > 2 And RS.Worksheets("Automatic Removal Dups").Range("M" & i - 1).Value2 = RS.Worksheets("Automatic Removal Dups").Range("M" & i).Value2 Then
        RS.Worksheets("Automatic Removal Dups").Range("Q" & i).Value2 = RS.Worksheets("Automatic Removal Dups").Range("Q" & i - 1).Value2
    Else
        RS.Worksheets("Automatic Removal Dups").Range("Q" & i) = RS.Worksheets("Automatic Removal Dups").Range("E" & i) & RS.Worksheets("Automatic Removal Dups").Range("H" & i) & RS.Worksheets("Automatic Removal Dups").Range("M" & i)
    End If

DeletedLine:

Next i

RS.Worksheets("Automatic Removal Dups").Sort.SortFields.Add2 key:=Range("Q:Q"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With RS.Worksheets("Automatic Removal Dups").Sort
    .SetRange Range("A:Q")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

lastrow = RS.Worksheets("Automatic Removal Dups").Range("M" & Rows.count).End(xlUp).row

If lastrow <> 1 Then
    RS.Worksheets("Automatic Removal Dups").Range("R2") = "=IF(AND(B2=""Annual"", C2<>""""),""LOCAL""&""|""&C2,IF(AND(B2=""Annual"",C2=""""),""CLEC""&""|""&D2,""EAS""&""|""&C2))"
    RS.Worksheets("Automatic Removal Dups").Range("R2").AutoFill Destination:=Worksheets("Automatic Removal Dups").Range("R2:R" & lastrow)
    
    RS.Worksheets("Automatic Removal Dups").Range("S2") = "=IF(G2="""",H2,H2&"", ""&G2)&IF(I2="""","""","" | ""&I2)&"" | ""&J2&"" ""&K2&IF(L2<>"""","" | ""&L2,"""")"
    RS.Worksheets("Automatic Removal Dups").Range("S2").AutoFill Destination:=Worksheets("Automatic Removal Dups").Range("S2:S" & lastrow)
    
    RS.Worksheets("Automatic Removal Dups").Range("T2") = "=M2"
    RS.Worksheets("Automatic Removal Dups").Range("T2").AutoFill Destination:=Worksheets("Automatic Removal Dups").Range("T2:T" & lastrow)
    
    RS.Worksheets("Automatic Removal Dups").Range("U2") = "=IF(P2<>"""",P2,"""")"
    RS.Worksheets("Automatic Removal Dups").Range("U2").AutoFill Destination:=Worksheets("Automatic Removal Dups").Range("U2:U" & lastrow)
End If

With RS.Worksheets("Automatic Removal Dups").UsedRange
.Value = .Value
End With


For i = lastrow To 2 Step -1

    If RS.Worksheets("Automatic Removal Dups").Range("M" & i).Value2 <> RS.Worksheets("Automatic Removal Dups").Range("M" & i - 1).Value2 Then

        RS.Worksheets("Automatic Removal Dups").Rows(i).EntireRow.Insert

    End If

Next i


RS.Worksheets("Automatic Removal Dups").Columns("N").Cut
RS.Worksheets("Automatic Removal Dups").Columns("V").Insert Shift:=xlToRight
RS.Worksheets("Automatic Removal Dups").Columns("F:P").Delete
RS.Worksheets("Automatic Removal Dups").Columns("A:D").Delete
RS.Worksheets("Automatic Removal Dups").Columns("E").Cut
RS.Worksheets("Automatic Removal Dups").Columns("A").Insert Shift:=xlToRight

RS.Worksheets("Automatic Removal Dups").Columns("F").Hidden = True
RS.Worksheets("Automatic Removal Dups").Range("A1:E1").Font.Bold = True
RS.Worksheets("Automatic Removal Dups").Columns("A:E").AutoFit
RS.Worksheets("Automatic Removal Dups").Protect Password:="stayaway", AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
RS.Worksheets("Automatic Removal Dups").Cells(1, 1).Select



'removing
lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row
For i = lastrow To 2 Step -1

    If RS.Worksheets("Dataset").Range("P" & i).Value2 <> vbNullString And RS.Worksheets("Dataset").Range("P" & i).Value2 <> "Keep" Then

        RS.Worksheets("Dataset").Rows(i).Delete

    End If
    If RS.Worksheets("Dataset").Range("P" & i).Value2 = "Keep" Then

        RS.Worksheets("Dataset").Range("P" & i).Clear

    End If

Next i
'RS.Worksheets("Dataset").Cells(1, 1).Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub
Sub Report_Near_Dups()

Dim ws As Worksheet, lastrow As Long, RS As Workbook, i As Long, j As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


Set RS = ThisWorkbook

        
RS.Worksheets("Dataset").Copy After:=Worksheets(Sheets.count)
RS.Worksheets(Sheets.count).name = "Near Dups"


RS.Worksheets("Near Dups").Range("Z1").Clear

RS.Worksheets("Near Dups").Range("E1") = "Section"
RS.Worksheets("Near Dups").Range("R1") = "Source"
RS.Worksheets("Near Dups").Range("S1") = "Listing"
RS.Worksheets("Near Dups").Range("T1") = "Phone"
RS.Worksheets("Near Dups").Range("U1") = "ACTION"
RS.Worksheets("Near Dups").Range("Q1") = "Sort"
RS.Worksheets("Near Dups").Range("Q2") = RS.Worksheets("Near Dups").Range("E2") & RS.Worksheets("Near Dups").Range("H2") & RS.Worksheets("Near Dups").Range("M2")
lastrow = RS.Worksheets("Near Dups").Range("M" & Rows.count).End(xlUp).row

For i = 2 To lastrow Step 1

    If RS.Worksheets("Near Dups").Range("M" & i).Value2 = vbNullString Then Exit For
    If RS.Worksheets("Near Dups").Range("P" & i).Value2 = vbNullString Then

        RS.Worksheets("Near Dups").Rows(i).Delete
        i = i - 1
        GoTo DeletedLine
    End If
    If RS.Worksheets("Near Dups").Range("P" & i).Value2 = "Action Required" Then RS.Worksheets("Near Dups").Range("P" & i).Clear
    If i > 2 And RS.Worksheets("Near Dups").Range("M" & i - 1).Value2 = RS.Worksheets("Near Dups").Range("M" & i).Value2 Then
        RS.Worksheets("Near Dups").Range("Q" & i).Value2 = RS.Worksheets("Near Dups").Range("Q" & i - 1).Value2
    Else
        RS.Worksheets("Near Dups").Range("Q" & i) = RS.Worksheets("Near Dups").Range("E" & i) & RS.Worksheets("Near Dups").Range("H" & i) & RS.Worksheets("Near Dups").Range("M" & i)
    End If

DeletedLine:

Next i

RS.Worksheets("Near Dups").Sort.SortFields.Add2 key:=Range("Q:Q"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With RS.Worksheets("Near Dups").Sort
    .SetRange Range("A:Q")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

lastrow = RS.Worksheets("Near Dups").Range("M" & Rows.count).End(xlUp).row

If lastrow <> 1 Then
    RS.Worksheets("Near Dups").Range("R2") = "=IF(AND(B2=""Annual"", C2<>""""),""LOCAL""&""|""&C2,IF(AND(B2=""Annual"",C2=""""),""CLEC""&""|""&D2,""EAS""&""|""&C2))"
    RS.Worksheets("Near Dups").Range("R2").AutoFill Destination:=Worksheets("Near Dups").Range("R2:R" & lastrow)
    
    RS.Worksheets("Near Dups").Range("S2") = "=IF(G2="""",H2,H2&"", ""&G2)&IF(I2="""","""","" | ""&I2)&"" | ""&J2&"" ""&K2&IF(L2<>"""","" | ""&L2,"""")"
    RS.Worksheets("Near Dups").Range("S2").AutoFill Destination:=Worksheets("Near Dups").Range("S2:S" & lastrow)
    
    RS.Worksheets("Near Dups").Range("T2") = "=M2"
    RS.Worksheets("Near Dups").Range("T2").AutoFill Destination:=Worksheets("Near Dups").Range("T2:T" & lastrow)
    
    
    '**** use below to debug which logic is being used for near dup matching
'    RS.Worksheets("Near Dups").Range("U2") = "=IF(P2<>"""",P2,"""")"
'    RS.Worksheets("Near Dups").Range("U2").AutoFill Destination:=Worksheets("Near Dups").Range("U2:U" & lastrow)
End If

With RS.Worksheets("Near Dups").UsedRange
.Value = .Value
End With


For i = lastrow To 2 Step -1

    If RS.Worksheets("Near Dups").Range("M" & i).Value2 <> RS.Worksheets("Near Dups").Range("M" & i - 1).Value2 Then
        RS.Worksheets("Near Dups").Range("U" & i).Locked = False
        RS.Worksheets("Near Dups").Rows(i).EntireRow.Insert
    Else
        RS.Worksheets("Near Dups").Range("U" & i).Locked = False

    End If

Next i


RS.Worksheets("Near Dups").Columns("N").Cut
RS.Worksheets("Near Dups").Columns("V").Insert Shift:=xlToRight
RS.Worksheets("Near Dups").Columns("F:P").Delete
RS.Worksheets("Near Dups").Columns("A:D").Delete
RS.Worksheets("Near Dups").Columns("E").Cut
RS.Worksheets("Near Dups").Columns("A").Insert Shift:=xlToRight

RS.Worksheets("Near Dups").Columns("F").Hidden = True
RS.Worksheets("Near Dups").Range("A1:E1").Font.Bold = True
RS.Worksheets("Near Dups").Columns("A:E").AutoFit
RS.Worksheets("Near Dups").Protect Password:="stayaway", AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
RS.Worksheets("Near Dups").Cells(1, 1).Select

'removing
lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row
For i = lastrow To 2 Step -1

    If RS.Worksheets("Dataset").Range("F" & i).Value2 = 0 Then

        RS.Worksheets("Dataset").Rows(i).Delete

    End If

Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Sub Report_Captions()

Dim ws As Worksheet, lastrow As Long, RS As Workbook, i As Long, j As Long, k As Long, caption_line As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

Set RS = ThisWorkbook
RS.Sheets.Add(After:=Sheets(Sheets.count)).name = "Caption Header Dups"
RS.Worksheets("Caption Header Dups").Range("A1") = "ACTION"
RS.Worksheets("Caption Header Dups").Range("B1") = "Section"
RS.Worksheets("Caption Header Dups").Range("C1") = "Source"
RS.Worksheets("Caption Header Dups").Range("D1") = "Caption Header"
RS.Worksheets("Caption Header Dups").Range("E1") = "Caption Line"
RS.Worksheets("Caption Header Dups").Range("F1") = "Phone"

lastrow = RS.Worksheets("Dataset").Range("M" & Rows.count).End(xlUp).row


j = 2

For i = 2 To lastrow Step 1

'fvs : I change column to E from O
k = RS.Worksheets("Dataset").Range("O" & i).Value2 + 1


    If FName_Exact_Dup(i, i - 1) = False Or Special_Char_Rem_Match(i, i - 1) = False Then
        j = j + 1
    End If
    
    RS.Worksheets("Caption Header Dups").Range("A" & j).Locked = False
    RS.Worksheets("Caption Header Dups").Range("B" & j).Value2 = RS.Worksheets("Dataset").Range("E" & i).Value2
    RS.Worksheets("Caption Header Dups").Range("G" & j).Value2 = RS.Worksheets("Dataset").Range("O" & i).Value2
    
    If RS.Worksheets("Dataset").Range("B" & i).Value2 = "Annual" Then
        If RS.Worksheets("Dataset").Range("C" & i).Value2 <> vbNullString Then
            RS.Worksheets("Caption Header Dups").Range("C" & j).Value2 = "LOCAL|" & RS.Worksheets("Dataset").Range("C" & i).Value2
        Else
            RS.Worksheets("Caption Header Dups").Range("C" & j).Value2 = "CLEC|" & RS.Worksheets("Dataset").Range("D" & i).Value2
        End If
    Else
        RS.Worksheets("Caption Header Dups").Range("C" & j).Value2 = "EAS|" & RS.Worksheets("Dataset").Range("C" & i).Value2
    End If
    
    RS.Worksheets("Caption Header Dups").Range("D" & j).Value2 = RS.Worksheets("Dataset").Range("H" & i).Value2
    
    Do While RS.Worksheets("CATS_FILE").Range("T" & k).Value2 <> 0 And RS.Worksheets("CATS_FILE").Range("T" & k).Value2 <> 10
        j = j + 1
        
        'creating caption lines
        caption_line = ""
        'concatenating caption text
        If RS.Worksheets("CATS_FILE").Range("E" & k).Value2 <> vbNullString Then caption_line = caption_line & RS.Worksheets("CATS_FILE").Range("E" & k).Value2
        'concatenating street number and address
        If RS.Worksheets("CATS_FILE").Range("AB" & k).Value2 <> vbNullString Then
            If caption_line <> "" Then caption_line = caption_line & " | "
            caption_line = caption_line & RS.Worksheets("CATS_FILE").Range("AC" & k).Value2 & " " & RS.Worksheets("CATS_FILE").Range("AB" & k).Value2
        End If
        'concatenating community
        If RS.Worksheets("CATS_FILE").Range("W" & k).Value2 <> vbNullString Then
            If caption_line <> "" Then caption_line = caption_line & " | "
            caption_line = caption_line & RS.Worksheets("CATS_FILE").Range("W" & k).Value2
        End If
        
        RS.Worksheets("Caption Header Dups").Range("E" & j).Value2 = caption_line
        RS.Worksheets("Caption Header Dups").Range("F" & j).Value2 = RS.Worksheets("CATS_FILE").Range("AE" & k).Value2
        
        
        k = k + 1
    Loop
    
    
    j = j + 1
Next i

RS.Worksheets("Caption Header Dups").Range("A1:F1").Font.Bold = True
RS.Worksheets("Caption Header Dups").Columns("A:F").AutoFit
RS.Worksheets("Caption Header Dups").Columns("G").Hidden = True
RS.Worksheets("Caption Header Dups").Protect Password:="stayaway", AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
RS.Worksheets("Caption Header Dups").Cells(1, 1).Select


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub


Sub Dupe_Dataset(fname As Variant)

Dim ws As Worksheet, lastrow As Long, NF As Workbook, RS As Workbook, file_Name As Variant, i As Long, j As Long, x As Long, y As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


Set RS = ThisWorkbook
        
        'Calculating the last row in CATS_FILE
        lastrow = RS.Worksheets("CATS_FILE").Range("N" & Rows.count).End(xlUp).row
           
        'Adding New Sheet - Dataset"

        Sheets.Add After:=Sheets("CATS_FILE")
        Set ws = ActiveSheet
        ws.name = "Dataset"
        
        'RS.Worksheets("CATS_FILE").Visible = xlSheetVeryHidden
    
        RS.Worksheets("Dataset").Range("A1") = "Bus_Res_Gov"
        RS.Worksheets("Dataset").Range("B1") = "Data_Feed_Type"
        RS.Worksheets("Dataset").Range("C1") = "Telco_Provider"
        RS.Worksheets("Dataset").Range("D1") = "CLEC_Provider"
        RS.Worksheets("Dataset").Range("E1") = "Section_Code"
        RS.Worksheets("Dataset").Range("F1") = "Caption Indicator"
        RS.Worksheets("Dataset").Range("G1") = "First_Name"
        RS.Worksheets("Dataset").Range("H1") = "Name"
        RS.Worksheets("Dataset").Range("I1") = "Content"
        RS.Worksheets("Dataset").Range("J1") = "Street_Number"
        RS.Worksheets("Dataset").Range("K1") = "Street"
        RS.Worksheets("Dataset").Range("L1") = "Community"
        RS.Worksheets("Dataset").Range("M1") = "Phone"
        RS.Worksheets("Dataset").Range("N1") = "Line Number"
        RS.Worksheets("Dataset").Range("O1") = "Line Number Caption Header"
        RS.Worksheets("Dataset").Range("P1") = "Action Taken"
        RS.Worksheets("Dataset").Range("Z1") = fname

        x = 2
        y = 3
        While y <= lastrow

            While RS.Worksheets("CATS_FILE").Cells(x, "N") = RS.Worksheets("CATS_FILE").Cells(y, "N")
                y = y + 1
            Wend

            For i = x To y - 1 Step 1


                If RS.Worksheets("CATS_FILE").Range("AE" & i).Value2 <> vbNullString And _
                Application.WorksheetFunction.CountIf(RS.Worksheets("CATS_FILE").Range("AE" & x & ":AE" & y), RS.Worksheets("CATS_FILE").Cells(i, "AE")) > 1 Then

                'Bus_Res CATS D
                    RS.Worksheets("Dataset").Range("A" & Range("A" & Rows.count).End(xlUp).row + 1).Value2 = RS.Worksheets("CATS_FILE").Range("D" & i).Value2
                'Data_Feed_Type CATS J
                    RS.Worksheets("Dataset").Range("B" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("J" & i).Value2
                'Telco Provider CATS AK
                    RS.Worksheets("Dataset").Range("C" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AK" & i).Value2
                'CLEC_Provider CATS H
                    RS.Worksheets("Dataset").Range("D" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("H" & i).Value2
                'Section_Code CATS N
                    RS.Worksheets("Dataset").Range("E" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("N" & i).Value2
                'Caption_Indicator CATS T
                    RS.Worksheets("Dataset").Range("F" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("T" & i).Value2

                'Straight Names
                    If RS.Worksheets("CATS_FILE").Range("AD" & i).Value2 <> vbNullString Then
                            'Fist_Name CATS Q
                            RS.Worksheets("Dataset").Range("G" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("Q" & i).Value2
                            'Last/Business Name CATS AD
                            RS.Worksheets("Dataset").Range("H" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AD" & i).Value2
                            
                            'fvs : line number caption header : this line wasnt here: You can remove it.
                            RS.Worksheets("Dataset").Range("O" & Range("A" & Rows.count).End(xlUp).row).Value2 = i
                    Else
                'Caption Names
                        j = i
                        While RS.Worksheets("CATS_FILE").Range("AD" & j).Value2 = vbNullString
                            j = j - 1
                        Wend
                        'first name
                        RS.Worksheets("Dataset").Range("G" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("Q" & j).Value2
                        'last name
                        RS.Worksheets("Dataset").Range("H" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AD" & j).Value2
                        'content
                        RS.Worksheets("Dataset").Range("I" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("E" & i).Value2
                        'line number caption header
                        RS.Worksheets("Dataset").Range("O" & Range("A" & Rows.count).End(xlUp).row).Value2 = j
                    End If

                'Street Number CATS AC
                    RS.Worksheets("Dataset").Range("J" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AC" & i).Value2
                'Street Name CATS AB
                    RS.Worksheets("Dataset").Range("K" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AB" & i).Value2
                'Community CATS W
                    RS.Worksheets("Dataset").Range("L" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("W" & i).Value2
                'Phone CATS AE
                    RS.Worksheets("Dataset").Range("M" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AE" & i).Value2
                'Line Number = i
                    RS.Worksheets("Dataset").Range("N" & Range("A" & Rows.count).End(xlUp).row).Value2 = i
                    
                ElseIf RS.Worksheets("CATS_FILE").Range("I" & i).Value2 <> vbNullString And _
                Application.WorksheetFunction.CountIf(RS.Worksheets("CATS_FILE").Range("I" & x & ":I" & y), RS.Worksheets("CATS_FILE").Cells(i, "I")) > 1 Then

                'Bus_Res CATS D
                    RS.Worksheets("Dataset").Range("A" & Range("A" & Rows.count).End(xlUp).row + 1).Value2 = RS.Worksheets("CATS_FILE").Range("D" & i).Value2
                'Data_Feed_Type CATS J
                    RS.Worksheets("Dataset").Range("B" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("J" & i).Value2
                'Telco Provider CATS AK
                    RS.Worksheets("Dataset").Range("C" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AK" & i).Value2
                'CLEC_Provider CATS H
                    RS.Worksheets("Dataset").Range("D" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("H" & i).Value2
                'Section_Code CATS N
                    RS.Worksheets("Dataset").Range("E" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("N" & i).Value2
                'Caption_Indicator CATS T
                    RS.Worksheets("Dataset").Range("F" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("T" & i).Value2

                'Straight Names
                    If RS.Worksheets("CATS_FILE").Range("AD" & i).Value2 <> vbNullString Then
                            'Fist_Name CATS Q
                            RS.Worksheets("Dataset").Range("G" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("Q" & i).Value2
                            'Last/Business Name CATS AD
                            RS.Worksheets("Dataset").Range("H" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AD" & i).Value2
                    Else
                'Caption Names
                        j = i
                        While RS.Worksheets("CATS_FILE").Range("AD" & j).Value2 = vbNullString
                            j = j - 1
                        Wend
                        RS.Worksheets("Dataset").Range("G" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("Q" & j).Value2
                        RS.Worksheets("Dataset").Range("H" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AD" & j).Value2
                        RS.Worksheets("Dataset").Range("I" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("E" & i).Value2
                        RS.Worksheets("Dataset").Range("O" & Range("A" & Rows.count).End(xlUp).row).Value2 = j
                    End If

                'Street Number CATS AC
                    RS.Worksheets("Dataset").Range("J" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AC" & i).Value2
                'Street Name CATS AB
                    RS.Worksheets("Dataset").Range("K" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("AB" & i).Value2
                'Community CATS W
                    RS.Worksheets("Dataset").Range("L" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("W" & i).Value2
                'Xref CATS I
                    RS.Worksheets("Dataset").Range("M" & Range("A" & Rows.count).End(xlUp).row).Value2 = RS.Worksheets("CATS_FILE").Range("I" & i).Value2
                'Line Number = i
                    RS.Worksheets("Dataset").Range("N" & Range("A" & Rows.count).End(xlUp).row).Value2 = i

                End If

            Next i

            x = y

        Wend

        RS.Worksheets("Dataset").Sort.SortFields.Clear
        RS.Worksheets("Dataset").Sort.SortFields.Add2 key:=Range( _
            "E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        RS.Worksheets("Dataset").Sort.SortFields.Add2 key:=Range( _
            "M:M"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        RS.Worksheets("Dataset").Sort.SortFields.Add2 key:=Range( _
            "H:H"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        RS.Worksheets("Dataset").Sort.SortFields.Add2 key:=Range( _
            "F:F"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With RS.Worksheets("Dataset").Sort
            .SetRange Range("A:O")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

'******OLD SUB TO DELETE******


'Sub Deleting_Process()
'
'
'Dim WS As Worksheet, lastrow As Long, RS As Workbook, i As Long, j As Long
'
'Set RS = ThisWorkbook
'
''Automatic Dup Delete Mark
'lastrow = RS.Worksheets("Automatic Removal Dups").Range("B" & Rows.Count).End(xlUp).Row
'
'For i = 2 To lastrow
'
'    If RS.Worksheets("Automatic Removal Dups").Range("A" & i).Value2 <> vbNullString Then
'        RS.Worksheets("CATS_FILE").Cells(RS.Worksheets("Automatic Removal Dups").Range("F" & i).Value2, 10).Clear
'    End If
'
'Next i
'
'
''Near Dup Delete Mark
'lastrow = RS.Worksheets("Near Dups").Range("B" & Rows.Count).End(xlUp).Row
'
'For i = 2 To lastrow
'
'    If RS.Worksheets("Near Dups").Range("A" & i).Value2 <> vbNullString Then
'        RS.Worksheets("CATS_FILE").Cells(RS.Worksheets("Near Dups").Range("F" & i).Value2, 10).Clear
'    End If
'
'Next i
'
''Caption Delete Mark
'lastrow = RS.Worksheets("Caption Header Dups").Range("F" & Rows.Count).End(xlUp).Row
'
'For i = 2 To lastrow
'
'    If RS.Worksheets("Caption Header Dups").Range("A" & i).Value2 <> vbNullString Then
'        j = RS.Worksheets("Caption Header Dups").Range("G" & i).Value2
'        RS.Worksheets("CATS_FILE").Cells(j, 10).Clear
'        j = j + 1
'
'        Do While RS.Worksheets("CATS_FILE").Range("T" & j).Value2 <> 0 And RS.Worksheets("CATS_FILE").Range("T" & j).Value2 <> 10
'            RS.Worksheets("CATS_FILE").Cells(j, 10).Clear
'            j = j + 1
'        Loop
'    End If
'
'Next i
'
''DELETING
'lastrow = RS.Worksheets("CATS_FILE").Range("D" & Rows.Count).End(xlUp).Row
'
'For i = lastrow To 2 Step -1
'
'    If RS.Worksheets("CATS_FILE").Range("J" & i).Value2 = vbNullString Then
'        RS.Worksheets("CATS_FILE").Rows(i).Delete
'    End If
'
'Next i
'
'
'
'
'End Sub
Sub Save_Report()


Dim ws As Worksheet, lastrow As Long, RS As Workbook, NF As Workbook, i As Long, j As Long

Set RS = ThisWorkbook

'RS.Worksheets("Dataset").Visible = xlSheetVeryHidden

Set NF = Workbooks(Workbooks.count)
NF.Activate
Application.DisplayAlerts = False

Do
    fname = Application.GetSaveAsFilename(InitialFileName:="Duplicate_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2, _
        fileFilter:="Excel Files (*.xlsm), *.xlsm", Title:="Save Duplicate Report As:")
    MsgBox fname
Loop Until fname <> False

    RS.SaveAs filename:=fname, FileFormat:=52
    Application.DisplayAlerts = True



End Sub

' ******OLD SUB TO SAVE******
'Sub Save_CATS()
'
'
'Dim WS As Worksheet, lastrow As Long, RS As Workbook, NF As Workbook, i As Long, j As Long
'
'Set RS = ThisWorkbook
'
'RS.Worksheets("CATS_FILE").Visible = xlSheetVisible
'RS.Worksheets("CATS_FILE").Copy
'
'Set NF = Workbooks(Workbooks.Count)
'NF.Activate
'Application.ScreenUpdating = False
'Application.DisplayAlerts = False
'
'fname = "\\172.28.136.38\DominicanRepublic-Shares\Dept\FO-DBM-Annual Load Files\--@Annual Load Duplicate Cleanup--\" & _
'    RS.Worksheets("Dataset").Cells(1, 26).Value2 & "."
'    NF.SaveAs filename:=fname, FileFormat:=20
'    NF.Close savechanges = False
'    RS.Worksheets("CATS_FILE").Visible = xlSheetVeryHidden
''    RS.SaveCopyAs "\\172.28.136.38\DominicanRepublic-Shares\Dept\FO-DBM-Annual Load Files\--@Annual Load Duplicate Cleanup--\" & _
''    "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2 & ".xlsx"
'
'
'
'RS.Sheets(Array("Automatic Removal Dups", "Near Dups", "Caption Header Dups")).Copy
'
'Set NF = Workbooks(Workbooks.Count)
'NF.Activate
'Application.DisplayAlerts = False
'fname = "\\172.28.136.38\DominicanRepublic-Shares\Dept\FO-DBM-Annual Load Files\--@Annual Load Duplicate Cleanup--\" & _
'    "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2
'    NF.SaveAs filename:=fname, FileFormat:=51
'    NF.Close savechanges = False
'
'
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True
'
'
'
'MsgBox "Finished"
'
'End Sub


Sub Save_CATS()


Dim ws As Worksheet, lastrow As Long, RS As Workbook, NF As Workbook, i As Long, j As Long

Set RS = ThisWorkbook


lastrow = RS.Worksheets("Automatic Removal Dups").Range("B" & Rows.count).End(xlUp).row

RS.Worksheets("CATS_FILE").Visible = xlSheetVisible
RS.Worksheets("CATS_FILE").Copy

Set NF = Workbooks(Workbooks.count)
NF.Activate
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Automatic Dup Delete Mark

For i = 2 To lastrow

    If RS.Worksheets("Automatic Removal Dups").Range("A" & i).Value2 <> vbNullString Then
        NF.Worksheets(1).Cells(RS.Worksheets("Automatic Removal Dups").Range("F" & i).Value2, 10).Clear
    End If

Next i


'Near Dup Delete Mark
lastrow = RS.Worksheets("Near Dups").Range("B" & Rows.count).End(xlUp).row

For i = 2 To lastrow

    If RS.Worksheets("Near Dups").Range("A" & i).Value2 <> vbNullString Then
        NF.Worksheets(1).Cells(RS.Worksheets("Near Dups").Range("F" & i).Value2, 10).Clear
    End If

Next i

'Caption Delete Mark
lastrow = RS.Worksheets("Caption Header Dups").Range("F" & Rows.count).End(xlUp).row

For i = 2 To lastrow

    If RS.Worksheets("Caption Header Dups").Range("A" & i).Value2 <> vbNullString Then
        j = RS.Worksheets("Caption Header Dups").Range("G" & i).Value2
        NF.Worksheets(1).Cells(j, 10).Clear
        j = j + 1
        
        Do While RS.Worksheets("CATS_FILE").Range("T" & j).Value2 <> 0 And RS.Worksheets("CATS_FILE").Range("T" & j).Value2 <> 10
            NF.Worksheets(1).Cells(j, 10).Clear
            j = j + 1
        Loop
    End If

Next i

'DELETING
lastrow = NF.Worksheets(1).Range("D" & Rows.count).End(xlUp).row

For i = lastrow To 2 Step -1

    If NF.Worksheets(1).Range("J" & i).Value2 = vbNullString Then
        NF.Worksheets(1).Rows(i).Delete
    End If

Next i

'fvs : Remove vivial force file and leave the thryv file
If NF.Worksheets(1).Range("AR2").Value <> vbNullString Then
    NF.Worksheets(1).Columns("A:AQ").Delete
End If

'For Tech team Debugging
'fname = "C:\Users\dc2462\OneDrive - Vivial\Documents\Excel VBA\VBA Projects\Windstream Annual Converter\WIN\" & _
'    RS.Worksheets("Dataset").Cells(1, 26).Value2 & "."
'fname = "C:\Users\dc2462\OneDrive - Vivial\Documents\Excel VBA\VBA Projects\" & _
    RS.Worksheets("Dataset").Cells(1, 26).Value2 & "."
'fname = "https://vivialnet.sharepoint.com/sites/DR-ServiceOrders/FODBMAnnual Load Files/--@Annual Load Duplicate Cleanup--/" & _
'    RS.Worksheets("Dataset").Cells(1, 26).Value2 & "."

'Upload and file at time:
If css_file = True Then
    UploadFiles
End If

NF.Worksheets("CATS_FILE").Activate

'Working code fvs
fname = "https://thryv.sharepoint.com/sites/DR-ServiceOrders/FODBMAnnual Load Files/--@Annual Load Duplicate Cleanup--/" & _
    "MERGED_" & RS.Worksheets("Dataset").Cells(1, 26).Value2 & "."
    NF.Worksheets("CATS_FILE").SaveAs filename:=fname, FileFormat:=20
    MsgBox fname
    NF.Close savechanges = False
    RS.Worksheets("CATS_FILE").Visible = xlSheetVeryHidden
    
'    RS.SaveCopyAs "\\172.28.136.38\DominicanRepublic-Shares\Dept\FO-DBM-Annual Load Files\--@Annual Load Duplicate Cleanup--\" & _
'    "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2 & ".xlsx"
    
RS.Sheets(Array("Automatic Removal Dups", "Near Dups", "Caption Header Dups")).Copy

Set NF = Workbooks(Workbooks.count)
NF.Activate
Application.DisplayAlerts = False

'For Tech Team debugging
'fname = "C:\Users\dc2462\OneDrive - Vivial\Documents\Excel VBA\VBA Projects\Windstream Annual Converter\WIN\" & _
'    "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2
'fname = "C:\Users\dc2462\OneDrive - Vivial\Documents\Excel VBA\VBA Projects\" & _
'    "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2
'fname = "https://vivialnet.sharepoint.com/sites/DR-ServiceOrders/FODBMAnnual Load Files/--@Annual Load Duplicate Cleanup--/" & _
'    "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2
fname = "https://thryv.sharepoint.com/sites/DR-ServiceOrders/FODBMAnnual Load Files/--@Annual Load Duplicate Cleanup--/" & _
    "Duplicates_Report_" & RS.Worksheets("Dataset").Cells(1, 26).Value2
    'fvs: Replace added to remove thryv files extention .txt in order to convert the file to the .xlsx format.
    NF.SaveAs filename:=Replace(fname, ".txt", ""), FileFormat:=51
    MsgBox fname
    NF.Close savechanges = False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RS.Worksheets("Dataset").Visible = xlSheetVeryHidden
    RS.Worksheets("CATS_FILE").Visible = xlSheetVeryHidden

MsgBox "Finished"

ThisWorkbook.Save

End Sub


