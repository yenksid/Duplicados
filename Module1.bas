Attribute VB_Name = "Module1"
'Option Private Module
Function IsFileOpen(filename As Variant)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
     'Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

     'Check to see which error occurred.
    Select Case errnum

         'No error occurred.
         'File is NOT already open by another user.
        Case 0
         IsFileOpen = False

         'Error number for "Permission Denied."
         'File is already opened by another user.
        Case 70
            IsFileOpen = True

         'Another error occurred.
        Case Else
            Error errnum
    End Select
End Function
Function Section_And_Phone_Dup(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook
    If UCase(RS.Worksheets("Dataset").Cells(i, "E").Value2) & _
        UCase(RS.Worksheets("Dataset").Cells(i, "M").Value2) = _
        UCase(RS.Worksheets("Dataset").Cells(j, "E").Value2) & _
        UCase(RS.Worksheets("Dataset").Cells(j, "M").Value2) Then
            Section_And_Phone_Dup = True
        Else
            Section_And_Phone_Dup = False
        End If
End Function
Function Bus_Res_Gov(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook
    If RS.Worksheets("Dataset").Cells(i, "A") = "R" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "R" Then
            Bus_Res_Gov = "Res_vs_Res"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "B" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "B" Then
            Bus_Res_Gov = "Bus_vs_Bus"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "B" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "R" Then
            Bus_Res_Gov = "Bus_vs_Res"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "R" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "B" Then
            Bus_Res_Gov = "Res_vs_Bus"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "B" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "G" Then
            Bus_Res_Gov = "Bus_vs_Gov"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "G" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "B" Then
            Bus_Res_Gov = "Gov_vs_Bus"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "G" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "G" Then
            Bus_Res_Gov = "Gov_vs_Gov"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "G" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "R" Then
            Bus_Res_Gov = "Gov_vs_Res"
            GoTo BRGResolved
    End If
    If RS.Worksheets("Dataset").Cells(i, "A") = "R" And _
        RS.Worksheets("Dataset").Cells(j, "A") = "G" Then
            Bus_Res_Gov = "Res_vs_Gov"
            GoTo BRGResolved
    End If
    
BRGResolved:
End Function
Function FName_Exact_Dup(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook
    If UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2) = _
        UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2) Then
            FName_Exact_Dup = True
        Else
            FName_Exact_Dup = False
        End If
End Function
Function LName_Exact_Dup(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook
    If UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2) = _
        UCase(RS.Worksheets("Dataset").Cells(j, "H").Value2) Then
            LName_Exact_Dup = True
        Else
            LName_Exact_Dup = False
        End If
End Function
'change log 1.4 matching of names where only one word is different
Function one_word_mismatch(i As Long, j As Long) As Boolean
Dim RS As Workbook, count As Integer
Set RS = ThisWorkbook

one_word_mismatch = False

count = 0
i_name = Split(UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2))
j_name = Split(UCase(RS.Worksheets("Dataset").Cells(j, "H").Value2))

If Application.CountA(i_name) = Application.CountA(j_name) And Application.CountA(i_name) > 1 Then
    For x = 0 To Application.CountA(i_name) - 1 Step 1
        If i_name(x) <> j_name(x) Then count = count + 1
    Next x
    If count = 1 Then one_word_mismatch = True
End If


End Function
Function DataFeedType_Telco_Clec(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

 If RS.Worksheets("Dataset").Cells(i, "B").Value2 = RS.Worksheets("Dataset").Cells(j, "B").Value2 And _
    ((RS.Worksheets("Dataset").Cells(i, "C") = vbNullString And RS.Worksheets("Dataset").Cells(j, "C") = vbNullString) Or _
    (RS.Worksheets("Dataset").Cells(i, "C") <> vbNullString And RS.Worksheets("Dataset").Cells(j, "C") <> vbNullString)) Then
    DataFeedType_Telco_Clec = "Same_Type"
    GoTo ResolvedData
 End If
 If RS.Worksheets("Dataset").Cells(i, "B").Value2 = RS.Worksheets("Dataset").Cells(j, "B").Value2 And _
    RS.Worksheets("Dataset").Cells(i, "C") = vbNullString And RS.Worksheets("Dataset").Cells(j, "C") <> vbNullString Then
    DataFeedType_Telco_Clec = "CLEC_vs_LOCAL"
    GoTo ResolvedData
 End If
 If RS.Worksheets("Dataset").Cells(i, "B").Value2 = RS.Worksheets("Dataset").Cells(j, "B").Value2 And _
    RS.Worksheets("Dataset").Cells(i, "C") <> vbNullString And RS.Worksheets("Dataset").Cells(j, "C") = vbNullString Then
    DataFeedType_Telco_Clec = "LOCAL_vs_CLEC"
    GoTo ResolvedData
 End If
 If RS.Worksheets("Dataset").Cells(i, "B").Value2 = "Annual" And RS.Worksheets("Dataset").Cells(i, "C") <> vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "B").Value2 = "EAS" Then
    DataFeedType_Telco_Clec = "LOCAL_vs_EAS"
    GoTo ResolvedData
 End If
 If RS.Worksheets("Dataset").Cells(i, "B").Value2 = "Annual" And RS.Worksheets("Dataset").Cells(i, "C") = vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "B").Value2 = "EAS" Then
    DataFeedType_Telco_Clec = "CLEC_vs_EAS"
    GoTo ResolvedData
 End If
 If RS.Worksheets("Dataset").Cells(i, "B").Value2 = "EAS" And _
    RS.Worksheets("Dataset").Cells(j, "B").Value2 = "Annual" And RS.Worksheets("Dataset").Cells(j, "C") <> vbNullString Then
    DataFeedType_Telco_Clec = "EAS_vs_LOCAL"
    GoTo ResolvedData
 End If
 If RS.Worksheets("Dataset").Cells(i, "B").Value2 = "EAS" And _
    RS.Worksheets("Dataset").Cells(j, "B").Value2 = "Annual" And RS.Worksheets("Dataset").Cells(j, "C") = vbNullString Then
    DataFeedType_Telco_Clec = "EAS_vs_CLEC"
 End If
 
ResolvedData:
End Function
Function Caption_Indicator(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

 If RS.Worksheets("Dataset").Cells(i, "F").Value2 = 0 And RS.Worksheets("Dataset").Cells(j, "F").Value2 = 0 Then
    Caption_Indicator = "Str_vs_Str"
    GoTo ResolvedIndicator
 End If
 If RS.Worksheets("Dataset").Cells(i, "F").Value2 = 0 And RS.Worksheets("Dataset").Cells(j, "F").Value2 <> 0 Then
    Caption_Indicator = "Str_vs_Cap"
    GoTo ResolvedIndicator
 End If
 If RS.Worksheets("Dataset").Cells(i, "F").Value2 <> 0 And RS.Worksheets("Dataset").Cells(j, "F").Value2 = 0 Then
    Caption_Indicator = "Cap_vs_Str"
    GoTo ResolvedIndicator
 End If
 If RS.Worksheets("Dataset").Cells(i, "F").Value2 <> 0 And RS.Worksheets("Dataset").Cells(j, "F").Value2 <> 0 Then
    Caption_Indicator = "Cap_vs_Cap"
 End If
 
' 'fvs added the last stament
' If RS.Worksheets("Dataset").Cells(i, "F").Value2 <> 0 And RS.Worksheets("Dataset").Cells(j, "F").Value2 <> 0 And RS.Worksheets("Dataset").Cells(i, "O").Value2 <> vbNullString Then
'    Caption_Indicator = "Cap_vs_Cap"
' End If

ResolvedIndicator:
End Function
Function Street_Number(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

 If UCase(RS.Worksheets("Dataset").Cells(i, "J").Value2) = UCase(RS.Worksheets("Dataset").Cells(j, "J").Value2) Then
    Street_Number = "Keep_Either"
    GoTo ResolvedSNumber
 End If
 If RS.Worksheets("Dataset").Cells(i, "J").Value2 <> vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "J").Value2 = vbNullString Then
    Street_Number = "YesTop_NoBottom"
    GoTo ResolvedSNumber
 End If
 If RS.Worksheets("Dataset").Cells(i, "J").Value2 = vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "J").Value2 <> vbNullString Then
    Street_Number = "NoTop_YesBottom"
    GoTo ResolvedSNumber
 End If
 If Street_Number_Cleanup(RS.Worksheets("Dataset").Cells(i, "J").Value2) = Street_Number_Cleanup(RS.Worksheets("Dataset").Cells(j, "J").Value2) Then
    Street_Number = "Number_Cleanup_Match"
 Else
 
    Street_Number = "Keep_Both"
 End If

ResolvedSNumber:
End Function
Function Street_Number_Cleanup(street As String)

Street_Number_Cleanup = Replace(street, "-", "")
Street_Number_Cleanup = Replace(Street_Number_Cleanup, "W", "")
Street_Number_Cleanup = Replace(Street_Number_Cleanup, "N", "")
Street_Number_Cleanup = Replace(Street_Number_Cleanup, "S", "")
Street_Number_Cleanup = Replace(Street_Number_Cleanup, "E", "")
Street_Number_Cleanup = Replace(Street_Number_Cleanup, " ", "")
If IsNumeric(Street_Number_Cleanup) = True Then Street_Number_Cleanup = Val(Street_Number_Cleanup)


End Function

Function Street_Name(i As Long, j As Long)
Dim RS As Workbook, cleaned_addresses As Variant
Set RS = ThisWorkbook

 If UCase(RS.Worksheets("Dataset").Cells(i, "K").Value2) = UCase(RS.Worksheets("Dataset").Cells(j, "K").Value2) Then
    Street_Name = "Keep_Either"
    GoTo ResolvedSName
 End If
 
 'fvs: mark as review_address to be added to the near dups Ex : 880 Peters Creek Pkwy vs 880 Peterscreek Pkwys
 If Replace(Replace(Replace(UCase(RS.Worksheets("Dataset").Cells(i, "K").Value2), " ", ""), "/", ""), "-", "") = Replace(Replace(Replace(UCase(RS.Worksheets("Dataset").Cells(j, "K").Value2), " ", ""), "/", ""), "-", "") Then
    Street_Name = "review_address"
    GoTo ResolvedSName
 End If
 
 If RS.Worksheets("Dataset").Cells(i, "K").Value2 <> vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "K").Value2 = vbNullString Then
    Street_Name = "YesTop_NoBottom"
    GoTo ResolvedSName
 End If
 If RS.Worksheets("Dataset").Cells(i, "K").Value2 = vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "K").Value2 <> vbNullString Then
    Street_Name = "NoTop_YesBottom"
    GoTo ResolvedSName
 End If

'change log 1.4 New NENA logic
 cleaned_addresses = address_cleanup(UCase(RS.Worksheets("Dataset").Cells(i, "K").Value2), UCase(RS.Worksheets("Dataset").Cells(j, "K").Value2))
' MsgBox "streetname: " & i & " - " & j
 'NENA
 If cleaned_addresses(0) = cleaned_addresses(1) Then
    Street_Name = "NENA"
' MsgBox "streetname: " & i & " - " & j & "value:" & Street_Name
    GoTo ResolvedSName
 End If

'  OLD NENA call
' If NENA(UCase(RS.Worksheets("Dataset").Cells(i, "K").Value2)) = NENA(UCase(RS.Worksheets("Dataset").Cells(j, "K").Value2)) Then
'    Street_Name = "NENA"
'    GoTo ResolvedSName
' End If

 'change log 1.3 partial match of address
 If UCase(RS.Worksheets("Dataset").Cells(i, "K").Value2) = UCase(Left(RS.Worksheets("Dataset").Cells(j, "K").Value2, Len(RS.Worksheets("Dataset").Cells(i, "K").Value2))) Or _
    UCase(RS.Worksheets("Dataset").Cells(j, "K").Value2) = UCase(Left(RS.Worksheets("Dataset").Cells(i, "K").Value2, Len(RS.Worksheets("Dataset").Cells(j, "K").Value2))) Then
    Street_Name = "Address_Partial_Match"
    GoTo ResolvedSName
 End If
 
 Street_Name = "Keep_Both"
ResolvedSName:
End Function

'change log 1.4 new NENA logic
Function address_cleanup(i As String, j As String)
count = 0
If i <> vbNullString And j <> vbNullString Then
    iaddr = Split(StrReverse(i))(0)
    jaddr = Split(StrReverse(j))(0)
    
    If iaddr = jaddr Then
        count = count + 1
        If Len(iaddr) < Len(i) Then
            i = Left(UCase(i), Len(i) - Len(iaddr) - 1)
        Else
            i = ""
        End If
        If Len(jaddr) < Len(j) Then
            j = Left(UCase(j), Len(j) - Len(jaddr) - 1)
        Else
            j = ""
        End If
    End If
    
    If count <> 0 Then
        address_cleanup = address_cleanup(i, j)
    End If
    
    address_cleanup = Array(NENA(i), NENA(j))
Else
    address_cleanup = Array(i, j)
End If
End Function


Function NENA(address As String)
count = 0

    If Right(address, 1) = "0" Or _
        Right(address, 1) = "1" Or _
        Right(address, 1) = "2" Or _
        Right(address, 1) = "3" Or _
        Right(address, 1) = "4" Or _
        Right(address, 1) = "5" Or _
        Right(address, 1) = "6" Or _
        Right(address, 1) = "7" Or _
        Right(address, 1) = "8" Or _
        Right(address, 1) = "9" Or _
        Right(address, 1) = " " Then
        address = Left(address, Len(address) - 1)
        count = count + 1
    End If


    If Right(address, 2) = " E" Or _
        Right(address, 2) = " N" Or _
        Right(address, 2) = " S" Or _
        Right(address, 2) = " W" Then
        address = Left(address, Len(address) - 2)
        count = count + 1
    End If
    
    If Right(address, 3) = " AV" Or _
        Right(address, 3) = " CT" Or _
        Right(address, 3) = " CV" Or _
        Right(address, 3) = " DR" Or _
        Right(address, 3) = " LN" Or _
        Right(address, 3) = " NO" Or _
        Right(address, 3) = " PL" Or _
        Right(address, 3) = " RD" Or _
        Right(address, 3) = " ST" Or _
        Right(address, 3) = " SQ" Or _
        Right(address, 3) = " NW" Or _
        Right(address, 3) = " NE" Or _
        Right(address, 3) = " SW" Or _
        Right(address, 3) = " SE" Or _
        Right(address, 3) = " TR" Then
        address = Left(address, Len(address) - 3)
        count = count + 1
    End If

    If Right(address, 4) = " APT" Or _
        Right(address, 4) = " AVE" Or _
        Right(address, 4) = " CRK" Or _
        Right(address, 4) = " CTY" Or _
        Right(address, 4) = " EST" Or _
        Right(address, 4) = " HLS" Or _
        Right(address, 4) = " HWY" Or _
        Right(address, 4) = " MNR" Or _
        Right(address, 4) = " OFC" Or _
        Right(address, 4) = " PLZ" Or _
        Right(address, 4) = " PRK" Or _
        Right(address, 4) = " RDG" Or _
        Right(address, 4) = " STE" Or _
        Right(address, 4) = " TRC" Or _
        Right(address, 4) = " TRL" Or _
        Right(address, 4) = " WAY" Then
        address = Left(address, Len(address) - 4)
        count = count + 1
    End If
    
    If Right(address, 5) = " APTS" Or _
        Right(address, 5) = " BLVD" Or _
        Right(address, 5) = " COVE" Or _
        Right(address, 5) = " EAST" Or _
        Right(address, 5) = " LANE" Or _
        Right(address, 5) = " PARK" Or _
        Right(address, 5) = " PKWY" Or _
        Right(address, 5) = " ROAD" Or _
        Right(address, 5) = " SPUR" Or _
        Right(address, 5) = " TRLR" Or _
        Right(address, 5) = " WEST" Then
        address = Left(address, Len(address) - 5)
        count = count + 1
    End If

    If Right(address, 6) = " CREEK" Or _
        Right(address, 6) = " COURT" Or _
        Right(address, 6) = " DRIVE" Or _
        Right(address, 6) = " HILLS" Or _
        Right(address, 6) = " MANOR" Or _
        Right(address, 6) = " NORTH" Or _
        Right(address, 6) = " PLACE" Or _
        Right(address, 6) = " PLAZA" Or _
        Right(address, 6) = " RIDGE" Or _
        Right(address, 6) = " SOUTH" Or _
        Right(address, 6) = " SPRNG" Or _
        Right(address, 6) = " SUITE" Or _
        Right(address, 6) = " TRACE" Or _
        Right(address, 6) = " TRAIL" Then
        address = Left(address, Len(address) - 6)
        count = count + 1
    End If

    If Right(address, 7) = " AVENUE" Or _
        Right(address, 7) = " COUNTY" Or _
        Right(address, 7) = " ESTATE" Or _
        Right(address, 7) = " NUMBER" Or _
        Right(address, 7) = " OFFICE" Or _
        Right(address, 7) = " STREET" Or _
        Right(address, 7) = " SPRNGS" Or _
        Right(address, 7) = " SPRING" Then
        address = Left(address, Len(address) - 7)
        count = count + 1
    End If
    
    If Right(address, 8) = " ESTATES" Or _
        Right(address, 8) = " HIGHWAY" Or _
        Right(address, 8) = " SPRINGS" Or _
        Right(address, 8) = " PARKWAY" Or _
        Right(address, 8) = " TRAILER" Then
        address = Left(address, Len(address) - 8)
        count = count + 1
    End If
    
    
    If Right(address, 10) = " APARTMENT" Or _
        Right(address, 10) = " BOULEVARD" Or _
        Right(address, 10) = " NORTHWEST" Or _
        Right(address, 10) = " NORTHEAST" Or _
        Right(address, 10) = " SOUTHWEST" Or _
        Right(address, 10) = " SOUTHEAST" Then
        address = Left(address, Len(address) - 10)
        count = count + 1
    End If
    
    If Right(address, 11) = " APARTMENTS" Then
        address = Left(address, Len(address) - 11)
        count = count + 1
    End If
    
    
    If Left(address, 2) = "E " Or _
        Left(address, 2) = "N " Or _
        Left(address, 2) = "S " Or _
        Left(address, 2) = "W " Then
        address = Right(address, Len(address) - 2)
        count = count + 1
    End If
    
    If Left(address, 3) = "NE " Or _
        Left(address, 3) = "NW " Or _
        Left(address, 3) = "SE " Or _
        Left(address, 3) = "SW " Then
        address = Right(address, Len(address) - 3)
        count = count + 1
    End If
    
    If Left(address, 10) = "NORTHWEST " Or _
        Left(address, 10) = "NORTHEAST " Or _
        Left(address, 10) = "SOUTHWEST " Or _
        Left(address, 10) = "SOUTHEAST " Then
        address = Right(address, Len(address) - 10)
        count = count + 1
    End If
    
If count <> 0 Then
    NENA = NENA(address)
End If


'line below is from new code
If InStr(1, address, " ") = 0 Then address = Single_Word_NENA(address)

address = Replace(address, "-", "")
address = Replace(address, "'", "")
address = Replace(address, " ", "")

NENA = address


End Function

'change log 1.4 new NENA Logic
Function Single_Word_NENA(address As String)

If Len(address) = 1 Then
    If Right(address, 1) = "0" Or _
        Right(address, 1) = "1" Or _
        Right(address, 1) = "2" Or _
        Right(address, 1) = "3" Or _
        Right(address, 1) = "4" Or _
        Right(address, 1) = "5" Or _
        Right(address, 1) = "6" Or _
        Right(address, 1) = "7" Or _
        Right(address, 1) = "8" Or _
        Right(address, 1) = "9" Then
        address = Left(address, Len(address) - 1)
    End If
    If Right(address, 1) = "E" Or _
        Right(address, 1) = "N" Or _
        Right(address, 1) = "S" Or _
        Right(address, 1) = "W" Then
        address = Left(address, Len(address) - 1)
    End If
End If
    
If Len(address) = 2 Then
    If Right(address, 2) = "AV" Or _
        Right(address, 2) = "CT" Or _
        Right(address, 2) = "CV" Or _
        Right(address, 2) = "DR" Or _
        Right(address, 2) = "LN" Or _
        Right(address, 2) = "NO" Or _
        Right(address, 2) = "PL" Or _
        Right(address, 2) = "RD" Or _
        Right(address, 2) = "ST" Or _
        Right(address, 2) = "SQ" Or _
        Right(address, 2) = "NW" Or _
        Right(address, 2) = "NE" Or _
        Right(address, 2) = "SW" Or _
        Right(address, 2) = "SE" Or _
        Right(address, 2) = "TR" Then
        address = Left(address, Len(address) - 2)
    End If
End If

If Len(address) = 3 Then
    If Right(address, 3) = "APT" Or _
        Right(address, 3) = "AVE" Or _
        Right(address, 3) = "CRK" Or _
        Right(address, 3) = "CTY" Or _
        Right(address, 3) = "EST" Or _
        Right(address, 3) = "HLS" Or _
        Right(address, 3) = "HWY" Or _
        Right(address, 3) = "MNR" Or _
        Right(address, 3) = "OFC" Or _
        Right(address, 3) = "PLZ" Or _
        Right(address, 3) = "PRK" Or _
        Right(address, 3) = "RDG" Or _
        Right(address, 3) = "STE" Or _
        Right(address, 3) = "TRC" Or _
        Right(address, 3) = "TRL" Or _
        Right(address, 3) = "WAY" Then
        address = Left(address, Len(address) - 3)
    End If
End If
    
If Len(address) = 4 Then
    If Right(address, 4) = "APTS" Or _
        Right(address, 4) = "BLVD" Or _
        Right(address, 4) = "COVE" Or _
        Right(address, 4) = "EAST" Or _
        Right(address, 4) = "LANE" Or _
        Right(address, 4) = "PARK" Or _
        Right(address, 4) = "PKWY" Or _
        Right(address, 4) = "ROAD" Or _
        Right(address, 4) = "SPUR" Or _
        Right(address, 4) = "TRLR" Or _
        Right(address, 4) = "WEST" Then
        address = Left(address, Len(address) - 4)
    End If
End If



If Len(address) = 5 Then
    If Right(address, 5) = "CREEK" Or _
        Right(address, 5) = "COURT" Or _
        Right(address, 5) = "DRIVE" Or _
        Right(address, 5) = "HILLS" Or _
        Right(address, 5) = "MANOR" Or _
        Right(address, 5) = "NORTH" Or _
        Right(address, 5) = "PLACE" Or _
        Right(address, 5) = "PLAZA" Or _
        Right(address, 5) = "RIDGE" Or _
        Right(address, 5) = "SOUTH" Or _
        Right(address, 5) = "SPRNG" Or _
        Right(address, 5) = "SUITE" Or _
        Right(address, 5) = "TRACE" Or _
        Right(address, 5) = "TRAIL" Then
        address = Left(address, Len(address) - 5)
    End If
End If



If Len(address) = 6 Then
    If Right(address, 6) = "AVENUE" Or _
        Right(address, 6) = "COUNTY" Or _
        Right(address, 6) = "ESTATE" Or _
        Right(address, 6) = "NUMBER" Or _
        Right(address, 6) = "OFFICE" Or _
        Right(address, 6) = "STREET" Or _
        Right(address, 6) = "SPRNGS" Or _
        Right(address, 6) = "SPRING" Then
        address = Left(address, Len(address) - 6)
    End If
End If
    
    
If Len(address) = 7 Then
    If Right(address, 7) = "ESTATES" Or _
        Right(address, 7) = "HIGHWAY" Or _
        Right(address, 7) = "SPRINGS" Or _
        Right(address, 7) = "PARKWAY" Or _
        Right(address, 7) = "TRAILER" Then
        address = Left(address, Len(address) - 7)
    End If
End If
    
If Len(address) = 9 Then
    If Right(address, 9) = "APARTMENT" Or _
        Right(address, 9) = "BOULEVARD" Or _
        Right(address, 9) = "NORTHWEST" Or _
        Right(address, 9) = "NORTHEAST" Or _
        Right(address, 9) = "SOUTHWEST" Or _
        Right(address, 9) = "SOUTHEAST" Then
        address = Left(address, Len(address) - 9)
    End If
End If
    
If Len(address) = 10 Then
    If Right(address, 10) = "APARTMENTS" Then
        address = Left(address, Len(address) - 10)
    End If
End If
            

Single_Word_NENA = address


End Function

Function Community(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

 If (RS.Worksheets("Dataset").Cells(i, "L").Value2 <> vbNullString And RS.Worksheets("Dataset").Cells(j, "L").Value2 <> vbNullString) Or _
    (RS.Worksheets("Dataset").Cells(i, "L").Value2 = vbNullString And RS.Worksheets("Dataset").Cells(j, "L").Value2 = vbNullString) Then
    Community = "Keep_Either"
    GoTo ResolvedComm
 End If
  If RS.Worksheets("Dataset").Cells(i, "L").Value2 <> vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "L").Value2 = vbNullString Then
    Community = "YesTop_NoBottom"
    GoTo ResolvedComm
 End If
  If RS.Worksheets("Dataset").Cells(i, "L").Value2 = vbNullString And _
    RS.Worksheets("Dataset").Cells(j, "L").Value2 <> vbNullString Then
    Community = "NoTop_YesBottom"
 End If
 
ResolvedComm:
End Function
Function FName_Split_String(i As Long)
Dim RS As Workbook, FName_Split() As String
Set RS = ThisWorkbook
If RS.Worksheets("Dataset").Cells(i, "G").Value2 <> vbNullString Then
    If InStr(2, UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " & ") <> 0 Then
        FName_Split = Split(UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " & ", 2)
        FName_Split_String = FName_Split(0) & " | " & FName_Split(1)
    Else
        If InStr(2, UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " AND ") <> 0 Then
            FName_Split = Split(UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " AND ", 2)
            FName_Split_String = FName_Split(0) & " | " & FName_Split(1)
        Else
            FName_Split_String = UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2)
        End If
    End If
Else
    FName_Split_String = UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2)
End If

End Function
Function FName_Switch_String(i As Long)
Dim RS As Workbook, FName_Split() As String
Set RS = ThisWorkbook
If RS.Worksheets("Dataset").Cells(i, "G").Value2 <> vbNullString Then
    If InStr(2, UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " & ") <> 0 Then
        FName_Split = Split(UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " & ", 2)
        FName_Switch_String = FName_Split(1) & " | " & FName_Split(0)
    Else
        If InStr(2, UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " AND ") <> 0 Then
            FName_Split = Split(UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " AND ", 2)
            FName_Switch_String = FName_Split(1) & " | " & FName_Split(0)
        Else
            FName_Switch_String = UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2)
        End If
    End If
Else
    FName_Switch_String = UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2)
End If

End Function
Function FName_Switch_Dup(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

If (FName_Split_String(i) = FName_Switch_String(j)) Then
    FName_Switch_Dup = True
    GoTo ResolvedFNameSwitch
End If
If (FName_Split_String(i) = FName_Split_String(j)) Then
    FName_Switch_Dup = True
    GoTo ResolvedFNameSwitch
End If

FName_Switch_Dup = False

ResolvedFNameSwitch:
End Function
Function Empty_Address_Match(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

If RS.Worksheets("Dataset").Cells(i, "L").Value2 = vbNullString And RS.Worksheets("Dataset").Cells(j, "L").Value2 = vbNullString _
    And RS.Worksheets("Dataset").Cells(i, "K").Value2 = vbNullString And RS.Worksheets("Dataset").Cells(j, "K").Value2 = vbNullString _
    And RS.Worksheets("Dataset").Cells(i, "J").Value2 = vbNullString And RS.Worksheets("Dataset").Cells(j, "L").Value2 = vbNullString _
    And RS.Worksheets("Dataset").Cells(i, "I").Value2 = vbNullString And RS.Worksheets("Dataset").Cells(j, "I").Value2 = vbNullString Then
        Empty_Address_Match = True
Else
        Empty_Address_Match = False
End If

End Function
Function Cross_Reference_Lines(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

If UCase(Left(RS.Worksheets("Dataset").Cells(i, "M").Value2, 3)) = "SEE" And UCase(Left(RS.Worksheets("Dataset").Cells(j, "M").Value2, 3)) = "SEE" Then
        Cross_Reference_Lines = True
Else
        Cross_Reference_Lines = False
End If

End Function
Function Partial_FName_Match(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

If Left(UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2), Len(RS.Worksheets("Dataset").Cells(i, "G").Value2)) = UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2) Then
    Partial_FName_Match = True
    GoTo ResolvedPartialFNameMatch
End If
If Left(UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), Len(RS.Worksheets("Dataset").Cells(j, "G").Value2)) = UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2) Then
    Partial_FName_Match = True
    GoTo ResolvedPartialFNameMatch
End If
If Left(FName_Switch_String(j), Len(RS.Worksheets("Dataset").Cells(i, "G").Value2)) = UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2) Then
    Partial_FName_Match = True
    GoTo ResolvedPartialFNameMatch
End If
If Left(FName_Switch_String(i), Len(RS.Worksheets("Dataset").Cells(j, "G").Value2)) = UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2) Then
    Partial_FName_Match = True
    GoTo ResolvedPartialFNameMatch
End If


Partial_FName_Match = False

ResolvedPartialFNameMatch:
End Function
Function Partial_LName_Match(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

If Left(UCase(RS.Worksheets("Dataset").Cells(j, "H").Value2), Len(RS.Worksheets("Dataset").Cells(i, "H").Value2)) = UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2) Then
    Partial_LName_Match = True
    GoTo ResolvedPartialLNameMatch
End If
If Left(UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2), Len(RS.Worksheets("Dataset").Cells(j, "H").Value2)) = UCase(RS.Worksheets("Dataset").Cells(j, "H").Value2) Then
    Partial_LName_Match = True
    GoTo ResolvedPartialLNameMatch
End If

''****need to revise these 2 conditions below... dont make sense in this function
If Left(FName_Switch_String(j), Len(RS.Worksheets("Dataset").Cells(i, "H").Value2)) = UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2) Then
    Partial_LName_Match = True
    GoTo ResolvedPartialLNameMatch
End If
If Left(FName_Switch_String(i), Len(RS.Worksheets("Dataset").Cells(j, "H").Value2)) = UCase(RS.Worksheets("Dataset").Cells(j, "H").Value2) Then
    Partial_LName_Match = True
    GoTo ResolvedPartialLNameMatch
End If

Partial_LName_Match = False

ResolvedPartialLNameMatch:
End Function


'**** this is a test function, not currently used
Function Testfunction(i As Long, j As Long) As Boolean
Dim RS As Workbook, arr1 As Variant, arr2 As Variant
Set RS = ThisWorkbook
Testfun = True
arr1 = Split(RS.Worksheets("Dataset").Cells(i, "H").Value2)
arr2 = Split(RS.Worksheets("Dataset").Cells(j, "H").Value2)

For x = 0 To Application.CountA(arr1) - 1
    For y = 0 To Application.CountA(arr2) - 1
        If arr1(x) = arr2(y) Then Exit For
        If y = Application.CountA(arr2) - 1 Then GoTo jvsi
    Next y
    If x = Application.CountA(arr1) - 1 Then GoTo Definedf
Next x

jvsi:
For x = 0 To Application.CountA(arr2) - 1
    For y = 0 To Application.CountA(arr1) - 1
        If arr2(x) = arr1(y) Then Exit For
        If y = Application.CountA(arr1) - 1 Then
            Testfun = False
            GoTo Definedf
        End If
    Next y
Next x

Definedf:

End Function


Function Space_Partial_FName_Match(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

If InStr(1, UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " ") <> 0 Then
    If Left(UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), InStr(1, UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " ") - 1) = _
        Left(UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2), InStr(1, UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), " ") - 1) Then
            Space_Partial_FName_Match = True
            GoTo ResolvedSpacePartialFNameMatch
    End If
End If

If InStr(1, UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2), " ") <> 0 Then
    If Left(UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2), InStr(1, UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2), " ") - 1) = _
        Left(UCase(RS.Worksheets("Dataset").Cells(i, "G").Value2), InStr(1, UCase(RS.Worksheets("Dataset").Cells(j, "G").Value2), " ") - 1) Then
            Space_Partial_FName_Match = True
            GoTo ResolvedSpacePartialFNameMatch
    End If
End If


Space_Partial_FName_Match = False

ResolvedSpacePartialFNameMatch:
End Function
Function Special_Char_Rem(i As Long)
Dim RS As Workbook, name As String
Set RS = ThisWorkbook

name = UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2)
name = Replace(name, "'", "")
name = Replace(name, "-", "")
name = Replace(name, "&", "")
name = Replace(name, "/", "")
name = Replace(name, " AND ", "")
name = Replace(name, " ", "")

Special_Char_Rem = name
    
End Function
Function Special_Char_Rem_Match(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

'change log 1.3 - updated to match uppercase

If UCase(Special_Char_Rem(i)) = UCase(Special_Char_Rem(j)) Then
    Special_Char_Rem_Match = True
Else
    Special_Char_Rem_Match = False
End If

End Function
Function Dup_BusVsRes(i As Long, j As Long)
Dim RS As Workbook
Set RS = ThisWorkbook

If UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2) = UCase(RS.Worksheets("Dataset").Cells(j, "H").Value2 & " " & RS.Worksheets("Dataset").Cells(j, "G").Value2) Then
    Dup_BusVsRes = True
    GoTo ResolvedDupBusVsRes
End If


If UCase(RS.Worksheets("Dataset").Cells(j, "H").Value2) = UCase(RS.Worksheets("Dataset").Cells(i, "H").Value2 & " " & RS.Worksheets("Dataset").Cells(i, "G").Value2) Then
    Dup_BusVsRes = True
    GoTo ResolvedDupBusVsRes
End If

Dup_BusVsRes = False

ResolvedDupBusVsRes:
End Function









