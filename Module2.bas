Attribute VB_Name = "Module2"
Sub Report_Build(fname As Variant)
Attribute Report_Build.VB_ProcData.VB_Invoke_Func = "R\n14"

Dupe_Dataset (fname)
FindDuplicates
Duplicates_Mark_Automatic_Removals
Report_Automatic_Removal
Duplicates_Mark_Near
Report_Near_Dups
Duplicates_Mark_Captions
Report_Captions

'Save_Report

End Sub


Sub Delete_Save()

'Deleting_Process
Save_CATS

End Sub
