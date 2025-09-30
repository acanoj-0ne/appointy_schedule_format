Attribute VB_Name = "modReadabilityFunctions"
Option Explicit





'function for readability
Function ADJUST_COLUMNS()
    Columns("F").Delete
    Columns("A:G").EntireColumn.AutoFit
End Function
'function to sort incenter from online students
Function SORT_LOCATIONNAME()

    Columns("A:E").Sort key1:=Columns("A"), _
    order1:=xlDescending, Header:=xlYes
End Function
'function to rename online students for readability'
Function RENAME_ONLINE(ByVal nameFromFile As String, ByVal nameToPrint As String)
    Dim i As Integer
    Dim homeflag As Boolean
    homeflag = True
    i = 0
    
    While homeflag = True
        If Cells(2 + i, 1).Value = nameFromFile Then
            Cells(2 + i, 1).Value = nameToPrint
            i = i + 1
        Else
            homeflag = False
        End If
    Wend
End Function

