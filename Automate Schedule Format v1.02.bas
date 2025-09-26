Attribute VB_Name = "modAutomateScheduleFormat"
'Program: Automate Schedule Format'
'Developed by Alejandro Cano'
'Date: 9/22/2025'
'ver 1.02'

'Goal: Completely automate the scheduling format downloaded from Appointy.'
'after downloading the file and running the program, the file will separate Online students from in-center students,
'and then separate incenter students based on the start time of the session. Additionally
'it will change the font of students that are scheduled for 90min to red and 30 min to blue for improved readability







'Subroutine for readability
Sub ADJUST_COLUMNS()
    Columns("F").Delete
    Columns("A:G").EntireColumn.AutoFit
End Sub
'Subroutine to sort incenter from online students
Sub SORT_LOCATIONNAME()

    Columns("A:E").Sort key1:=Columns("A"), _
    order1:=xlDescending, Header:=xlYes
End Sub
'Subroutine to rename online students for readability'
Sub SEPARATE_TIME()
    Dim i As Integer
    Dim homeflag As Boolean
    homeflag = True
    i = 0
    
    While homeflag = True
        If Cells(2 + i, 1).Value = "Schaumburg@Home" Then
            Cells(2 + i, 1).Value = "Home"
            i = i + 1
        Else
            homeflag = False
        End If
    Wend
End Sub

'main'
Sub AutomateScheduleFormat_main()
Dim i, j, k As Integer
Dim emptyflag As Boolean
Dim locationHandle As String
Dim durationHandle As String
Dim dateHandle As String
Dim homeAmount As Integer
Dim centerAmount As Integer
Dim prevDate As Boolean
Dim prevdateHandle As String

'Subroutines to improve readability, and prepare for main function
Call ADJUST_COLUMNS
Call SORT_LOCATIONNAME
Call SEPARATE_TIME

emptyflag = False
i = 0
homeAmount = 0
centerAmount = 0

'Determines which students are online, and which ones are in center for 30,60 or 90 min'
While emptyflag = False
    
    locationHandle = Cells(2 + i, 1)
    durationHandle = Cells(2 + i, 5)
    
    Select Case locationHandle
    Case "Home"
        homeAmount = homeAmount + 1
    
    Case "Schaumburg"
        centerAmount = centerAmount + 1
        If durationHandle = "90m" Then
            For j = 1 To 5
                Cells(2 + i, j).Font.Color = -16776961
            Next
        ElseIf durationHandle = "30m" Then
            For j = 1 To 5
                Cells(2 + i, j).Font.Color = RGB(0, 248, 242)
            Next
        
        End If
        
    End Select
    
    i = i + 1
    emptyflag = IsEmpty(Cells(2 + i, 1).Value)
    
Wend


Rows(2 + homeAmount).Insert (xlShiftDown)

emptyflag = False
k = 0


'separates home students by arrival time'
While emptyflag = False
        dateHandle = Cells(4 + homeAmount + k, 2).Value
        prevdateHandle = Cells(3 + homeAmount + k, 2).Value
        prevDate = Not (IsEmpty(Cells(3 + homeAmount + k, 2)))
        If dateHandle <> prevdateHandle And prevDate Then
            Rows(4 + homeAmount + k).Insert (x1Shiftdown)
        End If
        k = k + 1
        emptyflag = IsEmpty(Cells(4 + homeAmount + k, 2))
Wend
    'formatting
    Rows(2).Insert (x1Shiftdown)
    'convenience msg displaying an overall amount of students scheduled in center and online
    MsgBox ("The amount of students at home is: " + CStr(homeAmount) + " and the amount of students at the center is : " + CStr(centerAmount))

End Sub
