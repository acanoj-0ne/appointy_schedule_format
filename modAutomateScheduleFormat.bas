Attribute VB_Name = "modAutomateScheduleFormat"
'Program: Automate Schedule Format'
'Developed by Alejandro Cano'
'Date: 9/22/2025'
'ver 1.03'

'Goal: Completely automate the scheduling format downloaded from Appointy.'
'after downloading the file and running the program, the file will separate Online students from in-center students,
'and then separate incenter students based on the start time of the session. Additionally
'it will change the font of students that are scheduled for 90min to red and 30 min to blue for improved readability

'It is implemented so that every center can use it, as it reads a data file to get names from the appointy download, and then changes it to a value specified by the user.

'main'
Sub AutomateScheduleFormat()
Dim i, j, k As Integer
Dim emptyflag As Boolean
Dim locationHandle As String
Dim durationHandle As String
Dim dateHandle As String
Dim homeAmount As Integer
Dim centerAmount As Integer
Dim prevDate As Boolean
Dim prevdateHandle As String
Dim centerNames As Variant

centerNames = GetCenterNames()
'centerNames(0) is name of center from appointy
'centerNames(1) is name of online center from appointy
'centerNames(2) is name of center to be printed
'centerNames(3) is name of online center to be printed
'


'Subroutines to improve readability, and prepare for main function
Call ADJUST_COLUMNS

Call SORT_LOCATIONNAME

If centerNames(1) <> centerNames(3) Then
    Call RENAME_ONLINE(centerNames(1), centerNames(3))
End If



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
    
    Case centerNames(0)
        centerAmount = centerAmount + 1
        Cells(2 + i, 1).Value = centerNames(2)
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
