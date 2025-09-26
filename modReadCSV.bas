Attribute VB_Name = "modReadCSV"
Option Explicit


'Pending'
Sub ReadCSVtoArray(ByVal filePath As String)
    Dim fso As Object
    Dim ts As Object
    Dim fileLine As String
    Dim i As LoadPictureConstants
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '1 is reading mode
    Set tso = fso.OpenTextFile(filePath, 1)
    
End Sub

