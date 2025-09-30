Option Explicit


'Function to take a CSV file and turn it into an array'
Function ReadCSVtoArray(ByVal filePath As String) As Variant

    Dim fso As Object
    Dim ts As Object
    Dim fileLine As String
    Dim dataEntries As Variant
    Dim i As Long
    
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '1 is reading mode
    Set ts = fso.OpenTextFile(filePath, 1)
    fileLine = ts.ReadLine
    dataEntries = Split(fileLine, ", ")

    ReadCSVtoArray = dataEntries
    
    
    
End Function

'function to get the names of the center found in the Center Names file.
Function GetCenterNames() As Variant
Dim fileAddress As String
Dim wb As Workbook
Dim centerNames As Variant
Dim i As Integer

'wb in PERSONAL.XLSB
Set wb = ThisWorkbook

'Relative path
fileAddress = wb.Path + "\data\Center Names.csv"
centerNames = ReadCSVtoArray(fileAddress)

GetCenterNames = centerNames
End Function
