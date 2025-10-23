Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Sub OpenSolidWorksFiles()
    
    Dim folderPath As String
    Dim partName As String
    Dim asmPath As String
    Dim drwPath As String
    Dim row As Integer
    
    ' Connect to existing SolidWorks session
    On Error Resume Next
    Set swApp = GetObject(, "SldWorks.Application")
    On Error GoTo 0
    
    If swApp Is Nothing Then
        MsgBox "SolidWorks is not running. Please open SolidWorks and try again.", vbCritical
        Exit Sub
    End If
    
    swApp.Visible = True
    
    ' Get folder path from A1
    folderPath = Trim(ThisWorkbook.Sheets(1).Range("A1").Value)
    
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Loop through part names from A3 to A6
    For row = 3 To 6
        
        partName = Trim(ThisWorkbook.Sheets(1).Cells(row, 1).Value)
        
        If partName <> "" Then
            
            asmPath = folderPath & partName & ".sldasm"
            drwPath = folderPath & partName & ".slddrw"
            
            ' Check if Assembly file exists and open
            If FileExists(asmPath) Then
                Set swModel = swApp.OpenDoc6(asmPath, swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
            Else
                MsgBox "Assembly file not found: " & asmPath, vbExclamation
            End If
            
            ' Check if Drawing file exists and open
            If FileExists(drwPath) Then
                Set swModel = swApp.OpenDoc6(drwPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
            Else
                MsgBox "Drawing file not found: " & drwPath, vbExclamation
            End If
            
        End If
        
    Next row
    
    MsgBox "Process Completed.", vbInformation

End Sub

' Function to check if file exists
Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function


Abhishek

