Option Explicit

Public Sub UnitTest_clsDataFile_Object()

    Dim objTest As clsDataFile
    Set objTest = New clsDataFile
    
    Debug.Assert True

End Sub




Public Sub UnitTest_clsDataFile_Filepaths()

    Dim objTest As clsDataFile
    Dim i As Integer
    
    For i = 3008 To 3008
        Set objTest = New clsDataFile
        objTest.LoadFile "FOV", Trim(Str(i))
        Debug.Print objTest.FilePath
        Set objTest = Nothing
    Next

End Sub




Public Sub UnitTest_clsDataFile_Sheets()

    Dim objTest As clsDataFile
    Dim i As Integer
    
    For i = 3008 To 3008
        Set objTest = New clsDataFile
        objTest.LoadFile "FOV", Trim(Str(i))
        Debug.Print objTest.FilePath
        Debug.Print objTest.WSheet.Name
        Debug.Print Trim(objTest.WSheet.Range("E11"))
        Debug.Print ""
        Set objTest = Nothing
    Next

End Sub




Public Sub UnitTest_clsDataFile_Inputs()

    On Error GoTo ErrorHandler
    
    Dim objTest As clsDataFile
    
    Set objTest = New clsDataFile
    objTest.LoadFile "FOVm", "1001"
    Debug.Print Now
    Debug.Print "FilePath: " & objTest.FilePath
    Debug.Print "Partition: " & objTest.Partition & "FOVm"
    Debug.Print "UnitNumber: " & objTest.UnitNumber & "1001"
    Debug.Print "NoPartition_Error: " & objTest.NoPartition_Error
    Debug.Print "NoUnitNumber_Error: " & objTest.NoUnitNumber_Error
    Debug.Print "NoDir_Error: " & objTest.NoDir_Error
    Debug.Print "NoFile_Error: " & objTest.NoFile_Error
    Debug.Print ""
    Debug.Assert objTest.NoPartition_Error = True And objTest.NoUnitNumber_Error = False And _
                 objTest.NoDir_Error = False And objTest.NoFile_Error = False
    Set objTest = Nothing
    
    Set objTest = New clsDataFile
    objTest.LoadFile "FOV", "100"
    Debug.Print Now
    Debug.Print "FilePath: " & objTest.FilePath
    Debug.Print "Partition: " & objTest.Partition & "FOV"
    Debug.Print "UnitNumber: " & objTest.UnitNumber & "100"
    Debug.Print "NoPartition_Error: " & objTest.NoPartition_Error
    Debug.Print "NoUnitNumber_Error: " & objTest.NoUnitNumber_Error
    Debug.Print "NoDir_Error: " & objTest.NoDir_Error
    Debug.Print "NoFile_Error: " & objTest.NoFile_Error
    Debug.Print ""
    Debug.Assert objTest.NoPartition_Error = False And objTest.NoUnitNumber_Error = True And _
                 objTest.NoDir_Error = False And objTest.NoFile_Error = False
    Set objTest = Nothing
    
    Set objTest = New clsDataFile
    objTest.LoadFile "FOV", "100m"
    Debug.Print Now
    Debug.Print "FilePath: " & objTest.FilePath
    Debug.Print "Partition: " & objTest.Partition & "FOV"
    Debug.Print "UnitNumber: " & objTest.UnitNumber & "100m"
    Debug.Print "NoPartition_Error: " & objTest.NoPartition_Error
    Debug.Print "NoUnitNumber_Error: " & objTest.NoUnitNumber_Error
    Debug.Print "NoDir_Error: " & objTest.NoDir_Error
    Debug.Print "NoFile_Error: " & objTest.NoFile_Error
    Debug.Print ""
    Debug.Assert objTest.NoPartition_Error = False And objTest.NoUnitNumber_Error = True And _
                 objTest.NoDir_Error = False And objTest.NoFile_Error = False
    Set objTest = Nothing
    
    Set objTest = New clsDataFile
    objTest.LoadFile "FOV", "1000"
    Debug.Print Now
    Debug.Print "FilePath: " & objTest.FilePath
    Debug.Print "Partition: " & objTest.Partition & "FOV"
    Debug.Print "UnitNumber: " & objTest.UnitNumber & "1000"
    Debug.Print "NoPartition_Error: " & objTest.NoPartition_Error
    Debug.Print "NoUnitNumber_Error: " & objTest.NoUnitNumber_Error
    Debug.Print "NoDir_Error: " & objTest.NoDir_Error
    Debug.Print "NoFile_Error: " & objTest.NoFile_Error
    Debug.Print ""
    Debug.Assert objTest.NoPartition_Error = False And objTest.NoUnitNumber_Error = False And _
                 objTest.NoDir_Error = True And objTest.NoFile_Error = False
    Set objTest = Nothing
    
    Set objTest = New clsDataFile
    objTest.LoadFile "FOV", "1001"
    Debug.Print Now
    Debug.Print "FilePath: " & objTest.FilePath
    Debug.Print "Partition: " & objTest.Partition & "FOV"
    Debug.Print "UnitNumber: " & objTest.UnitNumber & "1001"
    Debug.Print "NoPartition_Error: " & objTest.NoPartition_Error
    Debug.Print "NoUnitNumber_Error: " & objTest.NoUnitNumber_Error
    Debug.Print "NoDir_Error: " & objTest.NoDir_Error
    Debug.Print "NoFile_Error: " & objTest.NoFile_Error
    Debug.Print ""
    Debug.Assert objTest.NoPartition_Error = False And objTest.NoUnitNumber_Error = False And _
                 objTest.NoDir_Error = False And objTest.NoFile_Error = True
    Set objTest = Nothing
    
    Set objTest = New clsDataFile
    objTest.LoadFile "FOV", "3004"
    Debug.Print Now
    Debug.Print "FilePath: " & objTest.FilePath
    Debug.Print "Partition: " & objTest.Partition & "FOV"
    Debug.Print "UnitNumber: " & objTest.UnitNumber & "3004"
    Debug.Print "NoPartition_Error: " & objTest.NoPartition_Error
    Debug.Print "NoUnitNumber_Error: " & objTest.NoUnitNumber_Error
    Debug.Print "NoDir_Error: " & objTest.NoDir_Error
    Debug.Print "NoFile_Error: " & objTest.NoFile_Error
    Debug.Print ""
    Debug.Assert objTest.NoPartition_Error = False And objTest.NoUnitNumber_Error = False And _
                 objTest.NoDir_Error = False And objTest.NoFile_Error = False
    Set objTest = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error" & Err.Number & ": " & Err.Description & " in Testing.UnitTest_clsDataFile_Inputs()"

End Sub
