'=========================================================================================================
'VBA Object Name:   clsLogger
'Filename:          PDELGenerator_Sensor_v##.xlsm
'Date:              June 19, 2014
'Version:           0.9
'Author:            Jeremiah Valenzuela
'
'Description:       This class provides procedures for other classes to send text to a FORM TEXT CONSOLE
'                   or a LOGFILE at a given path. The only inputs for the procedures are text strings. The
'                   folder path for the logfile is set in the 'Class_Initialize()' subroutine and is
'                   pulled from the CONFIGURATION sheet of the workbook running this code.
'
'Code Outline:      Declarations, Properties, Subroutines, Functions
'
'Misc:              UDF (User Defined Function), UDS (User Defined Subroutine)
'=========================================================================================================

'## Force Explicit Variable Declaration
Option Explicit

'## Declaration of Class Variables/Properties
Private pLogFilePath As String  'set in Class_Initialize()

'=========================================================================================================
'## Class_Initialize
'   When this class is instantiated, this subroutine acts as a constructor for the new object. Initial
'   values for class variables and instantiation of objects are to be made here
'=========================================================================================================
Private Sub Class_Initialize()

    On Error GoTo ErrorHandler
    
    '// Set path values from Configuration sheet of PDELGenerator_Sensor_v##.xlsm
    pLogFilePath = ThisWorkbook.Sheets("Configuration").Range("B3").Value
    
    '// Write activity to logfile and console
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error" & Err.Number & " : " & Err.Description & " in clsLogger.Class_Initialize()"
    LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                  " in clsLogger.Class_Initialize()"
    Resume Next

End Sub




'=========================================================================================================
'## LogToConsole UDS
'   Writes out to the Main Form text console for the user to see what the program has just completed.
'=========================================================================================================
Public Sub LogToConsole(ByVal myText As String)

    On Error GoTo ErrorHandler
    
    '// Write to form text console
    'frmMainForm.tbStatus.SelText = Now & " : " & myText & vbCrLf
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error" & Err.Number & " : " & Err.Description & " in clsLogger.LogToConsole()"
    LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                  " in clsLogger.LogToConsole()"
    Resume Next

End Sub




'=========================================================================================================
'## LogToLogfile UDS
'   Writes out to the log file for a journal to review what was done when to what files.
'=========================================================================================================
Public Sub LogToLogfile(ByVal myText As String)

    On Error GoTo ErrorHandler
    
    '//Declare variables
    Dim lLogfile As Long    'Index for FreeFile
    
    '// Get the next free index value for opening a file
    lLogfile = FreeFile
    
    '// Open the specified path and point the 'lLogfile' file index to it. Write to logfile. Close file.
    Open pLogFilePath & "test_PDELGenerator.txt" For Append As #lLogfile
    Print #lLogfile, Now & " : " & myText
    Close #lLogfile
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error" & Err.Number & " : " & Err.Description & " in clsLogger.LogToLogfile()"
    LogToConsole Now & " : Error " & Err.Number & ": " & Err.Description & _
                  " in clsLogger.LogToLogfile()"
    Resume Next

End Sub




'=========================================================================================================
'## LogEverywhere UDS
'   Writes out to the log file and to the main form text console.
'=========================================================================================================
Public Sub LogEverywhere(ByVal myText As String)

    On Error GoTo ErrorHandler
    
    '// Log to both logfile and form text console
    LogToLogfile myText
    LogToConsole myText
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error" & Err.Number & " : " & Err.Description & " in clsLogger.LogEverywhere()"
    LogToConsole Now & " : Error " & Err.Number & ": " & Err.Description & _
                  " in clsLogger.LogEverywhere()"
    Resume Next

End Sub




'=========================================================================================================
'## LogInitialLine UDS
'   Writes out to the log file a new line with a star-bar.
'=========================================================================================================
Public Sub LogInitialLine()

    On Error GoTo ErrorHandler
    
    '// Declare variables
    Dim lLogfile As Long    'Index for FreeFile
    
    '// Get the next free index value for opening a file
    lLogfile = FreeFile
    
    '// Open the specified path and point the 'lLogfile' file index to it. Write to logfile. Close file.
    Open pLogFilePath & "test_PDELGenerator.txt" For Append As #lLogfile
    Print #lLogfile, vbCrLf & "*****************************" & vbCrLf
    Close #lLogfile
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error" & Err.Number & " : " & Err.Description & " in clsLogger.LogInitialLine()"
    LogToConsole Now & " : Error " & Err.Number & ": " & Err.Description & _
                  " in clsLogger.LogInitialLine()"
    Resume Next

End Sub
