'=========================================================================================================
'VBA Object Name:   clsDataFile
'Filename:          PDELGenerator_Sensor_v##.xlsm
'Date:              June 11, 2014
'Version:           0.6
'Author:            Jeremiah Valenzuela
'
'Description:       This class takes PARTITION and UNIT NUMBER as inputs and outputs a copy of the
'                   worksheet pertaining to the aforementioned inputs. The UNIT NUMBER is used to
'                   determine the correct directory and file path. Where there is ambiguity, the user is
'                   given a selection of either multiple directories or files that could possibly have the
'                   desired data the user wants converted into a PDEL. The filepath is used to open the
'                   workbook and make a copy of it in a temporary folder. The PARTITION is used to find
'                   the desired worksheet containing the desired data. This worksheet is the output of
'                   this class (by way of all the soubroutines and functions).
'
'Code Outline:      Declarations, Properties, Subroutines, Functions
'
'Misc:              UDF (User Defined Function), UDS (User Defined Subroutine)
'=========================================================================================================

'## Force Explicit Variable Declaration
Option Explicit

'## Declaration of Class Variables/Properties
Private pPDELPath As String             'set in Class_Initialize()
Private pSourcePath As String           'set in Class_Initialize()
Private pLogFilePath As String          'set in Class_Initialize()
Private pTempPath As String             'set in Class_Initialize()
Public objLogger As clsDataFile         'instance of clsLogger for logging purposes

Private pPartition As String            'set in caller
Private pUnitNumber As String           'set in caller
Private pFileDirectory As String        'set in GetDirectory sub
Private pFilename As String             'set in GetFilename sub
Private pFilePath As String             'set in GetFilename sub
Private pWB As New Workbook             'set in LoadWorkbook sub
Private pWS As New Worksheet            'set in LoadWorksheet sub
Private objNewApp As New Application    'set in LoadWorkbook sub

Private pNoDir_Error As Boolean         'set in GetDirectory to Flag "No Directory"
Private pNoFile_Error As Boolean        'set in GetFilename to Flag "No File"
Private pNoPartition_Error As Boolean   'set in LoadFile to Flag "No Partition"
Private pNoUnitNumber_Error As Boolean  'set in LoadFile to Flag "No UnitNumber"
Private pNoSheet_Error As Boolean       'set in LoadWorksheet to Flag "No Sheet"

'## Class Property Getters/Setters
'// Public call for value of 'pWB'
Public Property Get WBook() As Workbook: Set WBook = pWB: End Property
'// Public call for value of 'pWS'
Public Property Get WSheet() As Worksheet: Set WSheet = pWS: End Property
'// Public call for value of 'pUnitNumber'
Public Property Get UnitNumber() As String: UnitNumber = pUnitNumber: End Property
'// Public call for value of 'pPartition'
Public Property Get Partition() As String: Partition = pPartition: End Property
'// Public call for value of 'pFileDirectory'
Public Property Get FileDirectory() As String: FileDirectory = pFileDirectory: End Property
'// Public call for value of 'pFilename'
Public Property Get Filename() As String: Filename = pFilename: End Property
'// Public call for value of 'pFilePath'
Public Property Get FilePath() As String: FilePath = pFilePath: End Property
'// Public call for value of 'pNoDir_Error'
Public Property Get NoDir_Error() As Boolean: NoDir_Error = pNoDir_Error: End Property
'// Public call for value of 'pNoFile_Error'
Public Property Get NoFile_Error() As Boolean: NoFile_Error = pNoFile_Error: End Property
'// Public call for value of 'pNoPartition_Error'
Public Property Get NoPartition_Error() As Boolean: NoPartition_Error = pNoPartition_Error: End Property
'// Public call for value of 'pNoUnitNumber_Error'
Public Property Get NoUnitNumber_Error() As Boolean: NoUnitNumber_Error = pNoUnitNumber_Error: End Property
'// Public call for value of 'pNoSheet_Error'
Public Property Get NoSheet_Error() As Boolean: NoSheet_Error = pNoSheet_Error: End Property




'=========================================================================================================
'## Class_Initialize
'   When this class is instantiated, this subroutine acts as a constructor for the new object. Initial
'   values fo the class variables and instantiation of objects are to be made here.
'=========================================================================================================
Private Sub Class_Initialize()

    On Error GoTo ErrorHandler
    
    '// Instantiate workbook, worksheet
    Set pWB = Nothing
    Set pWS = Nothing
    Set objLogger = New clsLogger
    
    '// Set path values from Configuration sheet of PDELGenerator_Sensor_v##.xlsm
    pPDELPath = ThisWorkbook.Sheets("Configuration").Range("B1").Value
    pSourcePath = ThisWorkbook.Sheets("Configuration").Range("B2").Value
    pLogFilePath = ThisWorkbook.Sheets("Configuration").Range("B3").Value
    pTempPath = ThisWorkbook.Sheets("Configuration").Range("B4").Value
    
    '// Set initial values for variables an flags
    pNoPartition_Error = False
    pNoUnitNumber_Error = False
    pNoDir_Error = False
    pNoFile_Error = False
    pNoSheet_Error = False
    
    '// Write activity to logfile and console
    objLogger.LogInitialLine
    objLogger.LogEverywhere "clsDataFile class initialized"
    
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description & " in clsDataFile.Class_Initialize()"
    objLogger.LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                            " in clsDataFile.Class_Initialize()"
    Resume Next

End Sub




'=========================================================================================================
'## LoadFile UDS
'   Takes as input PARTITION (myPartition) and UNIT NUMBER (myUnitNumber), error checks those inputs,
'   calls 'GetDirectory' to find the desired directory. If one is not found, the subroutine ends. Then
'   'GetFilename' tries to find the desired datasheet. If one is not found, the subroutine ends. Finally,
'   'LoadWorkbook' and 'LoadWorksheet' are called to acquire the desired data and sets that data as a
'   WORKSHEET to 'pWS'
'=========================================================================================================
Public Sub LoadFile(ByVal myPartition As String, ByVal myUnitNumber As String)

    '// Error check 'myPartition' input for valid string FOV, OE, TFOE. Set Flag to TRUE on invalid
    If myPartition <> "FOV" Then
        If myPartition <> "OE" Then
            If myPartition <> "TFOE" Then
                MsgBox "ERROR: Partition: " & myPartition & " is not valid. MUST be FOV, OE, TFOE." & _
                       " Please have the code maintainer look into it."
                objLogger.LogEverywhere "ERROR: Partition: " & myPartition & " is not valid. MUST be " & _
                                        "FOV, OE, TFOE. Please have the code maintainer look into it."
                pNoPartition_Error = True
            End If
        End If
    End If
    
    '// If 'myPartition' is not a valid input, the skip subsequent code
    If pNoPartition_Error = False Then
        '// Error check 'myUnitNumber' input for valid numerical string. Set Flag to TRUE on invalid.
        If IsNumeric(myUnitNumber) Then
            '// Error check 'myUnitNumber' for correct length of 4 numbers
            If Len(myUnitNumber) <> 4 Then
                MsgBox "ERROR: Unit number: " & myUnitNumber & " is not valid. It MUST be 4 numbers long."
                objLogger.LogEverywhere "ERROR: Unit number: " & myUnitNumber & " is not valid. It " & _
                                        "MUST be 4 numbers long."
                pNoUnitNumber_Error = True
            End If
        Else
            MsgBox "ERROR: Unit number " & myUnitNumber & " contains NON-NUMERIC characters. It MUST " & _
                   "be 4 numbers only."
            objLogger.LogEverywhere "ERROR: Unit number " & myUnitNumber & " contains NON-NUMERIC " & _
                                    "characters. It MUST be 4 numbers only."
            pNoUnitNumber_Error = True
        End If
        
        '// If 'myUnitNumber' is not valid input, then skip subsequent code
        If pNoUnitNumber_Error = False Then
            '// Assign valid inputs to class properties
            pPartition = myPartition
            pUnitNumber = myUnitNumber
            objLogger.LogEverywhere "SUCCESS! INPUTS are VALID! Partition: " & pPartition & _
                                    ". Unit Number: " & pUnitNumber & "."
            
            '// Call 'GetDirectory'
            GetDirectory
            
            '// If 'GetDirectory' finds no valid directories, Flag is set to TRUE and this subroutine
            '   finishes. If Flag is FALSE then continue on to 'GetFilename'. If no valid filename is
            '   found, Flag is set to TRUE and subroutine finishes. If Flag is still FALSE, then proceed
            '   to pulling workbook and worksheet.
            If pNoDir_Error = False Then
                GetFilename
                If pNoFile_Error = False Then
                    LoadWorkbook
                    LoadWorksheet
                    If pNoSheet_Error = True Then
                        MsgBox "A sheet was not found for the PARTITION: " & pPartition
                        objLogger.LogEverywhere "A sheet was not found for the PARTITION: " & pPartition
                    End If
                Else
                    MsgBox "A file was not found for the UNIT NUMBER: " & pUnitNumber & ". A valid " & _
                           "file has a filename with the UNIT NUMBER entered, does not have ""Copy"" " & _
                           "in it, must be an "".xls"" file, and contain ""Rev_1"". It must be in the" & _
                           " path: " & pSourcePath & pFileDirectory & "\"
                    objLogger.LogEverything "A file was not found for the UNIT NUMBER: " & pUnitNumber & _
                                            ". A valid file has a filename with the UNIT NUMBER " & _
                                            "entered, does not have ""Copy"" in it, must be an" & _
                                            " "".xls"" file, and contain ""Rev_1"". It must be in the" & _
                                            " path: " & pSourcePath & pFileDirectory & "\"
                End If
            Else
                MsgBox "A directory was not faound for the UNIT NUMBER: " & pUnitNumber & ". A valid " & _
                       "directory must contain the UNIT NUMBER entered and be in the path: " & pSourcePath
                objLogger.LogEverywhere "A directory was not found for the UNIT NUMBER: " & _
                                        pUnitNumber & ". A valid directory must contain the UNIT " & _
                                        "NUMBER entered and be in the path: " & pSourcePath
            End If
        Else
        End If
    Else
    End If
        
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description & " in clsDataFile.LoadFile()"
    objLogger.LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                            " in clsDataFile.LoadFile()"
    Resume Next

End Sub




'=========================================================================================================
'## GetDirectory UDS
'   Searches for the directory which contains the UNIT NUMBER in the name. Offers choices if there are
'   multiple options. Result is the desired path 'pFileDirectory' to search for the datasheet.
'=========================================================================================================
Private Sub GetDirectory()

    On Error GoTo ErrorHandler
    
    '// Declare variables
    Dim strDataDir() As String          'Valid directory array declared with undefined dimensions
    Dim strDir As String                'Container for the next item to asses for including in strDataDir
    
    ReDim strDataDir(0 To 0) As String  'Redimension for initial value
    Dim j As Integer                    'Index
    
    '// Scan 'pSourcePath' including directories and set the initial value for 'strDataDir'
    strDir = Dir(pSourcePath, vbDirectory)
    strDataDir(0) = strDir
    
    '// Search for directories that contain the UNIT NUMBER
    Do While strDir <> ""
        '// Get next item in the root directory
        strDir = Dir
        '// If current value of 'strDir' is a null string then exit loop
        If strDir = "" Then Exit Do
        '// If current value of 'strDataDir' contains UNIT NUMBER and next directory does also: append
        If InStrRev(strDataDir(UBound(strDataDir)), pUnitNumber) <> 0 And _
           InStrRev(strDir, pUnitNumber) <> 0 Then
            ReDim Preserve strDataDir(0 To UBound(strDataDir) + 1) As String
            strDataDir(UBound(strDataDir)) = strDir
        '// If current value of 'strDataDir' doesn't have UNIT NUMBER but next one does: overwrite
        ElseIf InStrRev(strDataDir(UBound(strDataDir)), pUnitNumber) = 0 And _
               InStrRev(strDir, pUnitNumber) <> 0 Then
            strDataDir(UBound(strDataDir)) = strDir
        End If
    Loop
    
    '// If 'strDataDir' has more than one valid directory with UNIT NUMBER, give user a choice
    If UBound(strDataDir) > 0 Then
        '// Clear any data in the combobox
        frmMultiDirFile.ComboBox1.Clear
        '// Populate combobox with multiple choices
        For j = 0 To UBound(strDataDir)
            frmMultiDirFile.ComboBox1.AddItem strDataDir(j)
        Next
        '// Display combobox, redimension 'strDataDir' for the single choice, and pull value from combobox
        frmMultiDirFile.Caption = "Multiple Directories Found: Please Choose One"
        frmMultiDirFile.Label1.Caption = "There are MULTIPLE DIRECTORIES from which to choose. " & _
                                         "Please use the pulldown to select a directory."
        frmMultiDirFile.Show
        ReDim strDataDir(0 To 0) As String
        strDataDir(0) = frmMultiDirFile.Selection
    End If
    
    '// Check to see if chosen directory is valid. If not, set Flag to TRUE
    If InStrRev(strDataDir(0), pUnitNumber) = 0 Then
        pNoDir_Error = True
    Else
        '// Set class property 'pFileDirectory' to the chosen or only found directory
        pFileDirectory = strDataDir(0)
        objLogger.LogEverywhere "SUCCESS! DIRECTORY found: \" & pFileDirectory & "\. Full Path: " & _
                                pSourcePath & pFileDirectory & "\."
    End If
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description & " in clsDataFile.GetDirectory()"
    objLogger.LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                            " in clsDataFile.GetDirectory()"
    Resume Next

End Sub




'=========================================================================================================
'## GetFilename UDS
'   Searches for the file with the UNIT NUMBER in the title, '.xls' extension, 'Rev_1' in the title, and
'   does not have 'Copy' in the title. If there is more than one option, a choice is given to the user.
'   The result is the filename saved to the 'pFilename' property and the full filepath saved to the
'   'pFilePath' property.
'=========================================================================================================
Private Sub GetFilename()

    On Error GoTo ErrorHandler
    
    '// Declare variables
    Dim strDataDirectoryPath As String  'Concatenation of 'pSourcePath' & 'pFileDirectory'
    Dim strDataFile() As String         'Valid file array declared with undefined dimensions
    Dim strFiles As String              'Container for the next item to assess for entry in 'strDataFile'
    ReDim strDataFile(0 To 0) As String 'Redimension for initial value
    Dim j As Integer                    'Index
    
    '// Build absolute path in which to scan for the datafile. Then start scanning.
    strDataDirectoryPath = pSourcePath & pFileDirectory & "\"
    strFiles = Dir(strDataDirectoryPath)
    strDataFile(0) = strFiles
    
    '// Search for files that contain the UNIT NUMBER
    Do While strFiles <> ""
        '// Get next item in the specified directory
        strFiles = Dir
        '// if current value of 'strFiles' is a null string then exit loop
        If strFiles = "" Then Exit Do
        '// If current value of 'strDataFile' is valid and current 'strFiles' is valid: append
        If InStrRev(UCase(strDataFile(UBound(strDataFile))), ".XLS") <> 0 And _
           InStrRev(UCase(strDataFile(UBound(strDataFile))), "COPY") = 0 And _
           InStrRev(UCase(strDataFile(UBound(strDataFile))), pUnitNumber) <> 0 And _
           InStrRev(UCase(strDataFile(UBound(strDataFile))), "REV_A") <> 0 And _
           InStrRev(UCase(strFiles), ".XLS") <> 0 And _
           InStrRev(UCase(strFiles), "COPY") = 0 And _
           InStrRev(UCase(strFiles), pUnitNumber) <> 0 And _
           InStrRev(UCase(strFiles), "REV_A") <> 0 Then
            ReDim Preserve strDataFile(0 To UBound(strDataFile) + 1) As String
            strDataFile(UBound(strDataFile)) = strFiles
        '// If current value of 'strDataFile' is NOT valid and current 'strFiles' is valid: overwrite
        ElseIf (InStrRev(UCase(strDataFile(UBound(strDataFile))), ".XLS") = 0 Or _
                InStrRev(UCase(strDataFile(UBound(strDataFile))), "COPY") <> 0 Or _
                InStrRev(UCase(strDataFile(UBound(strDataFile))), pUnitNumber) = 0 Or _
                InStrRev(UCase(strDataFile(UBound(strDataFile))), "REV_A") = 0) And _
               (InStrRev(UCase(strFiles), ".XLS") <> 0 And _
                InStrRev(UCase(strFiles), "COPY") = 0 And _
                InStrRev(UCase(strFiles), pUnitNumber) <> 0 And _
                InStrRev(UCase(strFiles), "REV_A") <> 0) Then
            strDataFile(UBound(strDataFile)) = strFiles
        End If
    Loop
    
    '// If 'strDataFile' has more than one valid datafile, give user a choice between them
    If UBound(strDataFile) > 0 Then
        '// Clear any data in the combobox
        frmMultiDirFile.ComboBox1.Clear
        '// Populate combobox with multiple choices
        For j = 0 To UBound(strDataFile)
            frmMultiDirFile.ComboBox1.AddItem strDataFile(j)
        Next
        '// Display combobox, redimension 'strDataFile' for single choice, and pull value from combobox
        frmMultiDirFile.Caption = "Multiple Datafiles Found: Please Choose One"
        frmMultiDirFile.Label1.Caption = "There are MULTIPLE DATAFILES from which to choose. " & _
                                         "Please use the pulldown to select a file."
        frmMultiDirFile.Show
        ReDim strDataFile(0 To 0) As String
        strDataFile(0) = frmMultiDirFile.Selection
    End If
    
    '// Check to see if chosen file is valid. If not, set Flag to TRUE
    If InStrRev(strDataFile(0), pUnitNumber) = 0 Or InStrRev(UCase(strDataFile(0)), ".XLS") = 0 Or _
       InStrRev(UCase(strDataFile(0)), "COPY") <> 0 Or InStrRev(UCase(strDataFile(0)), "REV_A") = 0 Then
        pNoFile_Error = True
    Else
        '// Set class properties 'pFilename' & 'pFilePath' to chosen datafile & full filepath
        pFilename = strDataFile(0)
        pFilePath = pSourcePath & pFileDirectory & "\" & pFilename
        objLogger.LogEverywhere "SUCCESS! DATAFILE found: " & pFilename & ". Full FilePath: " & _
                                pFilePath & "."
    End If
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description & " in clsDataFile.GetFilename()"
    objLogger.LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                            " in clsDataFile.GetFilename()"
    Resume Next

End Sub




'=========================================================================================================
'## LoadWorkbook UDS
'   Takes the 'pFilePath' property value and opens the appropriate workbook. The worksheet is copied into
'   a temp folder and set to the 'pWB'
'=========================================================================================================
Private Sub LoadWorkbook()

    On Error GoTo ErrorHandler
    
    '// Declare variable
    Dim wFS As Object   'Expendable native object to employ 'CopyFile' method
    
    '// Makes a temporary copy of the found file into a temp folder
    Set wFS = CreateObject("Scripting.FileSystemObject")
    wFS.CopyFile pFilePath, pTempPath & "workingdatasheet.xls"
    Set wFS = Nothing
    
    '// Makes new instance of Excel open in a hidden state and opens temp workbook into 'pWB'
    objNewApp.Visible = False
    Set pWB = objNewApp.Workbooks.Add(pTempPath & "workingdatasheet.xls")
    
    '// Log success of copying the workbook and storing it in 'pWB'
    If pWB Is Nothing Then
    Else
        objLogger.LogEverywhere "SUCCESS! WORKBOOK copied to Temp folder and stored in 'pWB' object."
    End If
    
    Exit Sub

    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description & " in clsDataFile.LoadWorkbook()"
    objLogger.LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                            " in clsDataFile.LoadWorkbook()"
    Resume Next

End Sub




'=========================================================================================================
'## LoadWorksheet UDS
'   Takes the 'pFilePath' property value and opens the appropriate workbook. The worksheet is copied into
'   a temp folder and set to the 'pWB'
'=========================================================================================================
Private Sub LoadWorkbook()

    On Error GoTo ErrorHandler
    
    '// Declare variable
    Dim intSheets() As Integer          'Valid integer array set without dimensions
    ReDim intSheets(0 To 0) As Integer  'Redimension array for initial value
    Dim k As Integer                    'Index for looking through sheets
    Dim j As Integer                    'Index for identifying multiple valid sheets
    Dim tempIndex As Integer            'Index of finally chosen sheet name
    Dim strSheetName As String          'SheetName to search for from pPartition
    
    '// Set search string for sheet name based on 'pPartition' value
    Select Case pPartition
        Case "FOV"
            strSheetName = "IFOV_SFOV"
        Case "OE"
            strSheetName = "FINAL_OE"
        Case "TFOE"
            strSheetName = "FINAL_OE"
    End Select
    
    '// Iterate through sheets to find one(s) with names matching 'strSheetName'
    For k = 1 To pWB.Sheets.Count
        '// If not a match, go to next sheet
        If InStrRev(pWB.Sheets(k).Name, strSheetName) <> 0 Then
            '// If initial value of 'intSheets' is 0, populate the first array position with sheet index
            If intSheets(UBound(intSheets)) = 0 Then
                intSheets(UBound(intSheets)) = k
            '// if initial value of 'intSheets' is valid, then add to array and populate
            Else
                ReDim Preserve intSheets(0 To UBound(intSheets) + 1) As Integer
                intSheets(UBound(intSheets)) = k
            End If
        End If
    Next
    
    '// If multiple valid sheets are found, display choices between which users may choose
    If UBound(intSheets) > 0 Then
        '// Clear any data in the combobox
        frmMultiDirFile.ComboBox1.Clear
        '// Populate combobox with multiple choices
        For j = 0 To UBound(intSheets)
            frmMultiDirFile.ComboBox1.AddItem Trim(Str(intSheets(j))) & " " & _
                                              pWB.Sheets(intSheets(j)).Name
        Next
        '// Display combobox, redimension 'intSheets' for single choice, and pull value from combobox
        frmMultiDirFile.Caption = "Multiple Sheets Found: Please Choose One"
        frmMultiDirFile.Label1.Caption = "There are MULTIPLE SHEETS from which to choose. " & _
                                         "Please use the pulldown to select a sheet."
        frmMultiDirFile.Show
        ReDim intSheets(0 To 0) As Integer
        '// Split Index and Title combo string into index only and convert to integer
        tempIndex = InStr(frmMultiDirFile.Selection, " ")
        intSheets(0) = CInt(Left(frmMultiDirFile.Selection, tempIndex - 1))
    End If
    
    '// Check to see if chosen sheet is valid. If not, set Flag to TRUE
    If intSheets(0) = 0 Then
        pNoSheet_Error = True
    Else
        '// Set class property 'pWS' to chosen sheet
        Set pWS = pWB.Sheets(intSheets(0))
        objLogger.LogEverywhere "SUCCESS! WORKSHEET found for Partition: " & pPartition & " in sheet " & _
                                intSheets(0) & " """ & pWS.Name & """ " & ". Worksheet is stored in " & _
                                "'pWS' object."
    End If
        
    Exit Sub

    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description & " in clsDataFile.LoadWorksheet()"
    objLogger.LogEverywhere Now & " : Error " & Err.Number & ": " & Err.Description & _
                            " in clsDataFile.LoadWorksheet()"
    Resume Next

End Sub




'=========================================================================================================
'## Class_Terminate
'   When this class is terminated, this subroutine acts as a destructor for the object made from this
'   class. All objects internal to this class are closed and set to Nothing. This cleans up the memory and
'   gracefully closes the instance of this class.
'=========================================================================================================
Private Sub Class_Terminate()

    On Error GoTo ErrorHandler
    
    '// Close the opened workbook and worksheet
    pWB.Close SaveChanges:=False
    Set pWB = Nothing
    Set pWS = Nothing
    
    '// Close the hidden instance of Excel
    objNewApp.Quit
    Set objNewApp = Nothing
    
    '// Close the internal instance of 'clsLogger'
    Set objLogger = Nothing
    
    '// Delete the workbook file copy in the Temp directory
    Kill pTempPath & "workingdatasheet.xls"
    
    Exit Sub
    
    '// Error Handler will display error to user and log it
ErrorHandler:
    MsgBox "Error " & Err.Number & " : " & Err.Description & " in clsDataFile.Class_Terminate()"
    Resume Next
    
End Sub
