Option Explicit

Public Sub ExportAllModules()

    Dim nowDate As String 'today's date and time
    Dim wb As Workbook 'current workbook object
    Dim i As Long 'vbComponent object index
    Dim currentFilename As String 'current excel file's filename
    Dim currentMaker As String 'user who made current excel file
    Dim currentVer As String 'version of current excel file
    Dim nStartIndex As Long 'index
    Dim nEndIndex As Long 'index
    Dim outPath As String 'directory of export
    
    Set wb = ThisWorkbook 'instantiate current workbook
    
    currentFilename = ThisWorkbook.FullName 'pull path/filename of current excel file
    
    'pull out the user who made the current excel file
    nStartIndex = InStr(1, currentFilename, "\")
    nEndIndex = InStr(nStartIndex + 1, currentFilename, "\") - 1
    currentMaker = Mid$(currentFilename, nStartIndex, nEndIndex - nStartIndex)
    
    'pull out the version of the current excel file
    nStartIndex = InStr(1, currentFilename, "v")
    nEndIndex = InStr(nStartIndex + 1, currentFilename, ".")
    currentVer = Mid$(currentFilename, nStartIndex, nEndIndex - nStartIndex)
    
    'get current date/time and set output directory path and create the directory
    nowDate = Format(Now(), "YYYYMMDD_hhmmss")
    outPath = "C:\Documents and Settings\" & Environ("UserName") & "\Desktop\PDELCompare_" & _
              currentMaker & "_" & currentVer & "_" & nowDate & "\"
    MkDir outPath
    
    'iterate through workbook's components and export them based on type
    For i = 1 To wb.VBProject.VBComponents.Count
        With wb.VBProject.VBComponents(i)
            If .Type = 1 And .Name <> "Export" Then 'modules
                ThisWorkbook.VBProject.VBComponents(.Name).Export outPath & .Name & ".bas"
            ElseIf .Type = 2 Then 'class modules
                ThisWorkbook.VBProject.VBComponents(.Name).Export outPath & .Name & ".cls"
            ElseIf .Type = 3 Then 'forms
                ThisWorkbook.VBProject.VBComponents(.Name).Export outPath & .Name & ".frm"
            End If
        End With
    Next
        
    Set wb = Nothing 'close out workbook object

End Sub

Public Sub ManualRun()
    MainForm.Show
    
    MainForm.tbSerNum.BackColor = vbWhite
    MainForm.ComboBox1.BackColor = vbWhite
    MainForm.ComboBox2.BackColor = vbWhite
    MainForm.ComboBox3.BackColor = vbWhite
    
    MainForm.Label2.Caption = Sheet1.Range("B2").Value
End Sub

Public Sub ClearImmediateWindow()
    Application.SendKeys "^g ^a {DEL}"
End Sub
