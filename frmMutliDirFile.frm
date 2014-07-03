Option Explicit
'// Public call for selected value of ComboBox1
Public Property Get Selection() As String: Selection = Me.ComboBox1.Value: End Property




'=========================================================================================================
'## CommandButton1_Click
'   Hides 'frmMultiDirFile' when a selection has been made in the combobox. Returns a Dialog box if not.
'=========================================================================================================
Private Sub CommandButton1_Click()

    If Me.ComboBox1.Value = "" Then
        '// If no choice has been made
        MsgBox "You did not make a selection."
    Else
        '// If a choice was made
        Me.Hide
    End If
    
End Sub




'=========================================================================================================
'## CommandButton2_Click
'   End the program if the user needs to in case something goes wrong with populating the combobox
'=========================================================================================================
Private Sub CommandButton2_Click()

    End
    
End Sub




'=========================================================================================================
'## UserForm_QueryClose
'   Keeps the user from killing the 'frmMultiDirFile' window in a non-programmatic way (via the "X" button
'   in the top-right of the window. Redirects all action to the 'CommandButton1_Click' subroutine.
'=========================================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    '// If the Windows "X" button is clicked...
    If CloseMode = vbFormControlMenu Then
        '// Stop Windows from closing the window
        Cancel = True
        '// Run 'CommandButton1_Click' subroutine
        CommandButton1_Click
    End If
    
End Sub
