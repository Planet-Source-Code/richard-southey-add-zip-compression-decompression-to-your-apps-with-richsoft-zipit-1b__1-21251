Attribute VB_Name = "ZipView"
Public Sub LogError(Source As String, ErrorObject As ErrObject, Display As Boolean)
    On Error Resume Next
    'This procedure logs errors
    If Display = True Then
        'Display an error message
        MsgBox ErrorObject.Number & ":" & ErrorObject.Description, vbExclamation, App.ProductName
    End If
    
    'Open the log file and log the error
    F = FreeFile
    Open App.Path & "\Zipit Error Log.txt" For Append As #F
        Print #F, Now, Source, ErrorObject.Number, ErrorObject.Description
    Close #F
End Sub





