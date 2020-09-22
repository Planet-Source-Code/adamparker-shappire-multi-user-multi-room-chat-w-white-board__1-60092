Attribute VB_Name = "modOfflineMessages"
Public Sub SaveText(txtContents, FilesPath)
    On Error Resume Next
    
    Open FilesPath For Output As #1
        Print #1, txtContents
    Close #1
End Sub

Public Function LoadText(FilesPath) As String
    On Error Resume Next
    Dim txtContents As String
    
    Open FilesPath For Input As #1
        txtContents = Input(LOF(1), #1)
    Close #1
    
    LoadText = txtContents
End Function

Public Sub DeleteFile(FilesPath)
    If FileExists(FilesPath) = True Then
        Kill FilesPath
    End If
End Sub

Public Function FileExists(FilesPath) As Boolean
    If Len(FilesPath) = 0 Then
        FileExists = False
        Exit Function
    End If
    
    If Len(Dir$(FilesPath)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
