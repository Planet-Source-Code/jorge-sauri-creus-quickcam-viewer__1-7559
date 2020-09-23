Attribute VB_Name = "globals"
Public strFilesave As String
Public Sub SaveAs()
    strFilesave = InputBox("File name: ", "Save As...", "")
    strFilesave = Trim(strFilesave)
    If Len(strFilesave) <> 0 Then
        SavePicture imgScreen.Picture, strFilesave
    End If
End Sub




