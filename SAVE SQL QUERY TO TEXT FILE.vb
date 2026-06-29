Public Sub SaveSqlQueryToFile(ByVal filePath As String, ByVal textData As String)

    On Error GoTo ErrHandler

    Dim fileNum As Integer

    fileNum = FreeFile

    Open filePath For Output As #fileNum

        Print #fileNum, textData

    Close #fileNum

    Exit Sub

ErrHandler:

    MsgBox "Error saving file: " & Err.Description

End Sub