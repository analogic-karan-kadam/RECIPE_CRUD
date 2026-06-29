Public Sub SaveQueryLog(ByVal logText As String)

    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection

    Dim sqlQuery As String

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"

    sqlQuery = "INSERT INTO checkLogTime ([condition], [datetime]) " & _

               "VALUES ('" & Replace(logText, "'", "''") & "', GETDATE())"

    conn.Execute sqlQuery

CleanUp:

    If Not conn Is Nothing Then

        If conn.State = 1 Then conn.Close

        Set conn = Nothing

    End If

    Exit Sub

ErrHandler:

    MsgBox "Error saving log: " & Err.Description, vbExclamation

    Resume CleanUp

End Sub
'USE OF 
SaveQueryLog("BEFORE QUERY START")