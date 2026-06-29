Private Sub updateRecipeStatus(RECIPE_STATUS As Integer, RECIPE_NAME As String, COMMENT As String, PerformerUser As String)

    Dim conn As ADODB.Connection
    Dim sql As String

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"

    sql = "UPDATE RECIPE " & _

          "SET STATUS = " & RECIPE_STATUS & ", " & _

          "[COMMENT] = '" & Replace(COMMENT, "'", "''") & "', " & _

          "APPROVED_BY = '" & Replace(PerformerUser, "'", "''") & "', " & _

          "APPROVED_DATE = GETDATE() " & _

          "WHERE RECIPE_NAME = '" & Replace(RECIPE_NAME, "'", "''") & "'"


 

    conn.Execute sql

    If RECIPE_STATUS Then

     MsgBox "Recipe Approved SuccessFully"

    Else

     MsgBox "Recipe Released For Modification"

    End If

    conn.Close
    Set conn = Nothing

End Sub
'------HOW TO USE 
 Call updateRecipeStatus(1, recipeTagGroup.Item("recipe\RECIPE_NAME").Value, PerformerComment, PerformerUser)