Private Function CheckDuplicate(ByVal RECIPE_NAME As String) As Boolean

    Dim conn As ADODB.Connection

    Dim rs As ADODB.Recordset

    Dim sql As String


    Set conn = New ADODB.Connection

    Set rs = New ADODB.Recordset

    ' Open database connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"

       

    ' Count how many rows match the recipe name

    sql = "SELECT COUNT(*) FROM RECIPE WHERE RECIPE_NAME = '" & Replace(RECIPE_NAME, "'", "''") & "'"

   

    Set rs = conn.Execute(sql)

   

    ' If the count is greater than 0, a duplicate exists

    If rs.Fields(0).Value > 0 Then

        CheckDuplicate = True

    Else

        CheckDuplicate = False

    End If

   

    ' Clean up and close connections

    rs.Close

    conn.Close

    Set rs = Nothing

    Set conn = Nothing

End Function

'------------USE OF CHECK DUPLICATE 
    If CheckDuplicate(recipeTagGroup.Item("recipe\RECIPE_NAME").Value) Then

        MsgBox "Recipe Already Exists Choose Different Name for the Recipe"

        RECIPE_NAMEInput.Value = ""

        Exit Sub

    End If