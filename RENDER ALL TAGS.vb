Function renderRecipeParameters(ByVal RECIPE_NAME As String)

    'MsgBox ("Render function for " & RECIPE_NAME)

    Dim conn

    Dim rs

    Dim Field, xyz

    Dim ServerName, sqlUsername, sqlPassword, Table, Database As String

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"


    ' ==============================================================

    ' FETCH RECIPE ROW

    ' ==============================================================

    StrQuery = "SELECT * from RECIPE " & _

               "WHERE RECIPE_NAME = '" & RECIPE_NAME & "'"


    Set rs = conn.Execute(StrQuery)



    ' ==============================================================

    ' SET TAG VALUES FROM RECORDSET

    ' ==============================================================

    If rs.EOF Then

        'MsgBox "Recipe not found: " & RECIPE_NAME, vbExclamation

    Else

        Do While Not rs.EOF

            ' --- LOADING STRING TAGS

            For i = 1 To 50

                recipeTagGroup.Item("recipe\STRING_PARA_" & i).Value = _

                    rs("STRING_PARA_" & i)

            Next i
			
            ' --- LOADING NUMBER TAGS

            For i = 1 To 300

                recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value = _

                    rs("NUMBER_PARA_" & i)

            Next i


            ' --- Control Tags ---

            recipeTagGroup.Item("recipe\STATUS").Value = rs("STATUS").Value

            recipeTagGroup.Item("recipe\DISABLE").Value = rs("DISABLE").Value


            rs.MoveNext

        Loop

    End If


    ' ==============================================================

    ' CLEANUP

    ' ==============================================================

    rs.Close

    conn.Close

    Set rs = Nothing

    Set conn = Nothing

End Function