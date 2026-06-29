Function renderRecipeParameters(ByVal RECIPE_NAME As String)
    
	'LogDiagnosticsMessage ("Render Function Called For Required Tag Only ")
	'MsgBox ("Render Function Called For Required Tag Only ")


    Dim conn

    Dim Rs

    Dim Field, xyz

    Dim ServerName, sqlUsername, sqlPassword, Table, Database As String

    If IsNull(RECIPE_NAME) Or Trim(RECIPE_NAME) = "" Or RECIPE_NAME = "RECIPE NAME" Or RECIPE_NAME = "RECIPE_NAME" Then

        MsgBox "Please select a recipe", vbExclamation, "Missing Recipe"

    Exit Function

    End If

   

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

    Set Rs = conn.Execute(StrQuery)


    ' ==============================================================

    ' SET TAG VALUES FROM RECORDSET

    ' ==============================================================

    Dim displayValue As Integer

    Dim i As Integer

   

    displayValue = recipeTagGroup.Item("recipe\Display").Value

    If Rs.EOF Then

        MsgBox "Recipe not found: " & RECIPE_NAME, vbExclamation

    Else

    Do While Not Rs.EOF


    Select Case displayValue


    '==========================

    ' lODING TAGS FOR Display = 0 

    '==========================

    Case 0

   

    For i = 1 To 150

    recipeTagGroup.Item("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i).Value = _

        Rs("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i)

    Next i

   

    
    For i = 1 To 150

    recipeTagGroup.Item("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i).Value = _

        Rs("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i)

    Next i


    For i = 1 To 50

    recipeTagGroup.Item("recipeTagGroup.Add "recipe\STRING_PARA_" & i).Value = _

        Rs("recipeTagGroup.Add "recipe\STRING_PARA_" & i)

    Next i
   

    '==========================

    ' lODING TAGS FOR Display = 1 

    '==========================

    Case 1

   

    For i = 1 To 150

    recipeTagGroup.Item("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i).Value = _

        Rs("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i)

    Next i

   

    
    For i = 1 To 150

    recipeTagGroup.Item("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i).Value = _

        Rs("recipeTagGroup.Add "recipe\NUMBER_PARA_" & i)

    Next i


    For i = 1 To 50

    recipeTagGroup.Item("recipeTagGroup.Add "recipe\STRING_PARA_" & i).Value = _

        Rs("recipeTagGroup.Add "recipe\STRING_PARA_" & i)

    Next i

   

    End Select


    ' Always update control tags

    recipeTagGroup.Item("recipe\STATUS").Value = Rs("STATUS").Value

    recipeTagGroup.Item("recipe\DISPLAY").Value = Rs("DISPLAY").Value

 
    Rs.MoveNext

   
    Loop

    End If

    ' ==============================================================

    ' CLEANUP

    ' ==============================================================

    Rs.Close

    conn.Close

    Set Rs = Nothing

    Set conn = Nothing

End Function
