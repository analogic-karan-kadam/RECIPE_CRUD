 Private Sub updateRecipetoDB(user As String, comment As String)

    On Error GoTo ErrHandler

    LogDiagnosticsMessage ("Update recipe called by " & user)


    Dim conn       As ADODB.Connection
    Dim sqlQuery   As String
    Dim i          As Integer


    ' ==============================================================

    ' 1. VALIDATION

    ' ==============================================================

    If Trim(Me.RECIPE_NAMEInput.Value) = "" Then

        MsgBox "Error: Please Enter Recipe Name", vbCritical, "Missing Data"

        Exit Sub

    End If


 

    ' ==============================================================

    ' 2. DATABASE CONNECTION

    ' ==============================================================

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"


 

    ' ==============================================================

    ' 3. BUILD QUERY PARTS (WITH CLng FIX)

    ' ==============================================================


 

    Dim setString As String: setString = ""
    For i = 1 To 50

        setString = setString & _

            "STRING_PARA_" & i & " = " & _

            recipeTagGroup.Item("recipe\STRING_PARA_" & i).Value

        If i < 50 Then setString = setString & ", "

    Next i



	Dim setNumber1to50 As String: setNumber1to50 = ""
	For i = 1 To 50
		setNumber1to50 = setNumber1to50 & _
			"NUMBER_PARA_" & i & " = " & _
			CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value)
		If i < 50 Then setNumber1to50 = setNumber1to50 & ", "
	Next i

	Dim setNumber51to100 As String: setNumber51to100 = ""
	For i = 51 To 100
		setNumber51to100 = setNumber51to100 & _
			"NUMBER_PARA_" & i & " = " & _
			CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value)
		If i < 100 Then setNumber51to100 = setNumber51to100 & ", "
	Next i

	Dim setNumber101to150 As String: setNumber101to150 = ""
	For i = 101 To 150
		setNumber101to150 = setNumber101to150 & _
			"NUMBER_PARA_" & i & " = " & _
			CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value)
		If i < 150 Then setNumber101to150 = setNumber101to150 & ", "
	Next i

	Dim setNumber151to200 As String: setNumber151to200 = ""
	For i = 151 To 200
		setNumber151to200 = setNumber151to200 & _
			"NUMBER_PARA_" & i & " = " & _
			CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value)
		If i < 200 Then setNumber151to200 = setNumber151to200 & ", "
	Next i

	Dim setNumber201to250 As String: setNumber201to250 = ""
	For i = 201 To 250
		setNumber201to250 = setNumber201to250 & _
			"NUMBER_PARA_" & i & " = " & _
			CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value)
		If i < 250 Then setNumber201to250 = setNumber201to250 & ", "
	Next i

	Dim setNumber251to300 As String: setNumber251to300 = ""
	For i = 251 To 300
		setNumber251to300 = setNumber251to300 & _
			"NUMBER_PARA_" & i & " = " & _
			CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value)
		If i < 300 Then setNumber251to300 = setNumber251to300 & ", "
	Next i


    ' ==============================================================

    ' 4. FINAL QUERY

    ' ==============================================================

    sqlQuery = "UPDATE RECIPE SET " & _

               "updatedBy = '" & user & "', " & _

               "UPDATED_DATE = GETDATE(), " & _

               "STATUS = 2, " & _

               "version = version + 1, " & _

               "[Comment] = '" & Replace(comment, "'", "''") & "', " & _

               setString & ", " & _

				setNumber1to50 & ", " & _
				
				setNumber51to100 & ", " & _
				
				setNumber101to150 & ", " & _
				
				setNumber151to200 & ", " & _
				
				setNumber201to250 & ", " & _
				
				setNumber251to300 & _

               " WHERE RECIPE_NAME = '" & recipeTagGroup.Item("recipe\RECIPE_NAME").Value & "'"


 

    LogDiagnosticsMessage sqlQuery


 

    conn.Execute sqlQuery


 

    MsgBox "Recipe updated successfully!", vbInformation, "Success"

    LogDiagnosticsMessage "Recipe updated successfully"


 

    conn.Close

    Set conn = Nothing

    Exit Sub


 

ErrHandler:

    LogDiagnosticsMessage "Update Recipe Error: " & Err.Description

    MsgBox "Update Recipe Error: " & Err.Description, vbCritical, "Database Error"

    If Not conn Is Nothing Then

        If conn.State = 1 Then conn.Close

        Set conn = Nothing

    End If


 

End Sub