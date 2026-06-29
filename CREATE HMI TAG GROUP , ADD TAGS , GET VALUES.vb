Dim recipeTagGroup As TagGroup
If recipeTagGroup Is Nothing Then
    Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)
        '---------ADDING TAGS
        recipeTagGroup.Add "recipe\RECIPE_NAME"
        recipeTagGroup.Add "recipe\SCADA_DATA_Batch_No"
        '--------- GETTING VALUES FORM TAGS
		recipeTagGroup.Item("recipe\STRING_PARA_1" ).Value = ""
End If