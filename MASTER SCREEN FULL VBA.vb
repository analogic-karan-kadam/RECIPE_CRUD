Dim recipeTagGroup As TagGroup
Private Function InitRecipeTagGroup() As TagGroup
    If recipeTagGroup Is Nothing Then
        Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)
        ' --- Meta Data ---
RECIPE_NAME VARCHAR(50) NOT NULL UNIQUE,
CREATED_BY"         '---------string data type
UPDATED_BY"         '---------string data type
APPROVED_BY"         '---------string data type
REJECTED_BY"         '---------string data type
STATUS"         '---------analog data type
DISPLAY"         '---------digital data type
VERSION"         '---------analog data type
DISPLAY"         '---------analog data type
        ' --- String Data type ---
recipeTagGroup.Add "recipe\STRING_PARA_1"
recipeTagGroup.Add "recipe\STRING_PARA_2"
recipeTagGroup.Add "recipe\STRING_PARA_3"
recipeTagGroup.Add "recipe\STRING_PARA_4"
recipeTagGroup.Add "recipe\STRING_PARA_5"
recipeTagGroup.Add "recipe\STRING_PARA_6"
recipeTagGroup.Add "recipe\STRING_PARA_7"
recipeTagGroup.Add "recipe\STRING_PARA_8"
recipeTagGroup.Add "recipe\STRING_PARA_9"
recipeTagGroup.Add "recipe\STRING_PARA_10"
recipeTagGroup.Add "recipe\STRING_PARA_11"
recipeTagGroup.Add "recipe\STRING_PARA_12"
recipeTagGroup.Add "recipe\STRING_PARA_13"
recipeTagGroup.Add "recipe\STRING_PARA_14"
recipeTagGroup.Add "recipe\STRING_PARA_15"
recipeTagGroup.Add "recipe\STRING_PARA_16"
recipeTagGroup.Add "recipe\STRING_PARA_17"
recipeTagGroup.Add "recipe\STRING_PARA_18"
recipeTagGroup.Add "recipe\STRING_PARA_19"
recipeTagGroup.Add "recipe\STRING_PARA_20"
recipeTagGroup.Add "recipe\STRING_PARA_21"
recipeTagGroup.Add "recipe\STRING_PARA_22"
recipeTagGroup.Add "recipe\STRING_PARA_23"
recipeTagGroup.Add "recipe\STRING_PARA_24"
recipeTagGroup.Add "recipe\STRING_PARA_25"
recipeTagGroup.Add "recipe\STRING_PARA_26"
recipeTagGroup.Add "recipe\STRING_PARA_27"
recipeTagGroup.Add "recipe\STRING_PARA_28"
recipeTagGroup.Add "recipe\STRING_PARA_29"
recipeTagGroup.Add "recipe\STRING_PARA_30"
recipeTagGroup.Add "recipe\STRING_PARA_31"
recipeTagGroup.Add "recipe\STRING_PARA_32"
recipeTagGroup.Add "recipe\STRING_PARA_33"
recipeTagGroup.Add "recipe\STRING_PARA_34"
recipeTagGroup.Add "recipe\STRING_PARA_35"
recipeTagGroup.Add "recipe\STRING_PARA_36"
recipeTagGroup.Add "recipe\STRING_PARA_37"
recipeTagGroup.Add "recipe\STRING_PARA_38"
recipeTagGroup.Add "recipe\STRING_PARA_39"
recipeTagGroup.Add "recipe\STRING_PARA_40"
recipeTagGroup.Add "recipe\STRING_PARA_41"
recipeTagGroup.Add "recipe\STRING_PARA_42"
recipeTagGroup.Add "recipe\STRING_PARA_43"
recipeTagGroup.Add "recipe\STRING_PARA_44"
recipeTagGroup.Add "recipe\STRING_PARA_45"
recipeTagGroup.Add "recipe\STRING_PARA_46"
recipeTagGroup.Add "recipe\STRING_PARA_47"
recipeTagGroup.Add "recipe\STRING_PARA_48"
recipeTagGroup.Add "recipe\STRING_PARA_49"
recipeTagGroup.Add "recipe\STRING_PARA_50"
        ' --- Real ( Int , float , double , analog ) Data type ---
recipeTagGroup.Add "recipe\NUMBER_PARA_1"
recipeTagGroup.Add "recipe\NUMBER_PARA_2"
recipeTagGroup.Add "recipe\NUMBER_PARA_3"
recipeTagGroup.Add "recipe\NUMBER_PARA_4"
recipeTagGroup.Add "recipe\NUMBER_PARA_5"
recipeTagGroup.Add "recipe\NUMBER_PARA_6"
recipeTagGroup.Add "recipe\NUMBER_PARA_7"
recipeTagGroup.Add "recipe\NUMBER_PARA_8"
recipeTagGroup.Add "recipe\NUMBER_PARA_9"
recipeTagGroup.Add "recipe\NUMBER_PARA_10"
recipeTagGroup.Add "recipe\NUMBER_PARA_11"
recipeTagGroup.Add "recipe\NUMBER_PARA_12"
recipeTagGroup.Add "recipe\NUMBER_PARA_13"
recipeTagGroup.Add "recipe\NUMBER_PARA_14"
recipeTagGroup.Add "recipe\NUMBER_PARA_15"
recipeTagGroup.Add "recipe\NUMBER_PARA_16"
recipeTagGroup.Add "recipe\NUMBER_PARA_17"
recipeTagGroup.Add "recipe\NUMBER_PARA_18"
recipeTagGroup.Add "recipe\NUMBER_PARA_19"
recipeTagGroup.Add "recipe\NUMBER_PARA_20"
recipeTagGroup.Add "recipe\NUMBER_PARA_21"
recipeTagGroup.Add "recipe\NUMBER_PARA_22"
recipeTagGroup.Add "recipe\NUMBER_PARA_23"
recipeTagGroup.Add "recipe\NUMBER_PARA_24"
recipeTagGroup.Add "recipe\NUMBER_PARA_25"
recipeTagGroup.Add "recipe\NUMBER_PARA_26"
recipeTagGroup.Add "recipe\NUMBER_PARA_27"
recipeTagGroup.Add "recipe\NUMBER_PARA_28"
recipeTagGroup.Add "recipe\NUMBER_PARA_29"
recipeTagGroup.Add "recipe\NUMBER_PARA_30"
recipeTagGroup.Add "recipe\NUMBER_PARA_31"
recipeTagGroup.Add "recipe\NUMBER_PARA_32"
recipeTagGroup.Add "recipe\NUMBER_PARA_33"
recipeTagGroup.Add "recipe\NUMBER_PARA_34"
recipeTagGroup.Add "recipe\NUMBER_PARA_35"
recipeTagGroup.Add "recipe\NUMBER_PARA_36"
recipeTagGroup.Add "recipe\NUMBER_PARA_37"
recipeTagGroup.Add "recipe\NUMBER_PARA_38"
recipeTagGroup.Add "recipe\NUMBER_PARA_39"
recipeTagGroup.Add "recipe\NUMBER_PARA_40"
recipeTagGroup.Add "recipe\NUMBER_PARA_41"
recipeTagGroup.Add "recipe\NUMBER_PARA_42"
recipeTagGroup.Add "recipe\NUMBER_PARA_43"
recipeTagGroup.Add "recipe\NUMBER_PARA_44"
recipeTagGroup.Add "recipe\NUMBER_PARA_45"
recipeTagGroup.Add "recipe\NUMBER_PARA_46"
recipeTagGroup.Add "recipe\NUMBER_PARA_47"
recipeTagGroup.Add "recipe\NUMBER_PARA_48"
recipeTagGroup.Add "recipe\NUMBER_PARA_49"
recipeTagGroup.Add "recipe\NUMBER_PARA_50"
recipeTagGroup.Add "recipe\NUMBER_PARA_51"
recipeTagGroup.Add "recipe\NUMBER_PARA_52"
recipeTagGroup.Add "recipe\NUMBER_PARA_53"
recipeTagGroup.Add "recipe\NUMBER_PARA_54"
recipeTagGroup.Add "recipe\NUMBER_PARA_55"
recipeTagGroup.Add "recipe\NUMBER_PARA_56"
recipeTagGroup.Add "recipe\NUMBER_PARA_57"
recipeTagGroup.Add "recipe\NUMBER_PARA_58"
recipeTagGroup.Add "recipe\NUMBER_PARA_59"
recipeTagGroup.Add "recipe\NUMBER_PARA_60"
recipeTagGroup.Add "recipe\NUMBER_PARA_61"
recipeTagGroup.Add "recipe\NUMBER_PARA_62"
recipeTagGroup.Add "recipe\NUMBER_PARA_63"
recipeTagGroup.Add "recipe\NUMBER_PARA_64"
recipeTagGroup.Add "recipe\NUMBER_PARA_65"
recipeTagGroup.Add "recipe\NUMBER_PARA_66"
recipeTagGroup.Add "recipe\NUMBER_PARA_67"
recipeTagGroup.Add "recipe\NUMBER_PARA_68"
recipeTagGroup.Add "recipe\NUMBER_PARA_69"
recipeTagGroup.Add "recipe\NUMBER_PARA_70"
recipeTagGroup.Add "recipe\NUMBER_PARA_71"
recipeTagGroup.Add "recipe\NUMBER_PARA_72"
recipeTagGroup.Add "recipe\NUMBER_PARA_73"
recipeTagGroup.Add "recipe\NUMBER_PARA_74"
recipeTagGroup.Add "recipe\NUMBER_PARA_75"
recipeTagGroup.Add "recipe\NUMBER_PARA_76"
recipeTagGroup.Add "recipe\NUMBER_PARA_77"
recipeTagGroup.Add "recipe\NUMBER_PARA_78"
recipeTagGroup.Add "recipe\NUMBER_PARA_79"
recipeTagGroup.Add "recipe\NUMBER_PARA_80"
recipeTagGroup.Add "recipe\NUMBER_PARA_81"
recipeTagGroup.Add "recipe\NUMBER_PARA_82"
recipeTagGroup.Add "recipe\NUMBER_PARA_83"
recipeTagGroup.Add "recipe\NUMBER_PARA_84"
recipeTagGroup.Add "recipe\NUMBER_PARA_85"
recipeTagGroup.Add "recipe\NUMBER_PARA_86"
recipeTagGroup.Add "recipe\NUMBER_PARA_87"
recipeTagGroup.Add "recipe\NUMBER_PARA_88"
recipeTagGroup.Add "recipe\NUMBER_PARA_89"
recipeTagGroup.Add "recipe\NUMBER_PARA_90"
recipeTagGroup.Add "recipe\NUMBER_PARA_91"
recipeTagGroup.Add "recipe\NUMBER_PARA_92"
recipeTagGroup.Add "recipe\NUMBER_PARA_93"
recipeTagGroup.Add "recipe\NUMBER_PARA_94"
recipeTagGroup.Add "recipe\NUMBER_PARA_95"
recipeTagGroup.Add "recipe\NUMBER_PARA_96"
recipeTagGroup.Add "recipe\NUMBER_PARA_97"
recipeTagGroup.Add "recipe\NUMBER_PARA_98"
recipeTagGroup.Add "recipe\NUMBER_PARA_99"
recipeTagGroup.Add "recipe\NUMBER_PARA_100"
recipeTagGroup.Add "recipe\NUMBER_PARA_101"
recipeTagGroup.Add "recipe\NUMBER_PARA_102"
recipeTagGroup.Add "recipe\NUMBER_PARA_103"
recipeTagGroup.Add "recipe\NUMBER_PARA_104"
recipeTagGroup.Add "recipe\NUMBER_PARA_105"
recipeTagGroup.Add "recipe\NUMBER_PARA_106"
recipeTagGroup.Add "recipe\NUMBER_PARA_107"
recipeTagGroup.Add "recipe\NUMBER_PARA_108"
recipeTagGroup.Add "recipe\NUMBER_PARA_109"
recipeTagGroup.Add "recipe\NUMBER_PARA_110"
recipeTagGroup.Add "recipe\NUMBER_PARA_111"
recipeTagGroup.Add "recipe\NUMBER_PARA_112"
recipeTagGroup.Add "recipe\NUMBER_PARA_113"
recipeTagGroup.Add "recipe\NUMBER_PARA_114"
recipeTagGroup.Add "recipe\NUMBER_PARA_115"
recipeTagGroup.Add "recipe\NUMBER_PARA_116"
recipeTagGroup.Add "recipe\NUMBER_PARA_117"
recipeTagGroup.Add "recipe\NUMBER_PARA_118"
recipeTagGroup.Add "recipe\NUMBER_PARA_119"
recipeTagGroup.Add "recipe\NUMBER_PARA_120"
recipeTagGroup.Add "recipe\NUMBER_PARA_121"
recipeTagGroup.Add "recipe\NUMBER_PARA_122"
recipeTagGroup.Add "recipe\NUMBER_PARA_123"
recipeTagGroup.Add "recipe\NUMBER_PARA_124"
recipeTagGroup.Add "recipe\NUMBER_PARA_125"
recipeTagGroup.Add "recipe\NUMBER_PARA_126"
recipeTagGroup.Add "recipe\NUMBER_PARA_127"
recipeTagGroup.Add "recipe\NUMBER_PARA_128"
recipeTagGroup.Add "recipe\NUMBER_PARA_129"
recipeTagGroup.Add "recipe\NUMBER_PARA_130"
recipeTagGroup.Add "recipe\NUMBER_PARA_131"
recipeTagGroup.Add "recipe\NUMBER_PARA_132"
recipeTagGroup.Add "recipe\NUMBER_PARA_133"
recipeTagGroup.Add "recipe\NUMBER_PARA_134"
recipeTagGroup.Add "recipe\NUMBER_PARA_135"
recipeTagGroup.Add "recipe\NUMBER_PARA_136"
recipeTagGroup.Add "recipe\NUMBER_PARA_137"
recipeTagGroup.Add "recipe\NUMBER_PARA_138"
recipeTagGroup.Add "recipe\NUMBER_PARA_139"
recipeTagGroup.Add "recipe\NUMBER_PARA_140"
recipeTagGroup.Add "recipe\NUMBER_PARA_141"
recipeTagGroup.Add "recipe\NUMBER_PARA_142"
recipeTagGroup.Add "recipe\NUMBER_PARA_143"
recipeTagGroup.Add "recipe\NUMBER_PARA_144"
recipeTagGroup.Add "recipe\NUMBER_PARA_145"
recipeTagGroup.Add "recipe\NUMBER_PARA_146"
recipeTagGroup.Add "recipe\NUMBER_PARA_147"
recipeTagGroup.Add "recipe\NUMBER_PARA_148"
recipeTagGroup.Add "recipe\NUMBER_PARA_149"
recipeTagGroup.Add "recipe\NUMBER_PARA_150"
recipeTagGroup.Add "recipe\NUMBER_PARA_151"
recipeTagGroup.Add "recipe\NUMBER_PARA_152"
recipeTagGroup.Add "recipe\NUMBER_PARA_153"
recipeTagGroup.Add "recipe\NUMBER_PARA_154"
recipeTagGroup.Add "recipe\NUMBER_PARA_155"
recipeTagGroup.Add "recipe\NUMBER_PARA_156"
recipeTagGroup.Add "recipe\NUMBER_PARA_157"
recipeTagGroup.Add "recipe\NUMBER_PARA_158"
recipeTagGroup.Add "recipe\NUMBER_PARA_159"
recipeTagGroup.Add "recipe\NUMBER_PARA_160"
recipeTagGroup.Add "recipe\NUMBER_PARA_161"
recipeTagGroup.Add "recipe\NUMBER_PARA_162"
recipeTagGroup.Add "recipe\NUMBER_PARA_163"
recipeTagGroup.Add "recipe\NUMBER_PARA_164"
recipeTagGroup.Add "recipe\NUMBER_PARA_165"
recipeTagGroup.Add "recipe\NUMBER_PARA_166"
recipeTagGroup.Add "recipe\NUMBER_PARA_167"
recipeTagGroup.Add "recipe\NUMBER_PARA_168"
recipeTagGroup.Add "recipe\NUMBER_PARA_169"
recipeTagGroup.Add "recipe\NUMBER_PARA_170"
recipeTagGroup.Add "recipe\NUMBER_PARA_171"
recipeTagGroup.Add "recipe\NUMBER_PARA_172"
recipeTagGroup.Add "recipe\NUMBER_PARA_173"
recipeTagGroup.Add "recipe\NUMBER_PARA_174"
recipeTagGroup.Add "recipe\NUMBER_PARA_175"
recipeTagGroup.Add "recipe\NUMBER_PARA_176"
recipeTagGroup.Add "recipe\NUMBER_PARA_177"
recipeTagGroup.Add "recipe\NUMBER_PARA_178"
recipeTagGroup.Add "recipe\NUMBER_PARA_179"
recipeTagGroup.Add "recipe\NUMBER_PARA_180"
recipeTagGroup.Add "recipe\NUMBER_PARA_181"
recipeTagGroup.Add "recipe\NUMBER_PARA_182"
recipeTagGroup.Add "recipe\NUMBER_PARA_183"
recipeTagGroup.Add "recipe\NUMBER_PARA_184"
recipeTagGroup.Add "recipe\NUMBER_PARA_185"
recipeTagGroup.Add "recipe\NUMBER_PARA_186"
recipeTagGroup.Add "recipe\NUMBER_PARA_187"
recipeTagGroup.Add "recipe\NUMBER_PARA_188"
recipeTagGroup.Add "recipe\NUMBER_PARA_189"
recipeTagGroup.Add "recipe\NUMBER_PARA_190"
recipeTagGroup.Add "recipe\NUMBER_PARA_191"
recipeTagGroup.Add "recipe\NUMBER_PARA_192"
recipeTagGroup.Add "recipe\NUMBER_PARA_193"
recipeTagGroup.Add "recipe\NUMBER_PARA_194"
recipeTagGroup.Add "recipe\NUMBER_PARA_195"
recipeTagGroup.Add "recipe\NUMBER_PARA_196"
recipeTagGroup.Add "recipe\NUMBER_PARA_197"
recipeTagGroup.Add "recipe\NUMBER_PARA_198"
recipeTagGroup.Add "recipe\NUMBER_PARA_199"
recipeTagGroup.Add "recipe\NUMBER_PARA_200"
recipeTagGroup.Add "recipe\NUMBER_PARA_201"
recipeTagGroup.Add "recipe\NUMBER_PARA_202"
recipeTagGroup.Add "recipe\NUMBER_PARA_203"
recipeTagGroup.Add "recipe\NUMBER_PARA_204"
recipeTagGroup.Add "recipe\NUMBER_PARA_205"
recipeTagGroup.Add "recipe\NUMBER_PARA_206"
recipeTagGroup.Add "recipe\NUMBER_PARA_207"
recipeTagGroup.Add "recipe\NUMBER_PARA_208"
recipeTagGroup.Add "recipe\NUMBER_PARA_209"
recipeTagGroup.Add "recipe\NUMBER_PARA_210"
recipeTagGroup.Add "recipe\NUMBER_PARA_211"
recipeTagGroup.Add "recipe\NUMBER_PARA_212"
recipeTagGroup.Add "recipe\NUMBER_PARA_213"
recipeTagGroup.Add "recipe\NUMBER_PARA_214"
recipeTagGroup.Add "recipe\NUMBER_PARA_215"
recipeTagGroup.Add "recipe\NUMBER_PARA_216"
recipeTagGroup.Add "recipe\NUMBER_PARA_217"
recipeTagGroup.Add "recipe\NUMBER_PARA_218"
recipeTagGroup.Add "recipe\NUMBER_PARA_219"
recipeTagGroup.Add "recipe\NUMBER_PARA_220"
recipeTagGroup.Add "recipe\NUMBER_PARA_221"
recipeTagGroup.Add "recipe\NUMBER_PARA_222"
recipeTagGroup.Add "recipe\NUMBER_PARA_223"
recipeTagGroup.Add "recipe\NUMBER_PARA_224"
recipeTagGroup.Add "recipe\NUMBER_PARA_225"
recipeTagGroup.Add "recipe\NUMBER_PARA_226"
recipeTagGroup.Add "recipe\NUMBER_PARA_227"
recipeTagGroup.Add "recipe\NUMBER_PARA_228"
recipeTagGroup.Add "recipe\NUMBER_PARA_229"
recipeTagGroup.Add "recipe\NUMBER_PARA_230"
recipeTagGroup.Add "recipe\NUMBER_PARA_231"
recipeTagGroup.Add "recipe\NUMBER_PARA_232"
recipeTagGroup.Add "recipe\NUMBER_PARA_233"
recipeTagGroup.Add "recipe\NUMBER_PARA_234"
recipeTagGroup.Add "recipe\NUMBER_PARA_235"
recipeTagGroup.Add "recipe\NUMBER_PARA_236"
recipeTagGroup.Add "recipe\NUMBER_PARA_237"
recipeTagGroup.Add "recipe\NUMBER_PARA_238"
recipeTagGroup.Add "recipe\NUMBER_PARA_239"
recipeTagGroup.Add "recipe\NUMBER_PARA_240"
recipeTagGroup.Add "recipe\NUMBER_PARA_241"
recipeTagGroup.Add "recipe\NUMBER_PARA_242"
recipeTagGroup.Add "recipe\NUMBER_PARA_243"
recipeTagGroup.Add "recipe\NUMBER_PARA_244"
recipeTagGroup.Add "recipe\NUMBER_PARA_245"
recipeTagGroup.Add "recipe\NUMBER_PARA_246"
recipeTagGroup.Add "recipe\NUMBER_PARA_247"
recipeTagGroup.Add "recipe\NUMBER_PARA_248"
recipeTagGroup.Add "recipe\NUMBER_PARA_249"
recipeTagGroup.Add "recipe\NUMBER_PARA_250"
recipeTagGroup.Add "recipe\NUMBER_PARA_251"
recipeTagGroup.Add "recipe\NUMBER_PARA_252"
recipeTagGroup.Add "recipe\NUMBER_PARA_253"
recipeTagGroup.Add "recipe\NUMBER_PARA_254"
recipeTagGroup.Add "recipe\NUMBER_PARA_255"
recipeTagGroup.Add "recipe\NUMBER_PARA_256"
recipeTagGroup.Add "recipe\NUMBER_PARA_257"
recipeTagGroup.Add "recipe\NUMBER_PARA_258"
recipeTagGroup.Add "recipe\NUMBER_PARA_259"
recipeTagGroup.Add "recipe\NUMBER_PARA_260"
recipeTagGroup.Add "recipe\NUMBER_PARA_261"
recipeTagGroup.Add "recipe\NUMBER_PARA_262"
recipeTagGroup.Add "recipe\NUMBER_PARA_263"
recipeTagGroup.Add "recipe\NUMBER_PARA_264"
recipeTagGroup.Add "recipe\NUMBER_PARA_265"
recipeTagGroup.Add "recipe\NUMBER_PARA_266"
recipeTagGroup.Add "recipe\NUMBER_PARA_267"
recipeTagGroup.Add "recipe\NUMBER_PARA_268"
recipeTagGroup.Add "recipe\NUMBER_PARA_269"
recipeTagGroup.Add "recipe\NUMBER_PARA_270"
recipeTagGroup.Add "recipe\NUMBER_PARA_271"
recipeTagGroup.Add "recipe\NUMBER_PARA_272"
recipeTagGroup.Add "recipe\NUMBER_PARA_273"
recipeTagGroup.Add "recipe\NUMBER_PARA_274"
recipeTagGroup.Add "recipe\NUMBER_PARA_275"
recipeTagGroup.Add "recipe\NUMBER_PARA_276"
recipeTagGroup.Add "recipe\NUMBER_PARA_277"
recipeTagGroup.Add "recipe\NUMBER_PARA_278"
recipeTagGroup.Add "recipe\NUMBER_PARA_279"
recipeTagGroup.Add "recipe\NUMBER_PARA_280"
recipeTagGroup.Add "recipe\NUMBER_PARA_281"
recipeTagGroup.Add "recipe\NUMBER_PARA_282"
recipeTagGroup.Add "recipe\NUMBER_PARA_283"
recipeTagGroup.Add "recipe\NUMBER_PARA_284"
recipeTagGroup.Add "recipe\NUMBER_PARA_285"
recipeTagGroup.Add "recipe\NUMBER_PARA_286"
recipeTagGroup.Add "recipe\NUMBER_PARA_287"
recipeTagGroup.Add "recipe\NUMBER_PARA_288"
recipeTagGroup.Add "recipe\NUMBER_PARA_289"
recipeTagGroup.Add "recipe\NUMBER_PARA_290"
recipeTagGroup.Add "recipe\NUMBER_PARA_291"
recipeTagGroup.Add "recipe\NUMBER_PARA_292"
recipeTagGroup.Add "recipe\NUMBER_PARA_293"
recipeTagGroup.Add "recipe\NUMBER_PARA_294"
recipeTagGroup.Add "recipe\NUMBER_PARA_295"
recipeTagGroup.Add "recipe\NUMBER_PARA_296"
recipeTagGroup.Add "recipe\NUMBER_PARA_297"
recipeTagGroup.Add "recipe\NUMBER_PARA_298"
recipeTagGroup.Add "recipe\NUMBER_PARA_299"
recipeTagGroup.Add "recipe\NUMBER_PARA_300"		
    End If
    Call ResetSelectedRecipeTags

End Function

Private Sub CreateRecipeBtn_Press()

recipeTagGroup.Item("recipe\RECIPE_NAME").Value = ""

recipeTagGroup.Item("recipe\SCADA_DATA_Batch_No").Value = ""

Call ResetSelectedRecipeTags

'Dim recipeTagGroup As TagGroup

'If recipeTagGroup Is Nothing Then

'    Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)

'

'        recipeTagGroup.Add "recipe\RECIPE_NAME"

'        recipeTagGroup.Add "recipe\SCADA_DATA_Batch_No"

'

'

'   End If

  

   

End Sub

Public Sub ResetSelectedRecipeTags()

    LogDiagnosticsMessage "reset hmi tag to 0 called"


    Dim i As Integer

     ' RESET STRING TAGS

    For i = 1 To 50

        recipeTagGroup.Item("recipe\STRING_PARA_" & i).Value = ""

    Next i

    ' RESET REAL( INT , FLOAT , ANALOG , NUMBER)

    For i = 1 To 300

        recipeTagGroup.Item("recipe\NUMBER_PARA_" & i).Value = 0

    Next i

End Sub

Private Sub DeleteRecipeBtn_Accepted(PerformerUser As String, PerformerCOMMENT As String, ApproverUser As String, ApproverCOMMENT As String)

'MsgBox "delete is clicked "

'    Dim recipeTagGroup As TagGroup

'

'   If recipeTagGroup Is Nothing Then

'    Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)

'        recipeTagGroup.Add "recipe\RECIPE_NAME"

'   End If

   'MsgBox "selected recipe name is : " & RecipeMasterTable.TextMatrix(RecipeMasterTable.Row, 1)

   recipeTagGroup.Item("recipe\RECIPE_NAME").Value = recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0)

    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection

    Dim sql As String

    Dim RECIPE_NAME As String

   

    RECIPE_NAME = recipeTagGroup.Item("recipe\RECIPE_NAME").Value

   

    If IsNull(RECIPE_NAME) Or Trim(RECIPE_NAME) = "" Or RECIPE_NAME = "RECIPE_NAME" Or RECIPE_NAME = "RECIPE NAME" Then

        MsgBox "Please select a recipe", vbExclamation, "Missing Recipe"

        Exit Sub

    End If

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"

    sql = "DELETE FROM RECIPE WHERE RECIPE_NAME = '" & RECIPE_NAME & "'"

    conn.Execute sql

   

    MsgBox "Recipe Deleted Successfully"

    Call LoadRecipeData

    conn.Close

    Set conn = Nothing

    Exit Sub

ErrHandler:

    LogDiagnosticsMessage _

        "Delete failed: " & Err.Description

    MsgBox _

        "Delete failed: " & Err.Description

End Sub

Private Sub Display_AnimationStart()

'----------------load recipe table data on start

        LoadRecipeData

      Call InitRecipeTagGroup

      'LoadRecipeData

End Sub

Private Function FormatColumnHeader(ByVal colName As String) As String

    Dim i As Integer

    Dim result As String

    Dim ch As String

    result = ""

    For i = 1 To Len(colName)

        ch = Mid(colName, i, 1)

        If i > 1 Then

            If ch Like "[A-Z]" Then

                result = result & " "

            End If

        End If

        result = result & ch

    Next i

    FormatColumnHeader = UCase(result)

End Function

'------------------- save sql to text file

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

'----------------load recipe data funtion

Public Sub LoadRecipeData()

    'On Error GoTo ErrHandler

    Dim conn As ADODB.Connection

    Dim Rs As ADODB.Recordset

    Dim sql As String

    Dim r As Long

    Dim c As Long

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"

        sql = "SELECT RECIPE_NAME,STATUS , CREATED_BY , CREATED_DATE ,APPROVED_BY , APPROVED_DATE , COMMENT , UPDATED_BY, UPDATED_DATE FROM RECIPE order by UPDATED_DATE desc"

        Set Rs = conn.Execute(sql)

        recipeMasterTable.Clear

   

    

        recipeMasterTable.Rows = 1

        recipeMasterTable.Cols = Rs.Fields.Count

        recipeMasterTable.ColWidth(0) = 5300   ' Recipe Name

        recipeMasterTable.ColWidth(1) = 4250   ' STATUS

        recipeMasterTable.ColWidth(2) = 1700   ' CREATED_BY

        recipeMasterTable.ColWidth(3) = 2300   ' CREATED_DATE

        recipeMasterTable.ColWidth(4) = 1900   ' APPROVED_BY

        recipeMasterTable.ColWidth(5) = 2400   ' APPROVED_DATE

        recipeMasterTable.ColWidth(6) = 6500   ' COMMENT

        recipeMasterTable.ColWidth(7) = 1700   ' UPDATED_BY

        recipeMasterTable.ColWidth(8) = 2300   ' UPDATED_DATE

     ' Column Names

        For c = 0 To recipeMasterTable.Cols - 1

            recipeMasterTable.TextMatrix(0, c) = _

            FormatColumnHeader(Rs.Fields(c).Name)

        Next c

       

        '==========================

        ' Grid Formatting

        '==========================

       

        With recipeMasterTable

       

        ' Entire grid font size

        .Font.Name = "Arial"

        .Font.Size = 13

        .Font.Bold = False

       

        End With

       

        ' Header formatting

        For c = 0 To recipeMasterTable.Cols - 1

       

        recipeMasterTable.Row = 0

        recipeMasterTable.Col = c

       

        recipeMasterTable.CellFontBold = True

        recipeMasterTable.CellAlignment = flexAlignCenterCenter

       

        ' Purple header

        recipeMasterTable.CellBackColor = RGB(199, 210, 232)

        recipeMasterTable.CellForeColor = vbBlack

        

        Next c

       

        ' Data alignment (left)

        Dim t As Long

       

        For t = 1 To recipeMasterTable.Rows - 1

       

        For c = 0 To recipeMasterTable.Cols - 1

       

        recipeMasterTable.Row = t

        recipeMasterTable.Col = c

       

        recipeMasterTable.CellAlignment = flexAlignLeftCenter

       

        Next c

       

        Next t

       

    If Rs.EOF Then

        LogDiagnosticsMessage _

            "No records found in RECIPEScreen"

        Rs.Close

        conn.Close

    Exit Sub

    End If

   

    ' Data Rows

    Do Until Rs.EOF

        recipeMasterTable.Rows = recipeMasterTable.Rows + 1

        r = recipeMasterTable.Rows - 1

        For c = 0 To recipeMasterTable.Cols - 1

            '-- RecipeMasterTable.TextMatrix(r, c) = _

                rs.Fields(c).Value & ""

             recipeMasterTable.TextMatrix(r, c) = GetDisplayValue(Rs.Fields(c).Name, Rs.Fields(c).Value & "")

        Next c

        Rs.MoveNext

    Loop

    Rs.Close

    conn.Close

    Set Rs = Nothing

    Set conn = Nothing

    LogDiagnosticsMessage _

        "Recipe data loaded successfully"

    Exit Sub

ErrHandler:

    LogDiagnosticsMessage _

        "LoadRecipe Error : " & Err.Description

    MsgBox _

        "LoadRecipe Error : " & Err.Description

End Sub

Private Sub Group2_Click()

End Sub

Private Sub recipeMasterTable_Click()

End Sub

Private Sub RefreshTableBtn_Released()

    LoadRecipeData

End Sub

'-------------------------------------get display value for 0 and 1

Private Function GetDisplayValue( _

    ByVal columnName As String, _

    ByVal rawValue As Variant) As String

    Select Case LCase(columnName)

        Case "STATUS"

            Select Case CLng(rawValue)

                Case 0

                    GetDisplayValue = "Not Approved"

                Case 1

                    GetDisplayValue = "Approved"

                Case 2

                    GetDisplayValue = "Pending Approval after Recipe Update"

                Case 3

                    GetDisplayValue = "DISPLAYd"

                Case Else

                    GetDisplayValue = "Unknown"

            End Select

           

        Case Else

            GetDisplayValue = CStr(rawValue)

    End Select

End Function

Private Sub UpdateRecipeVERSION()
    LogDiagnosticsMessage("UPDATE RECIPE VERSION CALLED")
    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection

    Dim sql As String

    Dim RECIPE_NAME As String

    If recipeMasterTable.Row <= 0 Then

        MsgBox "Please select a recipe"

        Exit Sub

    End If

    RECIPE_NAME = recipeMasterTable.TextMatrix( _

                    recipeMasterTable.Row, 0)

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"

    sql = "UPDATE RECIPE " & _

          "SET VERSION = VERSION + 1 " & _

          "WHERE RECIPE_NAME = '" & RECIPE_NAME & "'"

    LogDiagnosticsMessage sql

    conn.Execute sql

    MsgBox "Recipe VERSION updated successfully"

    LoadRecipeData

    conn.Close

    Set conn = Nothing

    Exit Sub

ErrHandler:

    LogDiagnosticsMessage _

        "VERSION update failed : " & Err.Description

    MsgBox _

        "VERSION update failed : " & Err.Description

End Sub

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


  

Private Sub ViewRecipeBtn1_Press()

    LogDiagnosticsMessage(" VIEW BUTTON PRESSED")

    recipeTagGroup.Item("recipe\RECIPE_NAME").Value = recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0)

   'MsgBox "selected recipe name is : " & recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0)

   '

   If (StrComp(recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0), "RECIPE_NAME", vbBinaryCompare) = 0) Or (StrComp(recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0), "RECIPE NAME", vbBinaryCompare) = 0) Then

        MsgBox ("Please select the recipe")

   Else

    Call renderRecipeParameters(recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0))

     LogDiagnosticsMessage (" getting Display RECIPE_PARAMETERS_WAX")

    ExecuteCommand "Display RECIPE_PARAMETERS_WAX"

    LogDiagnosticsMessage (" getting Display RECIPE_PARAMETERS_WAX")

    'ShowDisplay "RECIPE_PARAMETERS_WAX"

   End If

  

End Sub

Private Sub ViewRecipeBtn1_Released()

 

   

End Sub