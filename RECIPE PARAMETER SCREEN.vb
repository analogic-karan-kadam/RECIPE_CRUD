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
DISABLE"         '---------digital data type
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
End Function


 


 

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


 

Private Sub approveBtn1_Accepted(PerformerUser As String, PerformerComment As String, ApproverUser As String, ApproverComment As String)

    Call updateRecipeStatus(1, recipeTagGroup.Item("recipe\RECIPE_NAME").Value, PerformerComment, PerformerUser)

    ExecuteCommand "Display RecipeMasterScreen"

End Sub


 

Private Sub updateRecipeStatus(RECIPE_STATUS As Integer, RECIPE_NAME As String, COMMENT As String, PerformerUser As String)


 

    Dim conn As ADODB.Connection

    Dim sql As String

   ' Dim recipeTagGroup As TagGroup

    'Set recipeTagGroup = InitRecipeTagGroup

   

        

    

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


 

Private Sub batchNumberInput_Change()


 

End Sub


 

Private Sub Button1_Released()


 

If recipeTagGroup.Item("recipe\Download").Value Then

     MsgBox "Recipe Downloaded SuccessFully"

End If


 

End Sub


 

Private Sub Display_AnimationStart()

   Call InitRecipeTagGroup
    LogDiagnosticsMessage("RENDER ALL PARA FOR RECIPE : " & recipeTagGroup.Item("recipe\RECIPE_NAME").Value)
   Call renderRecipeParameters(recipeTagGroup.Item("recipe\RECIPE_NAME").Value)

End Sub


 

Private Sub rejectBtn__Accepted(PerformerUser As String, PerformerComment As String, ApproveUser As String, ApproveComment As String)


 

End Sub


Private Sub saveRecipetoDB(user As String)
LogDiagnosticsMessage ("Save function called")

    If CheckDuplicate(recipeTagGroup.Item("recipe\RECIPE_NAME").Value) Then

        MsgBox "Recipe Already Exists Choose Different Name for the Recipe"

        RECIPE_NAMEInput.Value = ""

        batchNumberInput.Value = ""

        Exit Sub

    End If  

    'On Error GoTo ErrHandler

    Dim conn As ADODB.Connection

    Dim sqlQuery As String

   

    ' --- 1. VALIDATION: Check if objects exist and have text ---

    ' Using Controls ensures VBA binds to the current display items

    If Trim(Me.RECIPE_NAMEInput.Value) = "" Then

        MsgBox "Error: Please Enter Recipe Name", vbCritical, "Missing Data"

        Exit Sub

    End If


    ' --- 3. DATABASE CONNECTION ---

    Set conn = New ADODB.Connection

    conn.Open _

        "Provider=MSOLEDBSQL;" & _

        "Data Source=DESKTOP-LA6V9TK\WINCCPLUSMIG2014;" & _

        "Initial Catalog=vbaLiveProjects;" & _

        "User ID=sa;" & _

        "Password=sa;" & _

        "TrustServerCertificate=yes;"

       

        LogDiagnosticsMessage ("Db Connection made")


    ' --- META DATA VARIABLES
    
	LogDiagnosticsMessage ("SAVE RECIPE CALLED BY USRE : "&user)
  
     If recipeTagGroup.Item("recipe\RECIPE_NAME").Value = "" Then

        MsgBox "Please Enter Recipe Name or Press enter after entering name"

        'RECIPE_NAMEInput.Value = ""

        Exit Sub

    End If


    ' ==============================================================

    ' COLUMN GROUPS

    ' ==============================================================

    Dim colsMeta As String

    colsMeta = "RECIPE_NAME , CREATED_BY, CREATED_DATE ,UPDATED_DATE "



    Dim colsString As String

    colsString ="STRING_PARA_1, STRING_PARA_2, STRING_PARA_3, STRING_PARA_4, " & _
				"STRING_PARA_5, STRING_PARA_6, STRING_PARA_7, STRING_PARA_8, " & _
				"STRING_PARA_9, STRING_PARA_10, STRING_PARA_11, STRING_PARA_12, " & _
				"STRING_PARA_13, STRING_PARA_14, STRING_PARA_15, STRING_PARA_16, " & _
				"STRING_PARA_17, STRING_PARA_18, STRING_PARA_19, STRING_PARA_20, " & _
				"STRING_PARA_21, STRING_PARA_22, STRING_PARA_23, STRING_PARA_24, " & _
				"STRING_PARA_25, STRING_PARA_26, STRING_PARA_27, STRING_PARA_28, " & _
				"STRING_PARA_29, STRING_PARA_30, STRING_PARA_31, STRING_PARA_32, " & _
				"STRING_PARA_33, STRING_PARA_34, STRING_PARA_35, STRING_PARA_36, " & _
				"STRING_PARA_37, STRING_PARA_38, STRING_PARA_39, STRING_PARA_40, " & _
				"STRING_PARA_41, STRING_PARA_42, STRING_PARA_43, STRING_PARA_44, " & _
				"STRING_PARA_45, STRING_PARA_46, STRING_PARA_47, STRING_PARA_48, " & _
				"STRING_PARA_49, STRING_PARA_50"

 

    Dim colsNumber1to50 As String
    colsNumber1to50 =   "NUMBER_PARA_1, NUMBER_PARA_2, NUMBER_PARA_3, NUMBER_PARA_4, " & _
						"NUMBER_PARA_5, NUMBER_PARA_6, NUMBER_PARA_7, NUMBER_PARA_8, " & _
						"NUMBER_PARA_9, NUMBER_PARA_10, NUMBER_PARA_11, NUMBER_PARA_12, " & _
						"NUMBER_PARA_13, NUMBER_PARA_14, NUMBER_PARA_15, NUMBER_PARA_16, " & _
						"NUMBER_PARA_17, NUMBER_PARA_18, NUMBER_PARA_19, NUMBER_PARA_20, " & _
						"NUMBER_PARA_21, NUMBER_PARA_22, NUMBER_PARA_23, NUMBER_PARA_24, " & _
						"NUMBER_PARA_25, NUMBER_PARA_26, NUMBER_PARA_27, NUMBER_PARA_28, " & _
						"NUMBER_PARA_29, NUMBER_PARA_30, NUMBER_PARA_31, NUMBER_PARA_32, " & _
						"NUMBER_PARA_33, NUMBER_PARA_34, NUMBER_PARA_35, NUMBER_PARA_36, " & _
						"NUMBER_PARA_37, NUMBER_PARA_38, NUMBER_PARA_39, NUMBER_PARA_40, " & _
						"NUMBER_PARA_41, NUMBER_PARA_42, NUMBER_PARA_43, NUMBER_PARA_44, " & _
						"NUMBER_PARA_45, NUMBER_PARA_46, NUMBER_PARA_47, NUMBER_PARA_48, " & _
						"NUMBER_PARA_49, NUMBER_PARA_50"

 
 
    Dim colsNumber51to100 As String
    colsNumber51to100 =     "NUMBER_PARA_51, NUMBER_PARA_52, NUMBER_PARA_53, NUMBER_PARA_54, " & _
							"NUMBER_PARA_55, NUMBER_PARA_56, NUMBER_PARA_57, NUMBER_PARA_58, " & _
							"NUMBER_PARA_59, NUMBER_PARA_60, NUMBER_PARA_61, NUMBER_PARA_62, " & _
							"NUMBER_PARA_63, NUMBER_PARA_64, NUMBER_PARA_65, NUMBER_PARA_66, " & _
							"NUMBER_PARA_67, NUMBER_PARA_68, NUMBER_PARA_69, NUMBER_PARA_70, " & _
							"NUMBER_PARA_71, NUMBER_PARA_72, NUMBER_PARA_73, NUMBER_PARA_74, " & _
							"NUMBER_PARA_75, NUMBER_PARA_76, NUMBER_PARA_77, NUMBER_PARA_78, " & _
							"NUMBER_PARA_79, NUMBER_PARA_80, NUMBER_PARA_81, NUMBER_PARA_82, " & _
							"NUMBER_PARA_83, NUMBER_PARA_84, NUMBER_PARA_85, NUMBER_PARA_86, " & _
							"NUMBER_PARA_87, NUMBER_PARA_88, NUMBER_PARA_89, NUMBER_PARA_90, " & _
							"NUMBER_PARA_91, NUMBER_PARA_92, NUMBER_PARA_93, NUMBER_PARA_94, " & _
							"NUMBER_PARA_95, NUMBER_PARA_96, NUMBER_PARA_97, NUMBER_PARA_98, " & _
							"NUMBER_PARA_99, NUMBER_PARA_100"

    Dim colsNumber101to150 As String
    colsNumber101to150 ="NUMBER_PARA_101, NUMBER_PARA_102, NUMBER_PARA_103, NUMBER_PARA_104, " & _
						"NUMBER_PARA_105, NUMBER_PARA_106, NUMBER_PARA_107, NUMBER_PARA_108, " & _
						"NUMBER_PARA_109, NUMBER_PARA_110, NUMBER_PARA_111, NUMBER_PARA_112, " & _
						"NUMBER_PARA_113, NUMBER_PARA_114, NUMBER_PARA_115, NUMBER_PARA_116, " & _
						"NUMBER_PARA_117, NUMBER_PARA_118, NUMBER_PARA_119, NUMBER_PARA_120, " & _
						"NUMBER_PARA_121, NUMBER_PARA_122, NUMBER_PARA_123, NUMBER_PARA_124, " & _
						"NUMBER_PARA_125, NUMBER_PARA_126, NUMBER_PARA_127, NUMBER_PARA_128, " & _
						"NUMBER_PARA_129, NUMBER_PARA_130, NUMBER_PARA_131, NUMBER_PARA_132, " & _
						"NUMBER_PARA_133, NUMBER_PARA_134, NUMBER_PARA_135, NUMBER_PARA_136, " & _
						"NUMBER_PARA_137, NUMBER_PARA_138, NUMBER_PARA_139, NUMBER_PARA_140, " & _
						"NUMBER_PARA_141, NUMBER_PARA_142, NUMBER_PARA_143, NUMBER_PARA_144, " & _
						"NUMBER_PARA_145, NUMBER_PARA_146, NUMBER_PARA_147, NUMBER_PARA_148, " & _
						"NUMBER_PARA_149, NUMBER_PARA_150"

    Dim colsNumber151to200 As String
    colsNumber151to200 ="NUMBER_PARA_151, NUMBER_PARA_152, NUMBER_PARA_153, NUMBER_PARA_154, " & _
						"NUMBER_PARA_155, NUMBER_PARA_156, NUMBER_PARA_157, NUMBER_PARA_158, " & _
						"NUMBER_PARA_159, NUMBER_PARA_160, NUMBER_PARA_161, NUMBER_PARA_162, " & _
						"NUMBER_PARA_163, NUMBER_PARA_164, NUMBER_PARA_165, NUMBER_PARA_166, " & _
						"NUMBER_PARA_167, NUMBER_PARA_168, NUMBER_PARA_169, NUMBER_PARA_170, " & _
						"NUMBER_PARA_171, NUMBER_PARA_172, NUMBER_PARA_173, NUMBER_PARA_174, " & _
						"NUMBER_PARA_175, NUMBER_PARA_176, NUMBER_PARA_177, NUMBER_PARA_178, " & _
						"NUMBER_PARA_179, NUMBER_PARA_180, NUMBER_PARA_181, NUMBER_PARA_182, " & _
						"NUMBER_PARA_183, NUMBER_PARA_184, NUMBER_PARA_185, NUMBER_PARA_186, " & _
						"NUMBER_PARA_187, NUMBER_PARA_188, NUMBER_PARA_189, NUMBER_PARA_190, " & _
						"NUMBER_PARA_191, NUMBER_PARA_192, NUMBER_PARA_193, NUMBER_PARA_194, " & _
						"NUMBER_PARA_195, NUMBER_PARA_196, NUMBER_PARA_197, NUMBER_PARA_198, " & _
						"NUMBER_PARA_199, NUMBER_PARA_200"


    Dim colsNumber201to250 As String
    colsNumber201to250 ="NUMBER_PARA_201, NUMBER_PARA_202, NUMBER_PARA_203, NUMBER_PARA_204, " & _
						"NUMBER_PARA_205, NUMBER_PARA_206, NUMBER_PARA_207, NUMBER_PARA_208, " & _
						"NUMBER_PARA_209, NUMBER_PARA_210, NUMBER_PARA_211, NUMBER_PARA_212, " & _
						"NUMBER_PARA_213, NUMBER_PARA_214, NUMBER_PARA_215, NUMBER_PARA_216, " & _
						"NUMBER_PARA_217, NUMBER_PARA_218, NUMBER_PARA_219, NUMBER_PARA_220, " & _
						"NUMBER_PARA_221, NUMBER_PARA_222, NUMBER_PARA_223, NUMBER_PARA_224, " & _
						"NUMBER_PARA_225, NUMBER_PARA_226, NUMBER_PARA_227, NUMBER_PARA_228, " & _
						"NUMBER_PARA_229, NUMBER_PARA_230, NUMBER_PARA_231, NUMBER_PARA_232, " & _
						"NUMBER_PARA_233, NUMBER_PARA_234, NUMBER_PARA_235, NUMBER_PARA_236, " & _
						"NUMBER_PARA_237, NUMBER_PARA_238, NUMBER_PARA_239, NUMBER_PARA_240, " & _
						"NUMBER_PARA_241, NUMBER_PARA_242, NUMBER_PARA_243, NUMBER_PARA_244, " & _
						"NUMBER_PARA_245, NUMBER_PARA_246, NUMBER_PARA_247, NUMBER_PARA_248, " & _
						"NUMBER_PARA_249, NUMBER_PARA_250"


    Dim colsNumber251to300 As String
    colsNumber251to300 ="NUMBER_PARA_251, NUMBER_PARA_252, NUMBER_PARA_253, NUMBER_PARA_254, " & _
						"NUMBER_PARA_255, NUMBER_PARA_256, NUMBER_PARA_257, NUMBER_PARA_258, " & _
						"NUMBER_PARA_259, NUMBER_PARA_260, NUMBER_PARA_261, NUMBER_PARA_262, " & _
						"NUMBER_PARA_263, NUMBER_PARA_264, NUMBER_PARA_265, NUMBER_PARA_266, " & _
						"NUMBER_PARA_267, NUMBER_PARA_268, NUMBER_PARA_269, NUMBER_PARA_270, " & _
						"NUMBER_PARA_271, NUMBER_PARA_272, NUMBER_PARA_273, NUMBER_PARA_274, " & _
						"NUMBER_PARA_275, NUMBER_PARA_276, NUMBER_PARA_277, NUMBER_PARA_278, " & _
						"NUMBER_PARA_279, NUMBER_PARA_280, NUMBER_PARA_281, NUMBER_PARA_282, " & _
						"NUMBER_PARA_283, NUMBER_PARA_284, NUMBER_PARA_285, NUMBER_PARA_286, " & _
						"NUMBER_PARA_287, NUMBER_PARA_288, NUMBER_PARA_289, NUMBER_PARA_290, " & _
						"NUMBER_PARA_291, NUMBER_PARA_292, NUMBER_PARA_293, NUMBER_PARA_294, " & _
						"NUMBER_PARA_295, NUMBER_PARA_296, NUMBER_PARA_297, NUMBER_PARA_298, " & _
						"NUMBER_PARA_299, NUMBER_PARA_300"

    LogDiagnosticsMessage ("Concatenation of SQL query")

    ' --- Final COLUMN concatenation ---

    Dim sqlCols As String

    sqlCols = "INSERT INTO RECIPE (" & _

                  colsMeta & ", " & _

                  colsString & ", " & _

                  colsNumber1to50 & ", " & _

                  colsNumber51to100 & ", " & _

                  colsNumber101to150 & ", " & _

                  colsNumber151to200 & ", " & _

                  colsNumber201to250 & ", " & _

                  colsNumber251to300 & ") "

    LogDiagnosticsMessage ("sqlCols " & sqlCols)

    ' ==============================================================

    Dim valsMeta As String

   valsMeta =  "'" & STRING_RECIPE_NAME & "', " & _

			   "'" & PerformerUser & "', " & _

			   "GETDATE(), " & _

			   "GETDATE()"


    Dim valsSring As String
    valsSring = recipeTagGroup.Item("recipe\STRING_PARA_" & 1).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 2).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 3).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 4).Value
	valsSring = valsSring & ", "recipeTagGroup.Item("recipe\STRING_PARA_" & 5).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 6).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 7).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 8).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 9).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 10).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 11).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 12).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 13).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 14).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 15).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 16).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 17).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 18).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 19).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 20).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 21).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 22).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 23).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 24).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 25).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 26).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 27).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 28).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 29).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 30).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 31).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 32).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 33).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 34).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 35).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 36).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 37).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 38).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 39).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 40).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 41).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 42).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 43).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 44).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 45).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 46).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 47).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 48).Value
	valsSring = valsSring & ", " recipeTagGroup.Item("recipe\STRING_PARA_" & 49).Value & ", " & recipeTagGroup.Item("recipe\STRING_PARA_" & 50).Value

	dim valsNumber1to50 as string
	valsNumber1to50 =CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 1).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 2).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 3).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 4).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 5).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 6).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 7).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 8).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 9).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 10).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 11).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 12).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 13).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 14).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 15).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 16).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 17).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 18).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 19).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 20).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 21).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 22).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 23).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 24).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 25).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 26).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 27).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 28).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 29).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 30).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 31).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 32).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 33).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 34).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 35).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 36).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 37).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 38).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 39).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 40).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 41).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 42).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 43).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 44).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 45).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 46).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 47).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 48).Value) & "," & _
				CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 49).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 50).Value)


	dim valsNumber51to100 as string
	valsNumber51to100 =CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 51).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 52).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 53).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 54).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 55).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 56).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 57).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 58).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 59).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 60).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 61).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 62).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 63).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 64).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 65).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 66).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 67).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 68).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 69).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 70).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 71).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 72).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 73).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 74).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 75).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 76).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 77).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 78).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 79).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 80).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 81).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 82).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 83).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 84).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 85).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 86).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 87).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 88).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 89).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 90).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 91).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 92).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 93).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 94).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 95).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 96).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 97).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 98).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 99).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 100).Value)



	dim valsNumber101to150 as string
	valsNumber101to150 = CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 101).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 102).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 103).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 104).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 105).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 106).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 107).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 108).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 109).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 110).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 111).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 112).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 113).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 114).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 115).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 116).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 117).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 118).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 119).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 120).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 121).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 122).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 123).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 124).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 125).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 126).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 127).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 128).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 129).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 130).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 131).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 132).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 133).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 134).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 135).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 136).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 137).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 138).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 139).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 140).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 141).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 142).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 143).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 144).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 145).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 146).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 147).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 148).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 149).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 150).Value)
	
	
	dim valsNumber151to200 as string
	valsNumber151to200 =    CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 150).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 151).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 152).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 153).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 154).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 155).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 156).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 157).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 158).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 159).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 160).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 161).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 162).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 163).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 164).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 165).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 166).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 167).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 168).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 169).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 170).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 171).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 172).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 173).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 174).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 175).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 176).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 177).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 178).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 179).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 180).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 181).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 182).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 183).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 184).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 185).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 186).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 187).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 188).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 189).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 190).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 191).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 192).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 193).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 194).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 195).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 196).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 197).Value) & "," & _
							CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 198).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 199).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 200).Value)
	
	
	dim valsNumber201to250 as string
	valsNumber201to250 =CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 201).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 202).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 203).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 204).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 205).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 206).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 207).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 208).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 209).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 210).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 211).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 212).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 213).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 214).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 215).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 216).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 217).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 218).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 219).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 220).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 221).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 222).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 223).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 224).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 225).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 226).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 227).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 228).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 229).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 230).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 231).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 232).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 233).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 234).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 235).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 236).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 237).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 238).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 239).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 240).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 241).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 242).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 243).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 244).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 245).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 246).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 247).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 248).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 249).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 250).Value)
	
	
	dim valsNumber251to300 as string
	valsNumber251to300 =CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 251).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 252).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 253).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 254).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 255).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 256).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 257).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 258).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 259).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 260).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 261).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 262).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 263).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 264).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 265).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 266).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 267).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 268).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 269).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 270).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 271).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 272).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 273).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 274).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 275).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 276).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 277).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 278).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 279).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 280).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 281).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 282).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 283).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 284).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 285).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 286).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 287).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 288).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 289).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 290).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 291).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 292).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 293).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 294).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 295).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 296).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 297).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 298).Value) & "," & _
						CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 299).Value) & ","CLng(recipeTagGroup.Item("recipe\NUMBER_PARA_" & 300).Value)
	


    sqlVals = "VALUES (" & _

                  valsMeta & ", " & _
				  
				  valsSring & ", " & _

                  valsNumber1to50 & ", " & _

                  valsNumber51to100 & ", " & _

                  valsNumber101to150 & ", " & _

                  valsNumber151to200 & ", " & _

                  valsNumber201to250 & ", " & _

                  valsNumber251to300 & ")"



    ' ==============================================================

    ' EXECUTE

     LogDiagnosticsMessage ("Concatenation end of SQL query")

    

     LogDiagnosticsMessage ("sqlVals " & sqlVals)

     sqlQuery = sqlCols & sqlVals

     LogDiagnosticsMessage ("SQL query built successfully")


 

          'MsgBox (sqlQuery)

    'LogDiagnosticsMessage (sqlQuery)

   


 

    ' --- 5. EXECUTE AND CLEANUP ---

    conn.Execute sqlQuery

   

    MsgBox "Recipe saved successfully", vbInformation, "Success"

    LogDiagnosticsMessage "Recipe saved successfully"


 

    conn.Close

    Set conn = Nothing

    Exit Sub


 

ErrHandler:

    LogDiagnosticsMessage "Save Recipe Error : " & Err.Description

    MsgBox "Save Recipe Error : " & Err.Description, vbCritical, "Database Error"

   

    If Not conn Is Nothing Then

        If conn.State = 1 Then conn.Close

        Set conn = Nothing

    End If

End Sub


 

Private Sub rejectBtn1_Accepted(PerformerUser As String, PerformerComment As String, ApproverUser As String, ApproverComment As String)

    Call updateRecipeStatus(0, recipeTagGroup.Item("recipe\RECIPE_NAME").Value, PerformerComment, PerformerUser)

    ExecuteCommand "Display RecipeMasterScreen"

End Sub


 

Private Sub rejectBtn1_Released()


 

End Sub


 

Private Sub saveToSqlBtn1_Accepted(PerformerUser As String, PerformerComment As String, ApproverUser As String, ApproverComment As String)

    Call saveRecipetoDB(PerformerUser)

    ExecuteCommand "Display RecipeMasterScreen"


 

End Sub


 


 

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


 


 

Private Sub updateBtn1_Accepted(PerformerUser As String, PerformerComment As String, ApproverUser As String, ApproverComment As String)

    'MsgBox "Make Sure You Pressed Enter Button After adding "

    Call updateRecipetoDB(PerformerUser, PerformerComment)

    ExecuteCommand "Display RecipeMasterScreen"

End Sub

Private Sub Button7_Released()


 

End Sub

Private Sub Button6_Released()


 

End Sub

Private Sub Button4_Released()


 

End Sub

Private Sub Button3_Released()


 

End Sub

Private Sub Button2_Released()


 

End Sub

Private Sub Button8_Released()


 

End Sub

Private Sub Button5_Released()


 

End Sub