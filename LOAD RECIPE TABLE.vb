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
'-------FORMAT COLUMN Header
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
'-------------------------------------GET STRING FOR STATUS COLUMN

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