'-----------------------------------------------------------------
' dbLib - database Library
'-----------------------------------------------------------------
' R.Bettinelli                                               2014
'-----------------------------------------------------------------
' 02/11/2013 - Added Insert Functionality.
' 10/09/2013 - Universal 'Convert.To' Added, dbNull Moved. 
'-----------------------------------------------------------------
Imports System.Data.SqlClient
Imports System.Globalization

Module DbLibIo

#Region "Connection"

    Private Function GetConnectionIO() As String
        Return "Data Source=IO-IIS03;Initial Catalog=dbDone;User Id=Test1;Password=pass@word1;MultipleActiveResultSets=true;"
    End Function

#End Region

#Region "Special Pull"
    Function GetStaffNo(txt As String) As Integer
        GetStaffNo = 0
        Dim var() As String = {"user"}
        Dim val() As Object = {txt}
        Dim dbDs As DataSet = spx_Uni("spg_getUserNo", var, val)
        If dbDs.Tables(0).Rows.Count > 0 Then
            GetStaffNo = CInt(dbDs.Tables(0).Rows(0)("stf_ID").ToString)
        End If
    End Function
#End Region

#Region "Get Methods"

    '----------------------------------------------------------------------------------
    '- spg_UniversalPull - In/Out
    '----------------------------------------------------------------------------------
    Function Spx_Uni(sp As String) As DataSet
        Dim var() As String = {}
        Dim val() As Object = {}
        Spx_Uni = spx_UniPull(sp, var, val, False)
    End Function

    '----------------------------------------------------------------------------------
    '- spg_UniversalPull - In/Out
    '----------------------------------------------------------------------------------
    Function Spx_Uni(sp As String, var() As String, val() As Object) As DataSet
        Spx_Uni = spx_UniPull(sp, var, val, False)
    End Function

    '----------------------------------------------------------------------------------
    '- spg_UniversalPull - Time
    '----------------------------------------------------------------------------------
    Function Spx_UniZero(sp As String, var() As String, val() As Object) As DataSet
        Spx_UniZero = spx_UniPull(sp, var, val, True)
    End Function


    Public Function CheckType(Of T)(ByVal arg As T) As Type
        Return GetType(T)
    End Function


    '----------------------------------------------------------------------------------
    '- spg_UniversalPull 
    '----------------------------------------------------------------------------------
    Function Spx_UniPull(sp As String, inVar() As String, inVal() As Object, acceptZero As Boolean) As DataSet
        Dim dbConn As SqlConnection '= New SqlConnection
        Dim dbCmd As SqlCommand
        Dim dbPar As SqlParameter
        Dim dbAdp As SqlDataAdapter
        Dim dbDs As New DataSet With {
            .Locale = CultureInfo.GetCultureInfo("en-US")
        }

        dbConn = New SqlConnection(GetConnectionIO())

        dbConn.Open()
        dbCmd = dbConn.CreateCommand()
        dbCmd.CommandText = sp
        dbCmd.CommandType = CommandType.StoredProcedure
        dbCmd.CommandTimeout = 300

        If inVar.Length > 0 Then
            For x As Integer = 0 To inVar.Length - 1
                Dim objct As Object = inVal(x)
                Dim sel As String = objct.GetType.ToString
                dbPar = New SqlParameter With {
                    .ParameterName = "@" + inVar(x),
                    .Value = DBNull.Value
                }
                If inVal(x) IsNot Nothing Then
                    Select Case sel
                        Case "System.Boolean"
                            dbPar.SqlDbType = SqlDbType.Bit
                            dbPar.Value = Convert.ToBoolean(inVal(x))
                        Case "System.DateTime", "System.Date"
                            If IsDate(inVal(x)) Then
                                Dim dtest As DateTime = Nothing
                                If inVal(x) <> dtest Then
                                    dbPar.Value = inVal(x)
                                End If
                                If Year(inVal(x)) = 1980 Then
                                    dbPar.Value = DBNull.Value
                                End If
                            End If
                            dbPar.SqlDbType = SqlDbType.DateTime
                        Case "System.Int32"

                            If acceptZero Then
                                dbPar.Value = Convert.ToInt32(inVal(x))
                            Else
                                If Convert.ToInt32(inVal(x)) <> 0 Then
                                    dbPar.Value = Convert.ToInt32(inVal(x))
                                End If
                            End If
                            dbPar.SqlDbType = SqlDbType.Int
                        Case "System.Single"
                            dbPar.Value = Convert.ToSingle(inVal(x))
                            dbPar.SqlDbType = SqlDbType.Float
                        Case "System.Double"
                            If acceptZero Then
                                dbPar.Value = Convert.ToDouble(inVal(x))
                            Else
                                If Convert.ToInt32(inVal(x)) <> 0 Then
                                    dbPar.Value = Convert.ToDouble(inVal(x))
                                End If
                            End If
                            dbPar.SqlDbType = SqlDbType.Float
                        Case "System.String"
                            If acceptZero Then
                                dbPar.Value = Convert.ToString(inVal(x))
                            Else
                                If CStr(inVal(x)).Trim <> "" Then
                                    dbPar.Value = Convert.ToString(inVal(x))
                                End If
                            End If
                            dbPar.SqlDbType = SqlDbType.NVarChar

                    End Select
                    dbPar.Direction = ParameterDirection.Input
                    dbCmd.Parameters.Add(dbPar)
                End If
            Next
        End If
        dbAdp = New SqlDataAdapter(dbCmd)
        dbAdp.Fill(dbDs)
        dbAdp.Dispose()
        dbConn.Close()
        Spx_UniPull = dbDs
    End Function

    Public Function GetPageLine(pge As String) As List(Of String)

        Dim fl As List(Of String) = New List(Of String)
        Dim varX() As String = {"Pge"}
        Dim valX() As Object = {pge}
        Dim dbDs As DataSet = Spx_Uni("spg_Main", varX, valX)
        If dbDs.Tables(0).Rows.Count > 0 Then
            Dim dbRow As DataRow = dbDs.Tables(0).Rows(0)
            fl.Add(dbRow("Title").ToString)
            fl.Add(HttpContext.Current.Server.HtmlDecode(dbRow("Detail").ToString))
            fl.Add(dbRow("Img").ToString)
        Else
            fl.Add("No Tite Data Set")
            fl.Add("No Detail Data Set")
            fl.Add("")
        End If
        Return fl
    End Function


#End Region

End Module
