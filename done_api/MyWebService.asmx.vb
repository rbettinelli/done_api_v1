Imports System.ComponentModel
Imports System.Web.Script.Services
Imports System.Web.Services
Imports Newtonsoft.Json

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class MyWebService
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Sub DNAC(dna As String)

        Dim DCall As New DNACall
        DCall.Decode(dna)

        Dim v1 As String() = DCall.v1.ToArray
        Dim v2 As Object() = DCall.v2.ToArray
        Dim dbDs As DataSet = Spx_UniZero(DCall.SP, v1, v2)

        JsonOutput(dbDs)

    End Sub

    <WebMethod()>
    Public Sub DNAEncode(dna As String)

        Dim DCall As New DNACall
        Dim str As String = DCall.Encode(dna)

        'Dim v1 As String() = DCall.v1.ToArray
        'Dim v2 As Object() = DCall.v2.ToArray
        'Dim dbDs As DataSet = Spx_UniZero(DCall.SP, v1, v2)

        JsonOutput(str)

    End Sub

    <WebMethod()>
    Public Sub DNADecode(dna As String)

        Dim DCall As New DNACall
        Dim str As String = DCall.DecodeTxt(dna)

        'Dim v1 As String() = DCall.v1.ToArray
        'Dim v2 As Object() = DCall.v2.ToArray
        'Dim dbDs As DataSet = Spx_UniZero(DCall.SP, v1, v2)

        JsonOutput(str)

    End Sub

    <WebMethod()>
    Public Sub VerifyUser(U As String, P As String)

        Dim var As String() = {"U", "P"}
        Dim val As Object() = {U, P}
        Dim dbDs As DataSet = Spx_UniZero("spg_UserVerify", var, val)

        JsonOutput(dbDs)
    End Sub


    <WebMethod()>
    Public Sub GetUsers(ID As String)
        Dim idx As Integer = FInt(ID)
        Dim var As String() = {"ID"}
        Dim val As Object() = {idx}
        Dim dbDs As DataSet = Spx_UniZero("spg_GetUsers", var, val)

        JsonOutput(dbDs)

    End Sub

    Private Sub JsonOutput(dbDs As DataSet)

        Dim json As String = JsonConvert.SerializeObject(dbDs, Formatting.Indented).ToString()

        Context.Response.ContentType = "application/json"
        Context.Response.Write(json)
        Context.Response.End()

    End Sub

    Private Sub JsonOutput(txtString As String)

        Dim json As String = JsonConvert.SerializeObject(txtString, Formatting.Indented).ToString()

        Context.Response.ContentType = "application/json"
        Context.Response.Write(json)
        Context.Response.End()

    End Sub



End Class