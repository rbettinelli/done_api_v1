Imports System.ComponentModel
Imports System.Web.Script.Services
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports Newtonsoft.Json

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
<System.Web.Script.Services.ScriptService()>
<System.Web.Services.WebService(Namespace:="http://doneapi.io-serv.com/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.None)>
<ToolboxItem(False)>
Public Class MyWebService
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function

    <WebMethod()>
    Public Sub GetUsers(ID As String)
        Dim idx As Integer = FInt(ID)
        Dim var() As String = {"ID"}
        Dim val() As Object = {idx}
        Dim dbDs As DataSet = Spx_UniZero("spg_GetUsers", var, val)

        Dim json As String = JsonConvert.SerializeObject(dbDs, Formatting.Indented).ToString()

        Context.Response.ContentType = "application/json"
        Context.Response.Write(json)
        Context.Response.End()

    End Sub






End Class