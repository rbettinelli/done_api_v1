'-----------------------------------------------------------------
' ComLib - Communication Lib.
'-----------------------------------------------------------------
' R.Bettinelli                                               2014
'-----------------------------------------------------------------
' 06/19/2014 - Doc.
'-----------------------------------------------------------------
Imports System.Net.Mail
Imports System.Threading
Imports System.Globalization
Imports System.IO
Imports System.Net

Module ComLib

    Public Const EofyFormat As String = "12/31/"
    Public Const Dfmt As String = "MM/dd/yyyy"

#Region "Misc"

    '---------------------------------------------------------------------------
    ' AddToDropDown - Drop Down List Population.
    '---------------------------------------------------------------------------
    Public Sub AddToDropDown(drp As DropDownList, dbDs As DataSet, txtRow() As String, typ As String, Optional tableId As Integer = 0)
        drp.Items.Clear()
        Try
            drp.Items.Add(New ListItem("Select " + typ, "0"))
            For x As Integer = 0 To dbDs.Tables(tableId).Rows.Count - 1

                Dim li As ListItem = New ListItem With {
                    .Text = dbDs.Tables(tableId).Rows(x)(txtRow(0)).ToString,
                    .Value = dbDs.Tables(tableId).Rows(x)(txtRow(1)).ToString
                }
                drp.Items.Add(li)
                drp.SelectedIndex = 0
            Next
        Catch ex As Exception
            ' No Records
        End Try
    End Sub

    '---------------------------------------------------------------------------
    ' ProvideWeekDateLimit - Limits Function.
    '---------------------------------------------------------------------------
    Public Function ProvideWeekDateLimit(yr As Integer) As List(Of String)
        Dim ddl As List(Of String) = New List(Of String)
        Dim lbs As List(Of String) = New List(Of String)
        Dim lbe As List(Of String) = New List(Of String)

        Dim var() As String = {"intYr"}
        Dim val() As Object = {yr}
        Dim dbDs As DataSet = spx_Uni("spg_getYearSpan", var, val)

        Dim sDate As Date = Convert.ToDateTime(dbDs.Tables(0).Rows(0)("yearStart").ToString)
        Dim eDate As Date = Convert.ToDateTime(dbDs.Tables(0).Rows(0)("yearEnd").ToString)

        Dim noWeeksDiff As Integer = Convert.ToInt32(DateDiff("w", sDate, eDate) / 2)

        Dim dteStart As Date = ToBeginWeek(sDate)
        Dim wks As Integer = Convert.ToInt32(noWeeksDiff)
        Do While wks > 0
            Dim dteSs As String = dteStart.ToString(Dfmt)
            Dim dteSe As String = dteStart.AddDays(13).ToString(Dfmt)
            Dim ns1 As String = dteSs.Replace("-", "/")
            Dim ns2 As String = dteSe.Replace("-", "/")
            lbs.Add(ns1)
            lbe.Add(ns2)
            dteStart = dteStart.AddDays(14)
            wks -= 1
        Loop
        wks = 1
        For x = 0 To lbs.Count - 1
            Dim ss As String = lbs.Item(x).ToString
            Dim se As String = lbe.Item(x).ToString
            Dim str As String = wks.ToString("D2") & ":" & ss & "-" & se
            ddl.Add(str)
            wks += 1
        Next
        Return ddl
    End Function

    '---------------------------------------------------------------------------
    ' Ddl_List - Drop Down List Create
    '---------------------------------------------------------------------------
    Public Function Ddl_List(ddl As DropDownList, yrStr As String) As Integer
        Dim cw As Integer = 0
        ddl.Items.Clear()
        Dim lbs As List(Of String) = New List(Of String)
        Dim lbe As List(Of String) = New List(Of String)

        Dim yr As Integer = Year(Today)
        If Not String.IsNullOrEmpty(yrStr) Then
            yr = Convert.ToInt32(yrStr)
        End If

        Dim var() As String = {"intYr"}
        Dim val() As Object = {yr}
        Dim dbDs As DataSet = spx_Uni("spg_getYearSpan", var, val)

        Dim sDate As Date = Convert.ToDateTime(dbDs.Tables(0).Rows(0)("yearStart").ToString)
        Dim eDate As Date = Convert.ToDateTime(dbDs.Tables(0).Rows(0)("yearEnd").ToString)

        Dim noWeeksDiff As Integer = Convert.ToInt32(DateDiff("w", sDate, eDate) / 2)

        Dim dteStart As Date = ToBeginWeek(sDate)
        Dim wks As Integer = Convert.ToInt32(noWeeksDiff)
        Dim flg As Boolean = False
        Do While wks > 0
            Dim dteSs As String = dteStart.ToString(Dfmt)
            Dim dteSe As String = dteStart.AddDays(13).ToString(Dfmt)
            Dim ns1 As String = dteSs.Replace("-", "/")
            Dim ns2 As String = dteSe.Replace("-", "/")
            lbs.Add(ns1)
            lbe.Add(ns2)
            dteStart = dteStart.AddDays(14)
            If dteStart >= Today() And Today() <= dteStart.AddDays(13) And flg = False Then
                cw = noWeeksDiff - (wks - 1)
                flg = True
                'Dim x As Integer = 0
            End If
            wks -= 1
        Loop
        wks = 1
        For x = 0 To lbs.Count - 1
            Dim ss As String = lbs.Item(x).ToString
            Dim se As String = lbe.Item(x).ToString
            Dim str As String = wks.ToString("D2") & ":" & ss & "-" & se

            ddl.Items.Add(New ListItem(str, wks.ToString))
            ' lb.Items.Add(New ListItem(str, wks.ToString))
            wks += 1
        Next
        Return cw
    End Function


    Function DateSet(txt As String) As DateTime

        DateSet = Convert.ToDateTime("1/1/1950")
        Dim firstInputDate As String = txt
        Dim firstDateTime As DateTime

        If DateTime.TryParse(firstInputDate, firstDateTime) Then
            DateSet = firstDateTime
        End If

        Return DateSet

    End Function


    '---------------------------------------------------------------------------
    ' FindYear - Find which year a date drops in.
    '---------------------------------------------------------------------------
    Function FindYear(dte As String) As Integer
        Dim varX() As String = {"dte"}
        Dim valX() As Object = {Convert.ToDateTime(dte)}
        Dim dbDs As DataSet = spx_Uni("spg_FindYear", varX, valX)
        Dim fys As Integer = FInt(dbDs.Tables(0).Rows(0)("year").ToString)
        Return fys
    End Function

    '---------------------------------------------------------------------------
    ' PullTh - Pull Time & Half
    '---------------------------------------------------------------------------
    Function PullTh(stfId As Integer, startDate As String, endDate As String) As Double

        Dim fys As Integer = FindYear(startDate)
        Dim fye As Integer = FindYear(endDate)
        Dim wks As List(Of String) = ProvideWeekDateLimit(fys)
        Dim wke As List(Of String) = ProvideWeekDateLimit(fye)

        Dim fwS As Integer = 0
        Dim fwE As Integer = 0

        Dim ps As DateTime = Convert.ToDateTime(startDate)
        Dim pe As DateTime = Convert.ToDateTime(endDate)
        Dim th As Double = 0.0

        If fys = fye Then
            For x As Integer = 0 To wks.Count - 1
                Dim wkline() As String = Split(wks(x), ":")
                Dim wkdetail() As String = Split(wkline(1), "-")
                Dim ds As DateTime = Convert.ToDateTime(wkdetail(0))
                Dim de As DateTime = Convert.ToDateTime(wkdetail(1))
                If ps >= ds And ps <= de And fwS = 0 Then
                    fwS = x
                End If
                If pe >= ds And pe <= de And fwE = 0 Then
                    fwE = x
                End If
            Next

            For x As Integer = fwS To fwE

                Dim wkline() As String = Split(wks(x), ":")
                Dim wkdetail() As String = Split(wkline(1), "-")

                Dim otwk1 As Double = OverTimeH(stfId, wkdetail(0))
                Dim we As String = Convert.ToDateTime(wkdetail(0)).AddDays(7).ToString("MM/dd/yyyy")
                Dim otwk2 As Double = OverTimeH(stfId, we)

                th += otwk1 + otwk2
            Next

        Else
            For x As Integer = 0 To wks.Count - 1
                Dim wkline() As String = Split(wks(x), ":")
                Dim wkdetail() As String = Split(wkline(1), "-")
                Dim ds As DateTime = Convert.ToDateTime(wkdetail(0))
                Dim de As DateTime = Convert.ToDateTime(wkdetail(1))
                If ps >= ds And ps <= de And fwS = 0 Then
                    fwS = x
                End If
            Next
            For x As Integer = 0 To wke.Count - 1
                Dim wkline() As String = Split(wke(x), ":")
                Dim wkdetail() As String = Split(wkline(1), "-")
                Dim ds As DateTime = Convert.ToDateTime(wkdetail(0))
                Dim de As DateTime = Convert.ToDateTime(wkdetail(1))
                If pe >= ds And pe <= de And fwE = 0 Then
                    fwE = x
                End If
            Next

            For x As Integer = fwS To wks.Count - 1

                Dim wkline() As String = Split(wks(x), ":")
                Dim wkdetail() As String = Split(wkline(1), "-")

                Dim otwk1 As Double = OverTimeH(stfId, wkdetail(0))
                Dim we As String = Convert.ToDateTime(wkdetail(0)).AddDays(7).ToString("MM/dd/yyyy")
                Dim otwk2 As Double = OverTimeH(stfId, we)

                th += otwk1 + otwk2
            Next

            For x As Integer = 0 To fwE

                Dim wkline() As String = Split(wke(x), ":")
                Dim wkdetail() As String = Split(wkline(1), "-")

                Dim otwk1 As Double = OverTimeH(stfId, wkdetail(0))
                Dim we As String = Convert.ToDateTime(wkdetail(0)).AddDays(7).ToString("MM/dd/yyyy")
                Dim otwk2 As Double = OverTimeH(stfId, we)

                th += otwk1 + otwk2
            Next


        End If

        Return th

    End Function

    '---------------------------------------------------------------------------
    ' OverTimeH - OvertTime Hours. 
    '---------------------------------------------------------------------------
    Function OverTimeH(stfId As Integer, wk As String) As Double

        Dim we As String = Convert.ToDateTime(wk).AddDays(6).ToString("MM/dd/yyyy")

        Dim varX() As String = {"stf_ID", "payD_Date_S1", "payD_Date_E1"}
        Dim valX() As Object = {stfId, wk, we}
        Dim dbDs As DataSet = spx_Uni("spg_PullWeek", varX, valX)

        Dim totbank As Double = 0.0
        For y As Integer = 0 To dbDs.Tables(0).Rows.Count - 1
            Dim dbRow As DataRow = dbDs.Tables(0).Rows(y)
            totbank += FDbl(dbRow("payD_RegHours").ToString) + FDbl(dbRow("payD_Bankarned").ToString)
        Next

        'Dim ot As Double
        Dim oh As Double
        Dim dh As Double
        Dim ah As Double

        If totbank > 44 Then
            'ot = totbank - 35
            oh = totbank - 44
            ah = oh * 1.5
            dh = ah - oh
        End If

        Return dh

    End Function

    '---------------------------------------------------------------------------
    ' GetSoureCodeFromFile - Pull aspx and convert to String for Literal.
    '---------------------------------------------------------------------------
    Function GetSoureCodeFromFile(url As String) As String
        Dim r As String
        Dim wc As WebClient = New WebClient
        r = wc.DownloadString(url)
        wc.Dispose()
        Return r
    End Function

    '---------------------------------------------------------------------------
    ' ChkNull - Check Null - Evaluate Missing Info Form Object
    '---------------------------------------------------------------------------
    Function ChkNull(inVar() As String, inVal() As Object) As String
        ChkNull = ""
        For x As Integer = 0 To inVal.Count - 1
            Dim iv As TextBox = CType(inVal(x), TextBox)
            If String.IsNullOrEmpty(iv.Text) Then
                ChkNull += inVar(x) + " Missing,"
            End If
        Next
        If ChkNull.Length > 1 Then
            ChkNull = Left(ChkNull, ChkNull.Length - 1) + "."
        End If
    End Function

    '---------------------------------------------------------------------------
    ' SaveMe - Global Save Function.  - One Save to Rule them all. 
    '---------------------------------------------------------------------------
    Public Function SaveMe(txtError As Label, tstVar() As String, tstVal() As Object, savVar() As String, savVal() As Object, storedProc As String, Optional err1 As String = "") As String
        Dim err As String = ChkNull(tstVar, tstVal)
        err += err1
        txtError.Text = ""
        Dim retn As String = "0"
        Dim dbDs As DataSet
        If String.IsNullOrEmpty(err) Then
            dbDs = spx_Uni(storedProc, savVar, savVal)
            txtError.Text = "Saved. " + Now.ToString()
            If dbDs.Tables.Count > 0 Then
                retn = dbDs.Tables(0).Rows(0)("ID").ToString
            End If
        Else
            txtError.Text = err + "Cant Save.Please Check."
        End If
        Return retn
    End Function

    '---------------------------------------------------------------------------
    ' GetDirCont
    '---------------------------------------------------------------------------
    Public Function GetDirCont(ByVal fullPath As String, ByVal extFilter As String, ByVal pullNo As Integer) As FileInfo()
        Dim dDir As New DirectoryInfo(fullPath)
        Dim files() As FileInfo = dDir.GetFiles(extFilter)
        Array.Sort(files, New ClsCompareFileInfo)
        If pullNo > 0 Then
            ReDim Preserve files(4)
        End If
        GetDirCont = files
    End Function

#End Region

#Region "Dates"
    '---------------------------------------------------------------------------
    ' ToBeginWeek - Calc To Beginging Of Week
    '---------------------------------------------------------------------------
    Public Function ToBeginWeek(dte As Date) As Date
        Dim tow As DayOfWeek = dte.DayOfWeek
        Dim toSun = tow - DayOfWeek.Sunday
        Return dte.AddDays(-toSun)
    End Function

    '---------------------------------------------------------------------------
    ' Wofy - Week of Year
    '---------------------------------------------------------------------------
    Public Function Wofy(dte As Date, Optional half As Boolean = False) As Integer
        Dim woy As Integer = Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dte, CalendarWeekRule.FirstDay, DayOfWeek.Sunday)
        Dim nw As Double = woy
        If half Then
            nw = (woy / 2)
            'woy = (woy / 2)
        End If
        Dim wkofYear As Integer = Convert.ToInt32(Math.Round(nw, MidpointRounding.AwayFromZero).ToString)
        Return wkofYear
    End Function

    '---------------------------------------------------------------------------
    ' GetNthDayOfNthWeek - Figure out where dates drop
    '---------------------------------------------------------------------------
    Private Function GetNthDayOfNthWeek(ByVal dt As Date, ByVal dayofWeek As Integer, ByVal whichWeek As Integer) As String
        'specify which day of which week of a month and this function will get the date
        'this function uses the month and year of the date provided

        'get first day of the given date
        Dim dtFirst As Date = DateSerial(dt.Year, dt.Month, 1)

        'get first DayOfWeek of the month
        Dim dtRet As Date = dtFirst.AddDays(6 - dtFirst.AddDays(-(dayofWeek + 1)).DayOfWeek)

        'get which week
        dtRet = dtRet.AddDays((whichWeek - 1) * 7)

        'if day is past end of month then adjust backwards a week
        If dtRet >= dtFirst.AddMonths(1) Then
            dtRet = dtRet.AddDays(-7)
        End If

        'return
        Return dtRet.ToString

    End Function

    '---------------------------------------------------------------------------
    ' EasterDate2 - Calulate What day Easter Drops On. 
    '---------------------------------------------------------------------------
    Public Function EasterDate2(yr As Integer) As Date
        Dim d As Integer
        d = (((255 - 11 * (yr Mod 19)) - 21) Mod 30) + 21
        EasterDate2 = DateSerial(yr, 3, 1).AddDays(d + Convert.ToInt32(d > 48) + 6 - ((yr + yr \ 4 + d + Convert.ToInt32(d > 48) + 1) Mod 7))
    End Function

    '---------------------------------------------------------------------------
    ' CalcHolidayDate - Calc Holiday Days Drop On.. 
    '---------------------------------------------------------------------------
    Public Function CalcHolidayDate(dy As DateTime) As String

        Dim str As String = ""

        Dim chkYr As Integer = Year(dy)
        Dim chkMn As Integer = Month(dy)
        Dim chkDy As Integer = Day(dy)

        Dim varX() As String = {"mn"}
        Dim valX() As Object = {chkMn}
        Dim dbDs As DataSet = spx_Uni("spg_pullHolidayList", varX, valX)

        Dim lbdates As ListBox = New ListBox

        If dbDs.Tables(0).Rows.Count > 0 Then

            For x = 0 To dbDs.Tables(0).Rows.Count - 1

                Dim ehd As String = dbDs.Tables(0).Rows(x)("Holiday").ToString.Trim
                Dim eMn As Integer = FInt(dbDs.Tables(0).Rows(x)("Mn").ToString)
                Dim eDy As Integer = FInt(dbDs.Tables(0).Rows(x)("Dy").ToString)
                Dim eMnInc As String = dbDs.Tables(0).Rows(x)("MnInc").ToString.Trim

                If eMnInc = "----" Then
                    ' Exact Match
                    If chkDy = eDy Then
                        lbdates.Items.Add(ehd + "|" + eMn.ToString + "/" + eDy.ToString + "/" + chkYr.ToString)
                    End If
                End If
                If Mid(eMnInc, 2, 1) = "M" Then
                    ' Xth Monday
                    Dim mNp As Integer = FInt(Left(eMnInc, 1))

                    lbdates.Items.Add(ehd + "|" + GetNthDayOfNthWeek(DateSerial(chkYr, chkMn, 1), DayOfWeek.Monday, mNp))

                End If
                If Mid(eMnInc, 3, 1) = "E" Then
                    Dim easter As DateTime = EasterDate2(chkYr)
                    Dim eM As DateTime = easter.AddDays(1)
                    Dim eF As DateTime = easter.AddDays(-2)
                    lbdates.Items.Add("Good Friday|" + eF.ToString)
                    lbdates.Items.Add("Easter|" + easter.ToString)
                    lbdates.Items.Add("Easter Monday|" + eM.ToString)
                End If
                If Left(eMnInc, 2) = "MB" Then
                    Dim fm As Integer = Right(eMnInc, 2)
                    For y = (fm - 1) To (fm - 7) Step -1
                        Dim dt As DateTime = Convert.ToDateTime(chkMn.ToString + "/" + y.ToString + "/" + chkYr.ToString)
                        If dt.DayOfWeek = DayOfWeek.Monday Then
                            lbdates.Items.Add(ehd + "|" + chkMn.ToString + "/" + y.ToString + "/" + chkYr.ToString)
                        End If
                    Next
                End If

            Next

            For x = 0 To lbdates.Items.Count - 1
                Dim sp() As String = Split(lbdates.Items(x).ToString(), "|")
                Dim ckd As DateTime = Convert.ToDateTime(sp(1))
                If ckd = dy Then
                    str = sp(0)
                End If
            Next

        End If
        lbdates.Dispose()

        Return str

    End Function

    '---------------------------------------------------------------------------
    ' ConvDates - Convert Dates to Useable Form  
    '---------------------------------------------------------------------------
    Public Function ConvDates(dtes() As Object) As DateTime()
        Dim conD(dtes.Length - 1) As DateTime
        For x As Integer = 0 To conD.Length - 1
            conD(x) = Nothing
            Dim txtb As TextBox = CType(dtes(x), TextBox)
            If Not String.IsNullOrEmpty(txtb.Text) Then
                conD(x) = Convert.ToDateTime(txtb.Text)
            End If
        Next
        Return conD
    End Function

#End Region

#Region "Conversion"
    '---------------------------------------------------------------------------
    ' IsInputNumeric - Confim if Entry is Number.
    '---------------------------------------------------------------------------
    Function IsInputNumeric(input As String) As Boolean
        If String.IsNullOrWhiteSpace(input) Then Return False
        If IsNumeric(input) Then Return True
        Dim parts() As String = input.Split("/"c)
        If parts.Length <> 2 Then Return False
        Return IsNumeric(parts(0)) AndAlso IsNumeric(parts(1))
    End Function

    '---------------------------------------------------------------------------
    ' FDbl - Return Double from String
    '---------------------------------------------------------------------------
    Public Function FDbl(txt As String) As Double
        Dim rt As Double = 0.0
        If Not String.IsNullOrEmpty(txt) Then
            txt = CleanString(txt.Trim)
            Try
                rt = Convert.ToDouble(txt)
            Catch ex As Exception
                ' cant do it. 
            End Try
        End If
        Return rt
    End Function

    '---------------------------------------------------------------------------
    ' FBool - Return Boolean from String
    '---------------------------------------------------------------------------
    Public Function FBool(txt As String) As Boolean

        Dim rt As Boolean = False
        If Not String.IsNullOrEmpty(txt) Then
            txt = CleanString(txt.Trim)
            Try
                rt = Convert.ToBoolean(txt)
            Catch ex As Exception
                ' cant.
            End Try
        End If
        Return rt
    End Function

    '---------------------------------------------------------------------------
    ' FInt - Return Integer from String
    '---------------------------------------------------------------------------
    Public Function FInt(txt As String) As Integer
        Dim rt As Integer = 0
        If IsInputNumeric(txt) Then
            txt = CleanString(txt.Trim)
            Try
                rt = Convert.ToInt32(txt)
            Catch ex As Exception
                ' cant
            End Try
        End If
        Return rt
    End Function

    '---------------------------------------------------------------------------
    ' CleanString - Strip unwanted Chars
    '---------------------------------------------------------------------------
    Public Function CleanString(txt As String) As String
        If txt IsNot Nothing Then
            Dim strpChar() = {",", "$", "/", "\", "{", "}", " "}
            ' ReSharper disable once LoopCanBeConvertedToQuery
            For Each str As String In strpChar
                txt = txt.Replace(str, "")
            Next
        End If
        Return txt
    End Function

    '---------------------------------------------------------------------------
    ' FSng - Return String To String
    '---------------------------------------------------------------------------
    Public Function FSng(txt As String) As Single

        Dim rt As Single = 0.0
        If Not String.IsNullOrEmpty(txt) And IsInputNumeric(txt) Then
            txt = CleanString(txt.Trim)
            Try
                rt = Convert.ToSingle(txt)
            Catch ex As Exception
                'cant
            End Try
        End If
        Return rt
    End Function

    '---------------------------------------------------------------------------
    ' Rounder - Round to neares (roundTo = .25 Default)
    '---------------------------------------------------------------------------
    Function Rounder(ByVal originalNumber As Double, Optional ByVal roundTo As Double = 0.25) As Double
        Rounder = originalNumber
        If originalNumber > 0.0 Then
            ' ReSharper disable once CompareOfFloatsByEqualityOperator
            If (originalNumber / roundTo) * 2 = CInt((originalNumber / roundTo) * 2) Then
                originalNumber += (originalNumber / Math.Abs(originalNumber)) * roundTo / 10
            End If
            Rounder = Math.Round(originalNumber / roundTo, 0) * roundTo
        End If
    End Function

    '---------------------------------------------------------------------------
    ' CdVal - String to TimeDate
    '---------------------------------------------------------------------------
    Public Function CdVal(txt As String) As DateTime
        CdVal = Convert.ToDateTime(txt)
        Return CdVal
    End Function


    '---------------------------------------------------------------------------
    ' RandGen - Rand Generator
    '---------------------------------------------------------------------------
    Public Function RandGen(indx As Integer) As String
        Randomize()
        Dim s As String = ""
        Const str As String = "123456789ABCDEFGHIJKLNMNPQRSTUVWXYZabcdefghijklmnpqrstuvwxyz"
        For x = 1 To indx
            Dim r As Integer = Int(Rnd() * str.Length) + 1
            s += Mid(str, r, 1)
        Next
        Return s
    End Function

#End Region


End Module
