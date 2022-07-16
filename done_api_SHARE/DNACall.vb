Public Class DNACall

    Public str As String

    Public SP As String

    Public v1 As List(Of String)
    Public v2 As List(Of Object)

    Public dict As Dictionary(Of String, Object)

    'binary values To nucleotide sequence
    ' remember:
    ' 00 = "A" (adenine)
    ' 01 = "G" (guanine)
    ' 10 = "C" (cytosine)
    ' 11 = "T" (thymine)

    Private ReadOnly dna_code As New Dictionary(Of String, String) From {
            {"00", "A"}, {"01", "G"}, {"10", "C"}, {"11", "T"}}

    Public Function DecodeTxt(dna As String) As String

        Dim bs As New StringBuilder
        For Each nt As String In dna
            For Each pair As KeyValuePair(Of String, String) In dna_code
                If pair.Value = nt Then
                    nt = pair.Key
                End If
            Next
            bs.Append(nt)
        Next

        Dim result As New StringBuilder
        For i As Integer = 0 To bs.Length - 1 Step 8
            Dim next_char As String = bs.ToString.Substring(i, 8)
            Dim ascii As Long = BinaryToLong(next_char)
            result.Append(Chr(ascii))
        Next i

        Return result.ToString

    End Function


    Public Sub Decode(dna As String)

        Dim result As String = DecodeTxt(dna)

        Dim StrtoSplit As String() = result.Split("~")
        If StrtoSplit.Length - 1 > 1 Then
            SP = StrtoSplit(0)
            For x As Integer = 1 To StrtoSplit.Length - 1
                Dim newsplit As String() = StrtoSplit(x).Split("~")
                v1.Add(newsplit(0))
                v2.Add(newsplit(1))
            Next
        End If

    End Sub

    Public Function Encode(txtStr As String) As String
        Dim bin As String = AsciiToBinary(txtStr)
        Dim dnaOut As New StringBuilder
        For x As Integer = 1 To bin.Length - 1 Step 2
            Dim nt As String = Mid(bin, x, 2)
            Dim r As String = dna_code.Item(nt)
            dnaOut.Append(r)
        Next
        Return dnaOut.ToString

    End Function

#Region "Convert"
    Private Function AsciiToBinary(txtStr As String) As String
        Dim result As New StringBuilder
        For i As Integer = 0 To txtStr.Length - 1
            Dim bin As String = LongToBinary(Asc(txtStr.Substring(i, 1)))
            result.Append(bin.Substring(bin.Length - 8))
        Next i
        Return result.ToString
    End Function
    Private Function LongToBinary(ByVal long_value As Long, Optional ByVal separate_bytes As Boolean = True) As String
        Dim hex_string As String = long_value.ToString("X")
        hex_string = hex_string.PadLeft(16, "0")
        Dim result_string As New StringBuilder
        For digit_num As Integer = 0 To 15
            ' Convert this hexadecimal digit into a
            ' binary nibble.
            Dim digit_value As Integer = Integer.Parse(hex_string.Substring(digit_num, 1), Globalization.NumberStyles.HexNumber)

            ' Convert the value into bits.
            Dim factor As Integer = 8
            Dim nibble_string As New StringBuilder
            For bit As Integer = 0 To 3
                If digit_value And factor Then
                    nibble_string.Append("1")
                Else
                    nibble_string.Append("0")
                End If
                factor \= 2
            Next bit

            ' Add the nibble's string to the left of the
            ' result string.
            result_string.Append(nibble_string.ToString)
        Next digit_num

        ' Add spaces between bytes if desired.
        If separate_bytes Then
            Dim tmp As New StringBuilder
            For i As Integer = 0 To result_string.ToString.Length - 8 Step 8
                tmp.Append(result_string.ToString.Substring(i, 8) & " ")
            Next i
            result_string.Append(tmp.ToString.Substring(0, tmp.Length - 1))
        End If

        ' Return the result.
        Return result_string.ToString
    End Function
    Private Function BinaryToLong(ByVal binary_value As String) As Long
        binary_value = binary_value.Trim().ToUpper()
        If binary_value.StartsWith("&B") Then binary_value = binary_value.Substring(2)

        binary_value = binary_value.Replace(" ", "")
        binary_value = binary_value.PadLeft(64, "0")

        Dim hex_result As New StringBuilder
        For nibble_num As Integer = 0 To 15
            Dim factor As Integer = 1
            Dim nibble_value As Integer = 0
            For bit As Integer = 3 To 0 Step -1
                If binary_value.Substring(nibble_num * 4 + bit, 1).Equals("1") Then
                    nibble_value += factor
                End If
                factor *= 2
            Next bit
            hex_result.Append(nibble_value.ToString("X"))
        Next nibble_num

        Return Long.Parse(hex_result.ToString, Globalization.NumberStyles.HexNumber)
    End Function

#End Region

End Class
