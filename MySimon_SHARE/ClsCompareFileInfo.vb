Imports System.IO
Public Class ClsCompareFileInfo
    Implements IComparer

    ''' <summary>
    ''' Compares the specified x.	
    ''' </summary>
    ''' <param name="x">The x.</param>
    ''' <param name="y">The y.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim file1 As FileInfo
        Dim file2 As FileInfo

        file1 = DirectCast(x, FileInfo)
        file2 = DirectCast(y, FileInfo)

        Compare = DateTime.Compare(file2.CreationTime, file1.CreationTime)
    End Function
End Class
