Imports System.Runtime.CompilerServices
Public Module Array
    <Extension()>
    Public Sub Append(Of T)(ByRef TArray As T(), value As T)
        ReDim Preserve TArray(UBound(TArray) + 1)
        TArray(UBound(TArray)) = value
    End Sub
    <Extension()>
    Public Sub Append(Of T)(ByRef Tarray As T(), value As T())
        For Each I In value
            Tarray.Append(value)
        Next
    End Sub
    <Extension()>
    Public Sub Delete(Of T)(ByRef TArray As T(), Position As Integer)
        Dim Tmp As Object()
        ReDim Tmp(0)
        For i As Integer = 0 To UBound(TArray) - 1
            If Position <> i Then
                Tmp(UBound(Tmp)) = TArray(i)
                If i < UBound(Tmp) - 1 Then
                    ReDim Preserve Tmp(UBound(Tmp) + 1)
                End If
            End If
        Next
        TArray = Tmp
    End Sub
    <Extension()>
    Public Sub SetLast(Of T)(ByRef TArray As T(), value As T)
        TArray(UBound(TArray)) = value
    End Sub
    <Extension()>
    Public Function Collect(ByRef TArray As String(), value As String) As String
        Dim Tmp As String = ""
        For i As Integer = 0 To UBound(TArray) - 1
            Tmp += TArray(i) + value
        Next
        Return Tmp + TArray.Last
    End Function
End Module
