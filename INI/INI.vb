
''' <summary>
''' 提供INI Class错误
''' </summary>
Public NotInheritable Class INIException
    Inherits System.Exception
    Dim Error_Code As Byte = 0 '错误代码
    Dim Error_Position(1) As UShort '错误出现位置(用坐标表示)
    Private Sub Initialize()
        Error_Position(0) = 0
        Error_Position(1) = 0
    End Sub
    ''' <summary>
    ''' 用指定的错误号查询错误的信息。
    ''' </summary>
    ''' <param name="ErrorCode">错误号</param>
    ''' <returns>错误描述</returns>
    Public Shared Function GetErrorExplain(ErrorCode As Byte) As String
        Select Case ErrorCode
            Case 0
                Return "No Error"
            Case 1
                Return "No Text To Load"
            Case 2
                Return "Can Not Find The File:%FileName%"
            Case 3
                Return "Can Not Use The Value"
            Case 4
                Return "Not Found '[]'"
            Case 5
                Return "Not Able To Set Empty"
            Case 6
                Return "Found The Same Name"
            Case 7
                Return "Not Found The Name"
            Case Else
                Throw New Exception("Unknown Error Code!")
        End Select
    End Function
    Sub New()
        Me.Initialize()
    End Sub
    Sub New(ErrorCode As Byte)
        Me.Error_Code = ErrorCode
        Initialize()
    End Sub
    Sub New(ErrorPosition As UShort())
        If UBound(ErrorPosition) = 1 Then
            Me.Error_Position = ErrorPosition
        Else
            Throw New Exception("Arrmy is not fit with form")
        End If
    End Sub
    Sub New(ErrorCode As Byte, ErrorPosition As UShort())
        Me.Error_Code = ErrorCode
        If UBound(ErrorPosition) = 1 Then
            Me.Error_Position = ErrorPosition
        Else
            Throw New Exception("Arrmy is not fit with form")
        End If
    End Sub
    ''' <summary>
    ''' 返回出现的错误号
    ''' </summary>
    ''' <returns>出现的错误号</returns>
    Public ReadOnly Property ErrorCode As Byte
        Get
            Return Me.Error_Code
        End Get
    End Property
    ''' <summary>
    ''' 返回出现的错误的位置
    ''' </summary>
    ''' <returns>出现错误的位置</returns>
    Public ReadOnly Property ErrorPosition As UShort()
        Get
            Return Me.Error_Position
        End Get
    End Property
    ''' <summary>
    ''' 泛海错误的解释
    ''' </summary>
    ''' <returns>错误的解释</returns>
    Public ReadOnly Property ErrorExplain As String
        Get
            Return INIException.GetErrorExplain(Me.Error_Code)
        End Get
    End Property
End Class
Public NotInheritable Class INI
    Dim _INIFathers As INIFather() '配置
    Public Shared Function IsAnnotation(Text As String) As Boolean
        Dim Tmp As Integer = InStr(Text, ";") - 1
        If Tmp < 0 Then Return False
        Return Not IsMeaingFul(Left(Text, Tmp))
    End Function
    Public Shared Function IsMeaingFul(Text As String) As Boolean
        Dim I As Byte() = New System.Text.ASCIIEncoding().GetBytes(Text)
        For Each value As Byte In I
            If value <> 8 And value <> 13 And value <> 10 Then Return True
        Next
        Return False
    End Function
    Public Shared Function LeaveAnnotation(Text As String) As String
        Dim Tmp As String() = Split(Text, vbCrLf), TmpValue As String = ""
        For i As Integer = 0 To UBound(Tmp)
            Tmp(i) = LeaveLineAnnotation(Tmp(i))
            TmpValue += Tmp(i) + vbCrLf
        Next
        Return Left(TmpValue, TmpValue.Length - 2)
    End Function
    Public Shared Function LeaveLineAnnotation(Text As String) As String
        Dim i As Integer = InStr(Text, ";")
        If i > 0 Then
            Return Left(Text, i - 1)
        Else
            Return Text
        End If
    End Function
    ''' <summary>
    ''' 加载INI文件
    ''' </summary>
    ''' <param name="INIText">INI文件内容</param>
    Sub New(INIText As String)
        Load(INIText)
    End Sub
    ''' <summary>
    ''' 建立一个空INI类
    ''' </summary>
    Sub New()
    End Sub
    ''' <summary>
    ''' 加载INI文件
    ''' </summary>
    ''' <param name="INIFile">文件读取</param>
    Sub New(INIFile As IO.StreamReader)
        Load(INIFile.ReadToEnd)
    End Sub
    ''' <summary>
    ''' 加载INI文本
    ''' </summary>
    ''' <param name="Text"></param>
    Public Sub Load(Text As String)
        Dim Tmp As String() = Split(Text, vbCrLf),'零时行储存
            TmpINIFathers As INIFather(),
            TmpText As String()
        ReDim TmpText(0)
        '对文本行进行分段
        For TmpLine As UShort = 0 To UBound(Tmp)
            Dim TText As String = Tmp(TmpLine)
            If Left(TText, 1) = "[" And Right(LeaveAnnotation(TText), 1) = "]" Then
                ReDim Preserve TmpText(UBound(TmpText) + 1)
                TmpText.SetLast(TText)
            Else
                TmpText(UBound(TmpText)) += vbCrLf + TText
            End If
        Next
        '如果没有注释
        If TmpText(0) = Nothing Then
            ReDim TmpINIFathers(0 To UBound(TmpText) - 1)
            For i As UShort = 1 To UBound(TmpText)
                TmpINIFathers(i - 1) = New INIFather(TmpText(i))
            Next
        Else
            ReDim TmpINIFathers(0 To UBound(TmpText))
            For i As UShort = 0 To UBound(TmpText)
                TmpINIFathers(i) = New INIFather(TmpText(i))
            Next
        End If
        _INIFathers = TmpINIFathers
    End Sub
    Public Sub LoadFile(File As String)
        Load(IO.File.ReadAllText(File))
    End Sub
    Public Sub Add(value As String)
        Dim Tmp As String() = Split(value, vbCrLf),'零时行储存
            TmpINIFathers As INIFather(),
            TmpText As String()
        ReDim TmpText(0)
        '对文本行进行分段
        For TmpLine As UShort = 0 To UBound(Tmp)
            Dim TText As String = Tmp(TmpLine)
            If Left(TText, 1) = "[" And Right(TText, 1) = "]" Then
                ReDim Preserve TmpText(UBound(TmpText) + 1)
                TmpText.SetLast(TText)
            Else
                TmpText(UBound(TmpText)) += vbCrLf + TText
            End If
        Next
        '如果没有注释
        If TmpText(0) = Nothing Then
            ReDim TmpINIFathers(0 To UBound(TmpText) - 1)
            For i As UShort = 1 To UBound(TmpText)
                TmpINIFathers(i - 1).Load(TmpText(i))
            Next
        Else
            ReDim TmpINIFathers(0 To UBound(TmpText))
            For i As UShort = 0 To UBound(TmpText)
                TmpINIFathers(i).Load(TmpText(i))
            Next
        End If
        _INIFathers.Append(TmpINIFathers)
    End Sub
    Public Property INIFathers(Number As Integer) As INIFather
        Get
            Return _INIFathers(Number)
        End Get

        Set(value As INIFather)
            _INIFathers(Number) = value
        End Set
    End Property
    Public Property INIFathers(Name As String) As INIFather
        Get
            For Each I In _INIFathers
                If UCase(I.Name) = UCase(Name) Then Return I
            Next
            Return Nothing
        End Get
        Set(value As INIFather)
            For i As Byte = 0 To UBound(_INIFathers)
                If UCase(_INIFathers(i).Name) = UCase(Name) Then
                    _INIFathers(i) = value
                    Exit Property
                End If
            Next
            Throw New Exception("没有找到对应的名称")
        End Set
    End Property

    Public Sub Save(File As String)
        Dim tmp As String() = {}
        For I As Integer = 0 To UBound(_INIFathers)
            tmp.Append(_INIFathers(I))
        Next
        Debug.Print(Collect(tmp, vbCrLf))
        IO.File.WriteAllText(File, Collect(tmp, vbCrLf))
    End Sub
End Class
Public Class INIFather
    Dim _Child As INIChild()
    Public Property Name As String
    Public Property Child(Number As Integer) As INIChild
        Get
            Return _Child(Number)
        End Get
        Set(value As INIChild)
            Dim TmpChild As INIChild() = _Child
            TmpChild(Number) = value
            If CheckForSame(TmpChild) Then
                Throw New INIException(6)
            Else
                _Child = TmpChild
            End If
        End Set
    End Property
    Public Property Child(Name As String) As INIChild
        Get
            For Each I As INIChild In _Child
                If UCase(I.Name) = UCase(Name) Then Return I
            Next
            Return Nothing
        End Get
        Set(value As INIChild)
            Dim TmpChild As INIChild() = _Child
            For I As Integer = 0 To UBound(TmpChild)
                If UCase(TmpChild(I).Name) = UCase(Name) Then
                    TmpChild(I) = value
                    If CheckForSame(TmpChild) Then
                        Throw New INIException(6)
                    Else
                        _Child = TmpChild
                    End If
                End If
            Next
            Throw New INIException(7)
        End Set
    End Property
    Public ReadOnly Property ChildNumber As Integer
        Get
            Return UBound(_Child) + 1
        End Get
    End Property
    Sub Append(INIChild As INIChild)
        If IsNothing(_Child) Then
            _Child = New INIChild() {INIChild}
            Exit Sub
        End If
        If CheckForSame(_Child, INIChild) Then
            Throw New Exception("有相同的Child标签")
        Else
            _Child.Append(INIChild)
        End If
    End Sub
    Sub New(Name As String, Child As INIChild())
        Me.Name = Name
        If CheckForSame(Child) Then '
            Throw New Exception("找到相同的Child名称")
        Else
            _Child = Child
        End If
    End Sub
    Sub New(Text As String)
        Load(Text)
    End Sub
    Sub New(Name As String, ChildText As String)
        Me.Name = Name
        Load(ChildText)
    End Sub
    Public Sub Load(Text As String)
        Dim TmpText As String() = Split(Text, vbCrLf), TmpName As String = INI.LeaveAnnotation(TmpText(0))
        '如果有名字
        If Left(TmpText(0), 1) = "[" AndAlso Right(TmpName, 1) = "]" Then
            Name = Mid(TmpName, 2, TmpName.Length - 2)
        End If
        If UBound(TmpText) = 0 Then Exit Sub
        Dim Tmp As String() = Split(Right(Text, Len(Text) - InStr(Text, vbCrLf) - 1), vbCrLf)
        Dim TmpINI As INIChild()
        ReDim TmpINI(0 To UBound(Tmp))
        For i As Byte = 0 To UBound(Tmp)
            TmpINI(i) = Tmp(i)
        Next
        If CheckForSame(TmpINI) Then Throw New Exception("有相同的Child标签")
        _Child = TmpINI
    End Sub
    Public Sub AddChild(Text As String)
        For Each i In Split(Text, vbCrLf)
            Append(i)
        Next
    End Sub
    Function CheckForSame(INIChilds As INIChild()) As Boolean
        Dim TmpName As String() = {""}
        For i As Byte = 0 To UBound(INIChilds)
            TmpName.SetLast(INIChilds(i))
            ReDim Preserve INIChilds(UBound(INIChilds) + 1)
        Next
        ReDim Preserve INIChilds(UBound(INIChilds) - 1)
        For I1 As Byte = 0 To UBound(TmpName)
            For I2 As Byte = I1 + 1 To UBound(TmpName)
                If TmpName(I1) = TmpName(I2) Then Return True
            Next
        Next
        Return False
    End Function
    Function CheckForSame(INIChilds As INIChild(), OtherChild As INIChild()) As Boolean
        For Each I In OtherChild
            If CheckForSame(INIChilds, I) Then Return True
        Next
        Return False
    End Function
    Function CheckForSame(INIChilds As INIChild(), OtherChild As INIChild) As Boolean
        For Each I In INIChilds
            If I.Name = OtherChild.Name Then Return True
        Next
        Return False
    End Function

    Public Shared Narrowing Operator CType(ByVal INIFather As INIFather) As String
        Dim tmpText As String() = {}
        If INIFather.Name = Nothing Then
            For I As UShort = 0 To UBound(INIFather._Child)
                tmpText.Append(INIFather._Child(I))
            Next
        Else
            tmpText.Append("[" & INIFather.Name & "]")
            For I As UShort = 0 To INIFather.ChildNumber - 1
                tmpText.Append(INIFather.Child(I))
            Next
        End If
        Return Collect(tmpText, vbCrLf)
    End Operator
End Class
Public Class INIChild
    Dim _Text As String, TType As Byte
    Public Property Name As String
        Get
            If TType = 0 Then Return Left(_Text, InStr(_Text, "=") - 1) Else Return Nothing
        End Get
        Set(value As String)
            If TType = 0 Then
                If InStr(value, "=", CompareMethod.Text) >= 1 Then
                    Throw New INIException(3)
                End If
                _Text = value + "=" + Right(_Text, _Text.Length - InStr(_Text, "="))
            Else
                Throw New INIException(5)
            End If
        End Set
    End Property
    Public Property Value As String
        Get
            If TType = 0 Then
                Dim i As Integer = InStr(_Text, ";"), Tmp As String
                If i > 0 Then
                    Tmp = Left(_Text, i - 1)
                Else
                    Tmp = _Text
                End If
                Return Right(Tmp, Len(Tmp) - InStr(Tmp, "=", CompareMethod.Text))
            Else
                Return Nothing
            End If
        End Get
        Set(value As String)
            If TType = 0 Then
                Dim i As Integer = InStr(_Text, ";")
                If i > 0 Then
                    _Text = Me.Name + "=" + value + Right(_Text, _Text.Length - i)
                Else
                    _Text = Name + "=" + value
                End If
            End If
        End Set
    End Property
    Sub New(Text As String)
        Me.Text = Text
    End Sub
    Public Shared Narrowing Operator CType(Text As String) As INIChild
        Return New INIChild(Text)
    End Operator
    Public Shared Narrowing Operator CType(Child As INIChild) As String
        Return Child.Text
    End Operator
    Public Property Text As String
        Get
            Return _Text
        End Get
        Set(value As String)
            If InStr(value, "=", CompareMethod.Text) >= 1 Then
                TType = 0
            ElseIf INI.IsMeaingFul(value) = False Then
                TType = 1
            ElseIf INI.IsAnnotation(value) Then
                TType = 2
            Else
                Throw New Exception("INI中没有=")
            End If
            Me._Text = value
        End Set
    End Property
End Class