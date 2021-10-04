'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6
'* Create Time: 11/3/2021
'* 1.0.2	6/4/2021 Add IsKeyNameToPigMD5Force
'* 1.0.3	6/5/2021 Modify New,mNew
'* 1.0.4	8/5/2021 Modify enmValueType,New, add BytesValue,StrValue
'* 1.0.5	10/5/2021 Add BytesBase64Value
'* 1.0.6	11/5/2021 Modify StrValue
'* 1.0.7	17/5/2021 Modify New
'* 1.0.8	8/8/2021  Modify New, and Add IsExpired
'* 1.0.9	10/8/2021  Add KeyValueLen, remove mstrStrValue, modify StrValue
'* 1.0.10	11/8/2021  Add SMNameHead,SMNameBody
'* 1.0.11	13/8/2021  Rename KeyValueLen to ValueLen, add ValueMD5Bytes
'* 1.0.12	13/8/2021  Modify ValueMD5Bytes
'* 1.0.13	16/8/2021  Modify mstrSMNameBody,SMNameBody,SMNameHead
'* 1.0.14	17/8/2021  Modify New
'* 1.0.15	25/8/2021 Remove Imports PigToolsLib, change to PigToolsWinLib, and add 
'* 1.1	    29/8/2021 Chanage PigToolsWinLib to PigToolsLiteLib
'* 1.2	    2/9/2021  Add IsValueTypeOK,ValueMD5Base64
'* 1.3	    17/9/2021  Modify ValueMD5Base64, Add Check
'* 1.4	    2/10/2021  Modify SMNameBody,SMNameHead
'* 1.5	    3/10/2021  Add CompareOther
'* 1.6	    4/10/2021  Add LastRefCacheTime,IsForceRefCache,CopyToMe
'************************************

Imports PigToolsLiteLib
Public Class PigKeyValue
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.6.2"
    Private mabKeyValue As Byte()
    ''' <summary>
    ''' 父对象
    ''' </summary>
    ''' <returns></returns>
    Public Property Parent As PigKeyValueApp

    ''' <summary>
    ''' 键值标识，为键名的PigMD5
    ''' </summary>
    ''' <returns>键值标识</returns>
    Public ReadOnly Property KeyName As String
    ''' <summary>
    ''' 键值类型
    ''' </summary>
    ''' <returns></returns>
    Private mintValueType As enmValueType
    Public Property ValueType As enmValueType
        Get
            Return mintValueType
        End Get
        Friend Set(value As enmValueType)
            mintValueType = value
        End Set
    End Property
    ''' <summary>
    '''   过期时间
    ''' </summary>
    Private mdteExpTime As DateTime
    Public Property ExpTime As DateTime
        Get
            Return mdteExpTime
        End Get
        Friend Set(value As DateTime)
            mdteExpTime = value
        End Set
    End Property
    Friend Property LastRefCacheTime As DateTime = Now

    ''' <summary>
    ''' 字符串值
    ''' </summary>
    'Private mstrStrValue As String
    Public ReadOnly Property StrValue As String
        Get
            Try
                Select Case Me.ValueType
                    Case enmValueType.Bytes, enmValueType.EncBytes, enmValueType.ZipBytes, enmValueType.ZipEncBytes
                        Return New PigText(mabKeyValue).Base64
                    Case enmValueType.Text
                        Return New PigText(mabKeyValue, PigText.enmTextType.UTF8).Text
                    Case Else
                        Return ""
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("StrValue", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ValueLen As Long
        Get
            Try
                Return mabKeyValue.Length
            Catch ex As Exception
                Me.SetSubErrInf("ValueLen", ex)
                Return -1
            End Try
        End Get
    End Property



    ''' <summary>
    ''' Value type, non text type, saved in byte array
    ''' </summary>
    Public Enum enmValueType
        ''' <summary>
        ''' text
        ''' </summary>
        Unknow = 0
        ''' <summary>
        ''' text
        ''' </summary>
        Text = 10
        ''' <summary>
        ''' Byte array
        ''' </summary>
        Bytes = 20
        ''' <summary>
        ''' Compressed byte array
        ''' </summary>
        ZipBytes = 30
        ''' <summary>
        ''' Encrypted byte array
        ''' </summary>
        EncBytes = 40
        ''' <summary>
        ''' Compressed encrypted byte array
        ''' </summary>
        ZipEncBytes = 50
    End Enum

    ''' <summary>
    ''' 键值的PigMD5
    ''' </summary>
    ''' 
    Private mstrValueMD5 As String = ""
    Public ReadOnly Property ValueMD5 As String
        Get
            Try
                Select Case Me.ValueType
                    Case enmValueType.Bytes, enmValueType.EncBytes, enmValueType.ZipBytes
                        If mstrValueMD5 = "" Then mstrValueMD5 = New PigMD5(mabKeyValue).PigMD5
                    Case enmValueType.Text
                        If mstrValueMD5 = "" Then mstrValueMD5 = New PigMD5(Me.StrValue, PigMD5.enmTextType.UTF8).PigMD5
                End Select
                Return mstrValueMD5
            Catch ex As Exception
                Me.SetSubErrInf("ValueMD5", ex)
                Return ""
            End Try
        End Get
    End Property

    Private mabValueMD5 As Byte()
    Public ReadOnly Property ValueMD5Bytes As Byte()
        Get
            Try
                Select Case Me.ValueType
                    Case enmValueType.Bytes, enmValueType.EncBytes, enmValueType.ZipBytes
                        If mabValueMD5 Is Nothing Then mabValueMD5 = New PigMD5(mabKeyValue).PigMD5Bytes
                    Case enmValueType.Text
                        If mabValueMD5 Is Nothing Then mabValueMD5 = New PigMD5(Me.StrValue, PigMD5.enmTextType.UTF8).PigMD5Bytes
                End Select
                Return mabValueMD5
            Catch ex As Exception
                Me.SetSubErrInf("ValueMD5Bytes", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property ValueMD5Base64 As String
        Get
            Try
                Dim pbMD5 As New PigBytes(Me.ValueMD5Bytes)
                ValueMD5Base64 = pbMD5.Base64Str
                pbMD5 = Nothing
            Catch ex As Exception
                Me.SetSubErrInf("ValueMD5Base64", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property BytesBase64Value As String
        Get
            Try
                Return New PigBytes(Me.BytesValue).Base64Str
            Catch ex As Exception
                Me.SetSubErrInf("BytesBase64Value", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property BytesValue As Byte()
        Get
            Try
                Select Case Me.ValueType
                    Case enmValueType.Bytes
                        Return mabKeyValue
                    Case enmValueType.Text
                        Return New PigText(Me.StrValue, PigText.enmTextType.UTF8).TextBytes
                    Case Else
                        Return Nothing
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("BytesValue", ex)
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    ''' <param name="MatchValueMD5">匹配PigMD5，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String, Optional MatchValueMD5 As String = "")
        MyBase.New(CLS_VERSION)
        Try
            If KeyName.Length > 128 Then Throw New Exception("The KeyName length cannot exceed 128 bytes.")
            Me.KeyName = KeyName
            If ExpTime < Now Then Throw New Exception("ExpTime unreasonable.")
            If MatchValueMD5.Length > 0 Then
                Dim strValueMD5 As String = New PigMD5(KeyValue, PigMD5.enmTextType.UTF8).PigMD5
                If strValueMD5 <> MatchValueMD5 Then Throw New Exception("ValueMD5 not match。")
                mstrValueMD5 = MatchValueMD5
            End If
            Me.ExpTime = ExpTime
            Me.ValueType = enmValueType.Text
            Dim oPigText As New PigText(KeyValue, PigText.enmTextType.UTF8)
            mabKeyValue = oPigText.TextBytes
            oPigText = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">字节数组键值</param>
    ''' <param name="ValueType">键值类型</param>
    ''' <param name="MatchValueMD5Bytes">匹配PigMD5数组，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), ValueType As enmValueType, Optional MatchValueMD5Bytes As Byte() = Nothing)
        MyBase.New(CLS_VERSION)
        Try
            If KeyName.Length > 128 Then Throw New Exception("The KeyName length cannot exceed 128 bytes.")
            Me.KeyName = KeyName
            If ExpTime < Now Then Throw New Exception("ExpTime unreasonable.")
            If Not MatchValueMD5Bytes Is Nothing Then
                Dim oPigMD5 As PigMD5 = New PigMD5(KeyValue)
                If Me.mIsBytesMatch(oPigMD5.PigMD5Bytes, MatchValueMD5Bytes) = False Then Throw New Exception("ValueMD5 not match.")
            End If
            Me.ExpTime = ExpTime
            Me.ValueType = ValueType
            mabKeyValue = KeyValue
            mabValueMD5 = MatchValueMD5Bytes
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Private Function mIsBytesMatch(ByRef SrcBytes As Byte(), ByRef MatchBytes As Byte()) As Boolean
        Try
#If NET40_OR_GREATER Then
            Return SrcBytes.SequenceEqual(MatchBytes)
#Else
            Dim i As Long
            If SrcBytes.Length <> MatchBytes.Length Then
                Return False
            Else
                mIsBytesMatch = True
                For i = 0 To SrcBytes.Length - 1
                    If SrcBytes(i) <> MatchBytes(i) Then
                        mIsBytesMatch = False
                        Exit For
                    End If
                Next
            End If

#End If
        Catch ex As Exception
            Me.SetSubErrInf("mIsBytesMatch", ex)
            Return False
        End Try
    End Function

    Public ReadOnly Property IsValueTypeOK As Boolean
        Get
            Select Case Me.ValueType
                Case enmValueType.Bytes, enmValueType.EncBytes, enmValueType.Text, enmValueType.ZipBytes, enmValueType.ZipEncBytes
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Public ReadOnly Property IsExpired As Boolean
        Get
            If Me.ExpTime < Now Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    Private mstrSMNameHead As String = ""
    Friend ReadOnly Property SMNameHead() As String
        Get
            Try
                If mstrSMNameHead.Length = 0 Then
                    If Me.Parent Is Nothing Then Throw New Exception("Parent Is Nothing")
                    Dim strSMName As String = Me.Parent.ShareMemRoot & "." & Me.KeyName & ".Head"
                    Dim oPigMD5 As New PigMD5(strSMName, PigMD5.enmTextType.UTF8)
                    mstrSMNameHead = oPigMD5.PigMD5
                    oPigMD5 = Nothing
                End If
                Return mstrSMNameHead
            Catch ex As Exception
                Me.SetSubErrInf("SMNameHead.Get", ex)
                mstrSMNameHead = ""
                Return ""
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 共享内存体
    ''' </summary>
    Private mstrSMNameBody As String = ""
    Friend ReadOnly Property SMNameBody() As String
        Get
            Try
                If mstrSMNameBody.Length = 0 Then
                    If Me.Parent Is Nothing Then Throw New Exception("Parent Is Nothing")
                    Dim strSMName As String = Me.Parent.ShareMemRoot & "." & Me.KeyName & ".Body"
                    Dim oPigMD5 As New PigMD5(strSMName, PigMD5.enmTextType.UTF8)
                    mstrSMNameBody = oPigMD5.PigMD5
                    oPigMD5 = Nothing
                End If
                Return mstrSMNameBody
            Catch ex As Exception
                Me.SetSubErrInf("SMNameBody.Get", ex)
                mstrSMNameBody = ""
                Return ""
            End Try
        End Get
    End Property

    Public Function Check() As String
        Try
            If Me.IsExpired = True Then
                Throw New Exception("IsExpired")
            ElseIf Me.ValueLen = 0 Then
                Throw New Exception("ValueLen is zero.")
            ElseIf Me.IsValueTypeOK = False Then
                Throw New Exception("ValueType is invalid.")
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Friend Function CopyToMe(ByRef SrcItem As PigKeyValue) As String
        Try
            If SrcItem Is Nothing Then
                Throw New Exception("SrcItem Is Nothing")
            Else
                With SrcItem
                    If .KeyName <> Me.KeyName Then Throw New Exception("KeyName mismatch")
                    If .ValueType <> Me.ValueType Then Me.ValueType = .ValueType
                    If .ExpTime <> Me.ExpTime Then Me.ExpTime = .ExpTime
                    Me.mabKeyValue = .BytesValue
                    Me.mabValueMD5 = .ValueMD5Bytes
                    Me.LastRefCacheTime = Now
                End With
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CopyToMe", ex)
        End Try
    End Function

    Friend Function CompareOther(ByRef OtherItem As PigKeyValue) As Boolean
        Try
            If OtherItem Is Nothing Then
                Return False
            Else
                CompareOther = False
                With OtherItem
                    If .KeyName <> Me.KeyName Then Exit Function
                    If .ValueType <> Me.ValueType Then Exit Function
                    If .ExpTime <> Me.ExpTime Then Exit Function
                    If .ValueLen <> Me.ValueLen Then Exit Function
                    If .ValueMD5 <> Me.ValueMD5 Then Exit Function
                End With
                Return True
            End If
        Catch ex As Exception
            Me.SetSubErrInf("CompareOther", ex)
            Return False
        End Try
    End Function

    Friend Function IsForceRefCache() As Boolean
        Try
            If Math.Abs(DateDiff(DateInterval.Second, Me.LastRefCacheTime, Now)) > Me.Parent.ForceRefCacheTime Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsForceRefCache", ex)
            Return False
        End Try
    End Function

End Class
