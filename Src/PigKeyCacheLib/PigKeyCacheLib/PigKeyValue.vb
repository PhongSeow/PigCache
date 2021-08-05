'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.8
'* Create Time: 11/3/2021
'* 1.0.2	6/4/2021 Add IsKeyNameToPigMD5Force
'* 1.0.3	6/5/2021 Modify New,mNew
'* 1.0.4	8/5/2021 Modify enmValueType,New, add BytesValue,StrValue
'* 1.0.5	10/5/2021 Add BytesBase64Value
'* 1.0.6	11/5/2021 Modify StrValue
'* 1.0.7	17/5/2021 Modify New
'* 1.0.8	8/8/2021  Modify New, and Add IsExpired
'************************************

Imports PigToolsLib
Public Class PigKeyValue
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.8.5"
    Private mabKeyValue As Byte()
    Private mbolIsKeyValueReady As Boolean = False
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
    Public ReadOnly Property ValueType As enmValueType
    ''' <summary>
    '''   过期时间
    ''' </summary>
    Public ReadOnly Property ExpTime As DateTime

    ''' <summary>
    ''' 字符串值
    ''' </summary>
    Private mstrStrValue As String
    Public ReadOnly Property StrValue As String
        Get
            Try
                Select Case Me.ValueType
                    Case enmValueType.Bytes, enmValueType.EncBytes, enmValueType.ZipBytes
                        Return New PigText(mabKeyValue).Text
                    Case enmValueType.Text
                        Return mstrStrValue
                    Case Else
                        Return ""
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("StrValue", ex)
                Return ""
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 值类型，非文本类型转为 base64 保存
    ''' </summary>
    Public Enum enmValueType
        Text = 0 '文本
        Bytes = 1 '字节数组
        ZipBytes = 2 '压缩的字节数组
        EncBytes = 3 '压缩后加密的字节数组
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
            If ExpTime < Now Then Throw New Exception("ExpTime unreasonable.")
            If MatchValueMD5.Length > 0 Then
                Dim strValueMD5 As String = New PigMD5(KeyValue, PigMD5.enmTextType.UTF8).PigMD5
                If strValueMD5 <> MatchValueMD5 Then Throw New Exception("ValueMD5 not match。")
                mstrValueMD5 = MatchValueMD5
            End If
            Me.KeyName = KeyName
            Me.ExpTime = ExpTime
            Me.ValueType = enmValueType.Text
            mbolIsKeyValueReady = True
            mstrStrValue = KeyValue
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
    ''' <param name="MatchValueMD5">匹配PigMD5，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), ValueType As enmValueType, Optional MatchValueMD5 As String = "")
        MyBase.New(CLS_VERSION)
        Try
            If KeyName.Length > 128 Then Throw New Exception("The KeyName length cannot exceed 128 bytes.")
            If ExpTime < Now Then Throw New Exception("ExpTime unreasonable.")
            If MatchValueMD5.Length > 0 Then
                Dim strValueMD5 As String = New PigMD5(KeyValue).PigMD5
                If strValueMD5 <> MatchValueMD5 Then Throw New Exception("ValueMD5 not match.")
                mstrValueMD5 = MatchValueMD5
            End If
            Me.KeyName = KeyName
            Me.ExpTime = ExpTime
            Select Case ValueType
                Case enmValueType.Text
                    mstrStrValue = New PigText(KeyValue, PigText.enmTextType.UTF8).Text
                Case enmValueType.Bytes
                    mabKeyValue = KeyValue
                    mbolIsKeyValueReady = False
                Case enmValueType.ZipBytes, enmValueType.EncBytes
                    mabKeyValue = KeyValue
                    mbolIsKeyValueReady = True
                Case Else
                    Throw New Exception("Unknow ValueType " & ValueType.ToString)
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Public ReadOnly Property IsExpired As Boolean
        Get
            If Me.ExpTime < Now Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property


End Class
