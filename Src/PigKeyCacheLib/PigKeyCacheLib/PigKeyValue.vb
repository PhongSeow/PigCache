'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 2.3
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
'* 1.7	    13/11/2021  Modify BytesValue,New
'* 1.8	    20/11/2021  Add OriginalBytesValue, modify BytesValue
'* 1.9	    21/11/2021  Modify New, add fSaveValueLen,fSaveBytesValue,fInitBytesBySave,IsDataReady, Rename fCompareOther,fCopyToMe,fIsForceRefCache,fSMNameBody,fSMNameHead
'* 1.10	    24/11/2021  Modify fSaveValueLen,fInitBytesBySave
'* 1.11	    25/11/2021  Modify StrValue,Check
'* 2.0	    27/11/2021  Add enmSaveType,TextType，IsValueTypeOK,IsTextTypeOK,mNew, and modify enmValueType,New
'* 2.1	    28/11/2021  Remove fSMNameHead,fSMNameBody
'* 2.2	    30/11/2021  Add fGetSaveData, modify ValueLen,StrValue,mInitSMNameHeadBody
'* 2.3	    2/12/2021  Add mNew for Byte, modify fGetSaveData,ValueLen,fCopyToMe
'************************************

Imports PigToolsLiteLib
Public Class PigKeyValue
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "2.3.10"
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
    Private mstrKeyName As String
    Public Property KeyName As String
        Get
            Return mstrKeyName
        End Get
        Friend Set(value As String)
            mstrKeyName = value
        End Set
    End Property

    ''' <summary>
    ''' 保存数据类型|Save data type
    ''' </summary>
    ''' <returns></returns>
    Private mintSaveType As enmSaveType
    Public Property SaveType As enmSaveType
        Get
            Return mintSaveType
        End Get
        Friend Set(value As enmSaveType)
            mintSaveType = value
        End Set
    End Property

    ''' <summary>
    ''' 值类型
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
    ''' 文本编码类型
    ''' </summary>
    ''' <returns></returns>
    Private mintTextType As PigText.enmTextType
    Public Property TextType As PigText.enmTextType
        Get
            Return mintTextType
        End Get
        Friend Set(value As PigText.enmTextType)
            mintTextType = value
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
    ''' 字符串值，非文本类型以 Base64 格式表示|String value, non text type, expressed in Base64 format
    ''' </summary>
    Private mstrValue As String = ""
    Public ReadOnly Property StrValue As String
        Get
            Dim strStepName As String = ""
            Try
                Select Case Me.ValueType
                    Case enmValueType.Bytes
                        strStepName = "New PigText(Bytes)"
                        Dim oPigText As New PigText(mabKeyValue)
                        If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
                        mstrValue = oPigText.Base64
                        oPigText = Nothing
                    Case enmValueType.Text
                        If mstrValue.Length = 0 Then
                            strStepName = "New PigText(Text)"
                            Dim oPigText As New PigText(mabKeyValue, Me.TextType)
                            If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
                            mstrValue = oPigText.Text
                            oPigText = Nothing
                        End If
                    Case Else
                        strStepName = Me.ValueType.ToString
                        Throw New Exception("Unsupported")
                End Select
                Return mstrValue
            Catch ex As Exception
                Dim strRet As String = Me.GetSubErrInf("StrValue", ex)
                Me.PrintDebugLog("As Exception", strRet)
                mstrValue = ""
                Return ""
            End Try
        End Get
    End Property

    'Private mlngSaveValueLen As Long = 0
    'Friend ReadOnly Property fSaveValueLen As Long
    '    Get
    '        Dim strStepName As String = ""
    '        Dim strRet As String = ""
    '        Try
    '            Select Case Me.ValueType
    '                Case enmValueType.Text
    '                    If mlngSaveValueLen <= 0 Then mlngSaveValueLen = mstrValue.Length
    '                Case enmValueType.Bytes
    '                    If mlngSaveValueLen <= 0 Then
    '                        strStepName = "fRefSaveValue"
    '                        strRet = Me.fRefSaveValue
    '                        If strRet <> "OK" Then Throw New Exception(strRet)
    '                    End If
    '                Case Else
    '                    Throw New Exception("Invalid ValueType is " & Me.ValueType)
    '            End Select
    '            Return mstrValue.Length
    '        Catch ex As Exception
    '            strRet = Me.GetSubErrInf("fSaveValueLen", strStepName, ex)
    '            Me.PrintDebugLog("As Exception", strRet)
    '            mlngSaveValueLen = -1
    '            Return mlngSaveValueLen
    '        End Try
    '    End Get
    'End Property

    Public ReadOnly Property ValueLen As Long
        Get
            Try
                Select Case Me.ValueType
                    Case enmValueType.Bytes
                        Return Me.BytesValue.Length
                    Case enmValueType.Text
                        Return Me.StrValue.Length
                    Case Else
                        Throw New Exception("Unsupported ValueType is " & Me.ValueType.ToString)
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("ValueLen", ex)
                Return -1
            End Try
        End Get
    End Property


    ''' <summary>
    ''' Value types, including text and binary
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
        'ZipBytes = 30
        '''' <summary>
        '''' Encrypted byte array
        '''' </summary>
        'EncBytes = 40
        '''' <summary>
        '''' Compressed encrypted byte array
        '''' </summary>
        'ZipEncBytes = 50
    End Enum


    ''' <summary>
    ''' Save data type, Decide whether to process the saved data.
    ''' </summary>
    Public Enum enmSaveType
        ''' <summary>
        ''' Original, not processed
        ''' </summary>
        Original = 0
        ''' <summary>
        ''' Save space and compress and save data
        ''' </summary>
        SaveSpace = 10
        ''' <summary>
        ''' It is confidential and saves space. The data is compressed and encrypted
        ''' </summary>
        EncSaveSpace = 20
    End Enum

    ''' <summary>
    ''' 键值的PigMD5
    ''' </summary>
    ''' 
    Public ReadOnly Property ValueMD5 As String
        Get
            Try
                Dim oPigMD5 As New PigMD5(mabValueMD5)
                If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
                ValueMD5 = oPigMD5.PigMD5
                oPigMD5 = Nothing
            Catch ex As Exception
                Dim strRet As String = Me.GetSubErrInf("ValueMD5", ex)
                Me.PrintDebugLog("As Exception", strRet)
                Return ""
            End Try
        End Get
    End Property

    Private mabValueMD5 As Byte()
    Public ReadOnly Property ValueMD5Bytes As Byte()
        Get
            Try
                If mabValueMD5 Is Nothing Then
                    Dim oPigMD5 As New PigMD5(mabKeyValue)
                    If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
                    mabValueMD5 = oPigMD5.PigMD5Bytes
                    oPigMD5 = Nothing
                End If
                Return mabValueMD5
            Catch ex As Exception
                Dim strRet As String = Me.GetSubErrInf("ValueMD5Bytes", ex)
                Me.PrintDebugLog("As Exception", strRet)
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
                Dim strRet As String = Me.GetSubErrInf("BytesBase64Value", ex)
                Me.PrintDebugLog("BytesBase64Value", strRet)
                Return ""
            End Try
        End Get
    End Property

    'Friend Function fRefSaveValue() As String
    '    Dim strStepName As String = ""
    '    Try
    '        Select Case Me.ValueType
    '            Case enmValueType.Bytes
    '                Select Case Me.SaveType
    '                    Case enmSaveType.Original
    '                        mlngSaveValueLen = mabKeyValue.Length
    '                    Case enmSaveType.SaveSpace
    '                        strStepName = "New PigBytes(ZipBytes)"
    '                        Dim oPigBytes As New PigBytes(mabKeyValue)
    '                        If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
    '                        strStepName = "Compress(ZipBytes)"
    '                        oPigBytes.Compress()
    '                        If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
    '                        mabSaveValue = oPigBytes.Main
    '                        oPigBytes = Nothing
    '                        mlngSaveValueLen = mabSaveValue.Length
    '                    Case enmSaveType.EncSaveSpace
    '                End Select
    '            Case Else
    '                Throw New Exception("Invalid ValueType is " & Me.ValueType.ToString)
    '        End Select
    '        Return "OK"
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf("fRefSaveValue", strStepName, ex)
    '    End Try
    'End Function

    Private mabSaveValue As Byte()

    Friend Function fGetSaveData(ByRef SaveBytes As Byte(), ByRef SavePigMD5 As Byte()) As String
        Dim strStepName As String = ""
        Dim strRet As String = ""
        Try
            Select Case Me.ValueType
                Case enmValueType.Text
                    strStepName = "Set SaveBytes(Text)"
                    SaveBytes = mabKeyValue
                Case enmValueType.Bytes
                    Select Case Me.SaveType
                        Case enmSaveType.Original
                            strStepName = "Set SaveBytes(Original)"
                            SaveBytes = mabKeyValue
                        Case enmSaveType.SaveSpace
                            strStepName = "New PigBytes(SaveSpace)"
                            Dim oPigBytes As New PigBytes(mabKeyValue)
                            If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
                            strStepName = "Compress(SaveSpace)"
                            oPigBytes.Compress()
                            If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
                            strStepName = "Set SaveBytes(SaveSpace)"
                            SaveBytes = oPigBytes.Main
                            oPigBytes = Nothing
                        Case enmSaveType.EncSaveSpace
                            strStepName = "Set SaveBytes(EncSaveSpace)"
                            Throw New Exception("Not yet supported")
                    End Select
                Case Else
                    Throw New Exception("Not supported ValueType is " & Me.ValueType.ToString)
            End Select
            strStepName = "New PigMD5"
            Dim oPigMD5 As New PigMD5(SaveBytes)
            If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
            SavePigMD5 = oPigMD5.PigMD5Bytes
            oPigMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            strRet = Me.GetSubErrInf("fGetSaveData", strStepName, ex)
            Return strRet
        End Try
    End Function

    'Friend ReadOnly Property fSaveBytesValue As Byte()
    '    Get
    '        Dim strStepName As String = ""
    '        Dim strRet As String = ""
    '        Try
    '            Select Case Me.ValueType
    '                Case enmValueType.Bytes
    '                    Select Case Me.SaveType
    '                        Case enmSaveType.Original
    '                            Return mabKeyValue
    '                        Case enmSaveType.SaveSpace, enmSaveType.EncSaveSpace
    '                            If Me.fSaveValueLen = 0 Then
    '                                strStepName = "fRefSaveValue"
    '                                strRet = Me.fRefSaveValue
    '                                If strRet <> "OK" Then Throw New Exception(strRet)
    '                            End If

    '                    End Select
    '                Case Else
    '                    Throw New Exception("Not supported ValueType is " & Me.ValueType.ToString)
    '            End Select
    '        Catch ex As Exception
    '            strRet = Me.GetSubErrInf("fSaveBytesValue", strStepName, ex)
    '            Me.PrintDebugLog("As Exception", strRet)
    '            mlngSaveValueLen = -1
    '            Return Nothing
    '        End Try
    '    End Get
    'End Property


    Public ReadOnly Property BytesValue As Byte()
        Get
            Dim strStepName As String = ""
            Try
                Select Case Me.ValueType
                    Case enmValueType.Text
                        Return mabKeyValue
                    Case enmValueType.Bytes
                        Return mabKeyValue
                    Case Else
                        strStepName = Me.ValueType.ToString
                        Throw New Exception("Not supported now")
                End Select
            Catch ex As Exception
                Dim strRet As String = Me.GetSubErrInf("BytesValue", ex)
                Me.PrintDebugLog("As Exception", strStepName, strRet)
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub New(KeyName As String)
        MyBase.New(CLS_VERSION)
        Dim strStepName As String = ""
        Try
            Me.KeyName = KeyName
            strStepName = "mInitSMNameHeadBody"
            Dim strRet As String = Me.mInitSMNameHeadBody()
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", "mInitSMNameHeadBody", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Initialize the class with the saved data
    ''' </summary>
    ''' <returns></returns>
    Friend Function fInitBytesBySave(KeyName As String, SuSMHead As PigKeyValueApp.StruSMHead, ByRef KeyValue As Byte()) As String
        Dim strStepName As String = ""
        Dim strRet As String = ""
        Try
            With Me
                .KeyName = KeyName
                .ExpTime = SuSMHead.ExpTime
                .ValueType = SuSMHead.ValueType
                .TextType = SuSMHead.TextType
                .SaveType = SuSMHead.SaveType
                Dim oPigBytes As PigBytes = Nothing
                Select Case Me.ValueType
                    Case enmValueType.Text, enmValueType.Bytes
                        Select Case Me.SaveType
                            Case enmSaveType.Original
                                mabKeyValue = KeyValue
                                mstrValue = ""
                            Case enmSaveType.SaveSpace
                                ReDim mabSaveValue(0)
                                strStepName = "New PigBytes(SaveSpace)"
                                oPigBytes = New PigBytes(KeyValue)
                                If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
                                strStepName = "UnCompress(SaveSpace)"
                                oPigBytes.UnCompress()
                                If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
                                mabKeyValue = oPigBytes.Main
                                oPigBytes = Nothing
                                mstrValue = ""
                            Case enmSaveType.EncSaveSpace
                                strStepName = "SaveType is EncSaveSpace"
                                Throw New Exception("Unsupported")
                        End Select
                End Select
                strStepName = "Check"
                strRet = .Check
                If strRet <> "OK" Then Throw New Exception(strRet)
                strStepName = "New PigBytes(ValueMD5Bytes)"
                oPigBytes = Nothing
                oPigBytes = New PigBytes(ValueMD5Bytes)
                If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
                strStepName = "Check(ValueMD5Bytes)"
                If oPigBytes.IsMatchBytes(.ValueMD5Bytes) = False Then Throw New Exception("Mismatch")
                oPigBytes = Nothing
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("fInitBytesBySave", strStepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    ''' <param name="MatchValueMD5">匹配PigMD5，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String, MatchValueMD5 As String)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, MatchValueMD5)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    ''' <param name="TextType">文本类型</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String, TextType As PigText.enmTextType)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, , TextType)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">文本键值</param>
    ''' <param name="TextType">文本类型</param>
    ''' <param name="MatchValueMD5">匹配PigMD5，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String, TextType As PigText.enmTextType, MatchValueMD5 As String)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, MatchValueMD5, TextType)
    End Sub



    Public Sub mNew(KeyName As String, ExpTime As DateTime, KeyValue As String, Optional MatchValueMD5 As String = "", Optional TextType As PigText.enmTextType = PigText.enmTextType.UTF8, Optional SaveType As enmSaveType = enmSaveType.Original)
        Dim strStepName As String = ""
        Dim strRet As String = ""
        Try
            If MatchValueMD5.Length > 0 Then
                strStepName = "Check KeyValue PigMD5"
                Dim oPigMD5 = New PigMD5(KeyValue, Me.TextType)
                If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
                If oPigMD5.PigMD5 <> MatchValueMD5 Then Throw New Exception("Not match.")
                mabValueMD5 = oPigMD5.PigMD5Bytes
                oPigMD5 = Nothing
            End If
            With Me
                .KeyName = KeyName
                .ExpTime = ExpTime
                .ValueType = enmValueType.Text
                .TextType = TextType
                .SaveType = SaveType
            End With
            strStepName = "New PigText"
            Dim oPigText As New PigText(KeyValue, TextType)
            If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
            mabKeyValue = oPigText.TextBytes
            oPigText = Nothing
            strStepName = "Check"
            strRet = Me.Check
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "mInitSMNameHeadBody"
            strRet = Me.mInitSMNameHeadBody()
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", strStepName, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">字节数组键值</param>
    ''' <param name="SaveType">保存类型</param>
    ''' <param name="MatchValueMD5Bytes">匹配PigMD5数组，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), SaveType As enmSaveType, MatchValueMD5Bytes As Byte())
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, MatchValueMD5Bytes, SaveType)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">字节数组键值</param>
    ''' <param name="SaveType">保存类型</param>
    ''' <param name="MatchValueMD5Bytes">匹配PigMD5数组，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), SaveType As enmSaveType)
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, Nothing, SaveType)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">字节数组键值</param>
    ''' <param name="MatchValueMD5Bytes">匹配PigMD5数组，如果指定则校验</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), MatchValueMD5Bytes As Byte())
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, MatchValueMD5Bytes, enmSaveType.SaveSpace)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="KeyName">键名</param>
    ''' <param name="ExpTime">键值过期时间</param>
    ''' <param name="KeyValue">字节数组键值</param>
    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte())
        MyBase.New(CLS_VERSION)
        Me.mNew(KeyName, ExpTime, KeyValue, Nothing, enmSaveType.SaveSpace)
    End Sub

    Private Sub mNew(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), Optional MatchValueMD5Bytes As Byte() = Nothing, Optional SaveType As enmSaveType = enmSaveType.SaveSpace)
        Dim strStepName As String = ""
        Dim strRet As String = ""
        Try
            If Not MatchValueMD5Bytes Is Nothing Then
                strStepName = "New PigMD5(ValueMD5)"
                Dim oPigMD5 As PigMD5 = New PigMD5(KeyValue)
                If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
                strStepName = "mIsBytesMatch(ValueMD5)"
                If Me.mIsBytesMatch(oPigMD5.PigMD5Bytes, MatchValueMD5Bytes) = False Then Throw New Exception("ValueMD5 not match.")
                mabValueMD5 = MatchValueMD5Bytes
                oPigMD5 = Nothing
            End If
            With Me
                .KeyName = KeyName
                .ExpTime = ExpTime
                .ValueType = enmValueType.Bytes
                .SaveType = SaveType
                .TextType = PigText.enmTextType.UnknowOrBin
            End With
            strStepName = "Set KeyValue"
            mabKeyValue = KeyValue
            strStepName = "Check"
            strRet = Me.Check
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "mInitSMNameHeadBody"
            strRet = Me.mInitSMNameHeadBody()
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", strStepName, ex)
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
                Case enmValueType.Bytes, enmValueType.Text
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Public ReadOnly Property IsTextTypeOK As Boolean
        Get
            Select Case Me.ValueType
                Case enmValueType.Bytes
                    Return False
                Case enmValueType.Text
                    Select Case Me.TextType
                        Case PigText.enmTextType.Ascii, PigText.enmTextType.Unicode, PigText.enmTextType.UTF8
                            Return True
                        Case Else
                            Return False
                    End Select
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Public ReadOnly Property IsSaveTypeOK As Boolean
        Get
            Select Case Me.ValueType
                Case enmValueType.Text
                    Select Case Me.SaveType
                        Case enmSaveType.Original
                            Return True
                        Case Else
                            Return False
                    End Select
                Case enmValueType.Bytes
                    Select Case Me.SaveType
                        Case enmSaveType.EncSaveSpace, enmSaveType.Original, enmSaveType.SaveSpace
                            Return True
                        Case Else
                            Return False
                    End Select
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

    Private mstrSMNameHead As String = ""
    Friend ReadOnly Property fSMNameHead() As String
        Get
            Return mstrSMNameHead
        End Get
    End Property

    Private mstrSMNameBody As String = ""
    Friend ReadOnly Property fSMNameBody() As String
        Get
            Return mstrSMNameBody
        End Get
    End Property

    Private Function mInitSMNameHeadBody() As String
        Try
            If mstrSMNameHead.Length = 0 Then
                Dim oPigMD5 As New PigMD5(Me.KeyName, PigMD5.enmTextType.UTF8)
                If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
                mstrSMNameHead = oPigMD5.PigMD5
                mstrSMNameBody = mstrSMNameHead & ".b"
                mstrSMNameHead = mstrSMNameHead & ".h"
                oPigMD5 = Nothing
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mInitSMNameHeadBody", ex)
        End Try
    End Function

    '''' <summary>
    '''' 共享内存体
    '''' </summary>
    'Private mstrSMNameBody As String = ""
    'Friend ReadOnly Property fSMNameBody() As String
    '    Get
    '        Try
    '            If mstrSMNameBody.Length = 0 Then
    '                If Me.Parent Is Nothing Then Throw New Exception("Parent Is Nothing")
    '                Dim strSMName As String = Me.Parent.ShareMemRoot & "." & Me.KeyName & ".Body"
    '                Dim oPigMD5 As New PigMD5(strSMName, PigMD5.enmTextType.UTF8)
    '                mstrSMNameBody = oPigMD5.PigMD5
    '                oPigMD5 = Nothing
    '            End If
    '            Return mstrSMNameBody
    '        Catch ex As Exception
    '            Me.SetSubErrInf("fSMNameBody.Get", ex)
    '            mstrSMNameBody = ""
    '            Return ""
    '        End Try
    '    End Get
    'End Property

    Public Function Check() As String
        Try
            Select Case Me.KeyName.Length
                Case 1 To 128
                Case Else
                    Throw New Exception("The length of the keyname must be between 1 and 128")
            End Select
            If Me.IsExpired = True Then Throw New Exception("KeyValue is IsExpired")
            If Me.IsValueTypeOK = False Then Throw New Exception("Invalid ValueType is " & Me.ValueType.ToString)
            If Me.IsSaveTypeOK = False Then Throw New Exception("Invalid SaveType is " & Me.SaveType.ToString)
            Select Case Me.ValueType
                Case enmValueType.Text
                    If Me.IsTextTypeOK = False Then Throw New Exception("Invalid TextType is " & Me.TextType.ToString)
                Case enmValueType.Bytes
                    If mabKeyValue Is Nothing Then
                        Throw New Exception("The value of keyValue is undefined")
                    End If
            End Select
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Friend Function fCopyToMe(ByRef SrcItem As PigKeyValue) As String
        Try
            If SrcItem Is Nothing Then
                Throw New Exception("SrcItem Is Nothing")
            Else
                With SrcItem
                    If .KeyName <> Me.KeyName Then Throw New Exception("KeyName mismatch")
                    If .ValueType <> Me.ValueType Then Me.ValueType = .ValueType
                    If .ExpTime <> Me.ExpTime Then Me.ExpTime = .ExpTime
                    If .SaveType <> Me.SaveType Then Me.SaveType = .SaveType
                    If .TextType <> Me.TextType Then Me.TextType = .TextType
                    mstrValue = ""
                    Me.mabKeyValue = .BytesValue
                    Me.mabValueMD5 = .ValueMD5Bytes
                    Me.LastRefCacheTime = Now
                End With
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("fCopyToMe", ex)
        End Try
    End Function

    Friend Function fCompareOther(ByRef OtherItem As PigKeyValue) As Boolean
        Try
            If OtherItem Is Nothing Then
                Return False
            Else
                fCompareOther = False
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
            Me.SetSubErrInf("fCompareOther", ex)
            Return False
        End Try
    End Function

    Friend Function fIsForceRefCache() As Boolean
        Try
            If Math.Abs(DateDiff(DateInterval.Second, Me.LastRefCacheTime, Now)) > Me.Parent.ForceRefCacheTime Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("fIsForceRefCache", ex)
            Return False
        End Try
    End Function

End Class
