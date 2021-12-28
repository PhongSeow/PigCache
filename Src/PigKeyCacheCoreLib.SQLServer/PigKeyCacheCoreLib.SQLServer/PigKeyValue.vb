'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 3.0
'* Create Time: 31/8/2021
'* 1.1	21/9/2021 
'* 1.2	3/12/2021 Add more new
'* 1.3	6/12/2021 Add GetSaveData,Check,InitBytesBySave
'* 2.0	15/12/2021 Modify New
'* 3.0	28/12/2021 Code rewriting
'************************************
Imports PigKeyCacheLib
Imports PigToolsLiteLib
Public Class PigKeyValue
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "3.0.18"
    Friend Property Obj As PigKeyCacheLib.PigKeyValue

    ''' <summary>
    ''' Save data type, Decide whether to process the saved data.
    ''' </summary>
    Public Enum EnmSaveType
        ''' <summary>
        ''' Original, not processed
        ''' </summary>
        Original = 0
        ''' <summary>
        ''' Save space and compress and save data
        ''' </summary>
        SaveSpace = 1
        ''' <summary>
        ''' It is confidential and saves space. The data is compressed and encrypted
        ''' </summary>
        EncSaveSpace = 2
    End Enum

    ''' <summary>
    ''' Value types, including text and binary
    ''' </summary>
    Public Enum EnmValueType
        ''' <summary>
        ''' text
        ''' </summary>
        Unknow = 0
        ''' <summary>
        ''' text
        ''' </summary>
        Text = 1
        ''' <summary>
        ''' Byte array
        ''' </summary>
        Bytes = 2
        ''' <summary>
        ''' Compressed byte array
        ''' </summary>
    End Enum

    Public Sub New(KeyName As String)
        MyBase.New(CLS_VERSION)
        Me.Obj = New PigKeyCacheLib.PigKeyValue(KeyName)
    End Sub

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String)
        MyBase.New(CLS_VERSION)
        Me.Obj = New PigKeyCacheLib.PigKeyValue(KeyName, ExpTime, KeyValue)
    End Sub

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String, TextType As PigText.enmTextType)
        MyBase.New(CLS_VERSION)
        Me.Obj = New PigKeyCacheLib.PigKeyValue(KeyName, ExpTime, KeyValue, TextType)
    End Sub

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String, TextType As PigText.enmTextType, SaveType As EnmSaveType)
        MyBase.New(CLS_VERSION)
        Me.Obj = New PigKeyCacheLib.PigKeyValue(KeyName, ExpTime, KeyValue, TextType, SaveType)
    End Sub

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte)
        MyBase.New(CLS_VERSION)
        Me.Obj = New PigKeyCacheLib.PigKeyValue(KeyName, ExpTime, KeyValue)
    End Sub
    Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte, SaveType As EnmSaveType)
        MyBase.New(CLS_VERSION)
        Me.Obj = New PigKeyCacheLib.PigKeyValue(KeyName, ExpTime, KeyValue, SaveType)
    End Sub

    Public ReadOnly Property KeyName As String
        Get
            Return Me.Obj.KeyName
        End Get
    End Property

    Public ReadOnly Property ExpTime As Date
        Get
            Return Me.Obj.ExpTime
        End Get
    End Property


    Public ReadOnly Property IsExpired As Boolean
        Get
            Return Me.Obj.IsExpired
        End Get
    End Property

    Public ReadOnly Property SaveType As EnmSaveType
        Get
            Return Me.Obj.SaveType
        End Get
    End Property


    Public ReadOnly Property BytesValue As Byte()
        Get
            Return Me.Obj.BytesValue
        End Get
    End Property

    Public ReadOnly Property StrValue As String
        Get
            Return Me.Obj.StrValue
        End Get
    End Property

    Public ReadOnly Property TextType As PigText.enmTextType
        Get
            Return Me.Obj.TextType
        End Get
    End Property

    Public ReadOnly Property ValueLen As Long
        Get
            Return Me.Obj.ValueLen
        End Get
    End Property


    Public ReadOnly Property ValueType As EnmValueType
        Get
            Return Me.Obj.ValueType
        End Get
    End Property

    Friend ReadOnly Property fHeadData As PigBytes
        Get
            Return Me.Obj.HeadData
        End Get
    End Property

    Friend ReadOnly Property fBodyData As PigBytes
        Get
            Try
                Return Me.Obj.BodyData
            Catch ex As Exception
                Me.PrintDebugLog("fBodyData", ex.Message.ToString)
                Return Nothing
            End Try
        End Get
    End Property

    Friend Overloads Function fCheck() As String
        Try
            Return Me.Obj.Check
        Catch ex As Exception
            Return Me.GetSubErrInf("fCheck", ex)
        End Try
    End Function

End Class
