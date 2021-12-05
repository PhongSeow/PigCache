'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 31/8/2021
'* 1.1	21/9/2021 
'* 1.2	4/12/2021 Add more new
'* 1.3	5/12/2021 Add PigBaseMini
'************************************
Imports PigKeyCacheLib
Imports PigToolsLiteLib
Public Class PigKeyValue
    Inherits PigKeyCacheLib.PigKeyValue
    Private Const CLS_VERSION As String = "1.3.1"
    Private mabKeyValue As Byte()

    '-----PigBaseMini-----Begin
    Private moPigBaseMini As New PigBaseMini(CLS_VERSION)
    Public Overloads ReadOnly Property LastErr As String
        Get
            Return moPigBaseMini.LastErr
        End Get
    End Property

    Private Sub ClearErr()
        moPigBaseMini.ClearErr()
    End Sub

    Private Sub SetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        moPigBaseMini.SetSubErrInf(SubName, exIn, IsStackTrace)
    End Sub

    Private Sub SetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        moPigBaseMini.SetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Sub

    Private Sub PrintDebugLog(SubName As String, LogInf As String)
        moPigBaseMini.PrintDebugLog(SubName, LogInf)
    End Sub

    Private Sub PrintDebugLog(SubName As String, StepName As String, LogInf As String)
        moPigBaseMini.PrintDebugLog(SubName, StepName, LogInf)
    End Sub

    Private Function GetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        Return moPigBaseMini.GetSubErrInf(SubName, exIn, IsStackTrace)
    End Function

    Private Function GetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        Return moPigBaseMini.GetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Function

    '-----PigBaseMini-----End

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String)
        MyBase.New(KeyName, ExpTime, KeyValue)
    End Sub


    Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte)
        MyBase.New(KeyName, ExpTime, KeyValue)
    End Sub
    Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte, SaveType As enmSaveType)
        MyBase.New(KeyName, ExpTime, KeyValue, SaveType)
    End Sub
    'Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte, MatchValueMD5Bytes() As Byte)
    '    MyBase.New(KeyName, ExpTime, KeyValue, MatchValueMD5Bytes)
    'End Sub
    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String, TextType As PigText.enmTextType)
        MyBase.New(KeyName, ExpTime, KeyValue, TextType)
    End Sub
    'Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String, MatchValueMD5 As String)
    '    MyBase.New(KeyName, ExpTime, KeyValue, MatchValueMD5)
    'End Sub
    'Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte, SaveType As enmSaveType, MatchValueMD5Bytes() As Byte)
    '    MyBase.New(KeyName, ExpTime, KeyValue, SaveType, MatchValueMD5Bytes)
    'End Sub
    'Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String, TextType As PigText.enmTextType, MatchValueMD5 As String)
    '    MyBase.New(KeyName, ExpTime, KeyValue, TextType, MatchValueMD5)
    'End Sub

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

End Class
