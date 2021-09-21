'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用 SQL Server 版|Piggy key value application for SQL Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6
'* Create Time: 31/8/2021
'* 1.1	1/9/2021 Add mCreateTableKeyValueInf,PigBaseMini,OpenDebug,mIsDBObjExists,GetPigKeyValue
'* 1.2	2/9/2021 Modify mNew,IsPigKeyValueExists,SavePigKeyValue,mCreateTableKeyValueInf, and remove mIsDBObjExists.
'* 1.3	4/9/2021 Modify SavePigKeyValue
'* 1.4	16/9/2021 Modify SavePigKeyValue
'* 1.5	17/9/2021 Modify mCreateTableKeyValueInf,SavePigKeyValue,IsPigKeyValueExists
'* 1.6	21/9/2021 Modify GetPigKeyValue,OpenDebug,SavePigKeyValue
'************************************

Imports PigKeyCacheLib
Imports PigToolsLiteLib
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If

Public Class PigKeyValueApp
    Inherits PigKeyCacheLib.PigKeyValueApp
    Private Const CLS_VERSION As String = "1.6.6"
    Private moConnSQLSrv As ConnSQLSrv
    Private moPigFunc As New PigFunc

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

    Public Overloads Sub OpenDebug()
        Dim strStepName As String = ""
        Try
            strStepName = "moConnSQLSrv.OpenDebug"
            moConnSQLSrv.OpenDebug()
            If moConnSQLSrv.LastErr <> "" Then Throw New Exception(moConnSQLSrv.LastErr)
            strStepName = "PigKeyValues.OpenDebug"
            Me.PigKeyValues.OpenDebug()
            If Me.PigKeyValues.LastErr <> "" Then Throw New Exception(Me.PigKeyValues.LastErr)
            strStepName = "moPigBaseMini.OpenDebug"
            moPigBaseMini.OpenDebug()
            If moPigBaseMini.LastErr <> "" Then Throw New Exception(moPigBaseMini.LastErr)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("OpenDebug", strStepName, ex)
        End Try
    End Sub

    Public Sub New(ConnSQLSrv As ConnSQLSrv)
        Me.mNew(ConnSQLSrv)
    End Sub

    Private Sub mNew(ConnSQLSrv As ConnSQLSrv)
        Dim strStepName As String = ""
        Dim strRet As String = ""
        Try
            strStepName = "Set MyClassName"
            moPigBaseMini.MyClassName = Me.GetType.Name.ToString
            strStepName = "Set moConnSQLSrv"
            moConnSQLSrv = ConnSQLSrv
            strStepName = "New SQLSrvTools"
            Dim oSQLSrvTools As New SQLSrvTools(moConnSQLSrv)
            strStepName = "IsDBObjExists"
            If oSQLSrvTools.IsDBObjExists(PigSQLSrvLib.SQLSrvTools.enmDBObjType.UserTable, "_ptKeyValueInf") = False Then
                strStepName = "mCreateTableKeyValueInf"
                strRet = mCreateTableKeyValueInf()
                If strRet <> "" Then Throw New Exception(strRet)
            End If
            oSQLSrvTools = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mNew", ex)
        End Try
    End Sub

    Public Overloads Function GetPigKeyValue(KeyName As String) As PigKeyValue
        Const SUB_NAME As String = "GetPigKeyValue"
        Dim strStepName As String = ""
        Try
            strStepName = "MyBase.GetPigKeyValue"
            Dim oPigKeyValue As PigKeyCacheLib.PigKeyValue = MyBase.GetPigKeyValue(KeyName)
            If oPigKeyValue Is Nothing Then
                Dim strSQL As String = "SELECT TOP 1 ValueType,ExpTime,KeyValue,ValueMD5 FROM dbo._ptKeyValueInf WITH(NOLOCK) WHERE KeyName=@KeyName AND ExpTime>GETDATE()"
                strStepName = "New CmdSQLSrvText"
                Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
                With oCmdSQLSrvText
                    .ActiveConnection = Me.moConnSQLSrv.Connection
                    .AddPara("@KeyName", SqlDbType.VarChar, 128)
                    .ParaValue("@KeyName") = KeyName
                    strStepName = "Execute"
                    Dim rsAny = .Execute()
                    If .LastErr <> "" Then
                        Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                        Throw New Exception(.LastErr)
                    End If
                    If rsAny.EOF = False Then
                        strStepName = "Check ValueType"
                        Dim intValueType As PigKeyCacheLib.PigKeyValue.enmValueType
                        intValueType = rsAny.Fields.Item("ValueType").IntValue
                        Dim abValue(0) As Byte
                        Select Case intValueType
                            Case PigKeyCacheLib.PigKeyValue.enmValueType.Text
                                Dim oPigText As New PigText(rsAny.Fields.Item("KeyValue").StrValue, PigText.enmTextType.UTF8)
                                ReDim abValue(oPigText.TextBytes.Length - 1)
                                abValue = oPigText.TextBytes
                                oPigText = Nothing
                            Case PigKeyCacheLib.PigKeyValue.enmValueType.Bytes, PigKeyCacheLib.PigKeyValue.enmValueType.EncBytes, PigKeyCacheLib.PigKeyValue.enmValueType.ZipBytes, PigKeyCacheLib.PigKeyValue.enmValueType.ZipEncBytes
                                strStepName &= "(" & KeyName & ")"
                                Throw New Exception("Not support ValueType " & intValueType.ToString & " now.")
                            Case Else
                                strStepName &= "(" & KeyName & ")"
                                Throw New Exception("Invalid ValueType " & intValueType.ToString)
                        End Select
                        strStepName = "New PigBytes ValueMD5"
                        Dim pbMD5 As New PigBytes(rsAny.Fields.Item("ValueMD5").StrValue)
                        If pbMD5.LastErr <> "" Then Throw New Exception(pbMD5.LastErr)
                        strStepName = "New PigKeyValue"
                        oPigKeyValue = New PigKeyCacheLib.PigKeyValue(KeyName, rsAny.Fields.Item("ExpTime").DateValue, abValue, intValueType, pbMD5.Main)
                        If oPigKeyValue.LastErr <> "" Then
                            strStepName &= "(" & KeyName & ")"
                            Throw New Exception(oPigKeyValue.LastErr)
                        End If
                        strStepName = "New PigKeyValue"
                        Me.PigKeyValues.Add(oPigKeyValue)
                        If Me.PigKeyValues.LastErr <> "" Then
                            strStepName &= "(" & KeyName & ")"
                            Throw New Exception(Me.PigKeyValues.LastErr)
                        End If
                    End If
                    rsAny.Close()
                    rsAny = Nothing
                    oCmdSQLSrvText = Nothing
                End With
            End If
            If Not oPigKeyValue Is Nothing Then
                With oPigKeyValue
                    GetPigKeyValue = New PigKeyValue(.KeyName, .ExpTime, .BytesValue, .ValueType, .ValueMD5Bytes)
                End With
                oPigKeyValue = Nothing
            Else
                GetPigKeyValue = Nothing
            End If
            oPigKeyValue = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return Nothing
        End Try
    End Function

    Public Overloads Function IsPigKeyValueExists(KeyName As String) As Boolean
        Const SUB_NAME As String = "IsPigKeyValueExists"
        Dim strStepName As String = ""
        Try
            Dim strSQL As String = "SELECT TOP 1 1 FROM dbo._ptKeyValueInf WITH(NOLOCK) WHERE KeyName=@KeyName AND ExpTime>GETDATE()"
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@KeyName", SqlDbType.VarChar, 128)
                .ParaValue("@KeyName") = KeyName
                strStepName = "Execute"
                Dim rsAny = .Execute()
                If .LastErr <> "" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                If rsAny.EOF = False Then
                    IsPigKeyValueExists = True
                Else
                    IsPigKeyValueExists = False
                End If
                strStepName = ""
                strStepName = "rsAny.Close"
                rsAny.Close()
                rsAny = Nothing
            End With
            oCmdSQLSrvText = Nothing
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return False
        End Try
    End Function


    Public Overloads Sub SavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True)
        Const SUB_NAME As String = "SavePigKeyValue"
        Dim strStepName As String = "", strRet As String = ""
        Try
            Dim strKeyName As String = NewItem.KeyName
            strStepName = "Check NewItem"
            strRet = NewItem.Check
            If strRet <> "OK" Then
                strStepName &= "(" & strKeyName & ")"
                Throw New Exception(strRet)
            End If
            If IsOverwrite = False Then
                If Me.IsPigKeyValueExists(strKeyName) = True Then
                    strStepName &= "(" & strKeyName & ")"
                    Throw New Exception("PigKeyValue Exists")
                End If
            End If
            Dim strSQL As String = ""
            moPigFunc.AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT TOP 1 1 FROM dbo._ptKeyValueInf WHERE KeyName=@KeyName)")
            moPigFunc.AddMultiLineText(strSQL, "INSERT INTO dbo._ptKeyValueInf(KeyName,ValueType,ExpTime,KeyValue,ValueMD5)VALUES(@KeyName,@ValueType,@ExpTime,@KeyValue,@ValueMD5)", 1)
            moPigFunc.AddMultiLineText(strSQL, "ELSE")
            moPigFunc.AddMultiLineText(strSQL, "UPDATE dbo._ptKeyValueInf SET ValueType=@ValueType,ExpTime=@ExpTime,KeyValue=@KeyValue,ValueMD5=@ValueMD5", 1)
            moPigFunc.AddMultiLineText(strSQL, "WHERE KeyName=@KeyName", 1)
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@KeyName", SqlDbType.VarChar, 128)
                .AddPara("@ValueType", SqlDbType.Int)
                .AddPara("@ExpTime", SqlDbType.DateTime)
                .AddPara("@KeyValue", SqlDbType.VarChar, -1)
                .AddPara("@ValueMD5", SqlDbType.VarChar, 64)
                .ParaValue("@KeyName") = strKeyName
                .ParaValue("@ExpTime") = NewItem.ExpTime
                .ParaValue("@ValueType") = NewItem.ValueType
                Select Case NewItem.ValueType
                    Case PigKeyCacheLib.PigKeyValue.enmValueType.Text
                        .ParaValue("@KeyValue") = NewItem.StrValue
                    Case Else
                        .ParaValue("@KeyValue") = NewItem.BytesBase64Value
                End Select
                .ParaValue("@ValueMD5") = NewItem.ValueMD5Base64
                strStepName = "ExecuteNonQuery"
                strRet = .ExecuteNonQuery
                If strRet <> "OK" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(strRet)
                ElseIf .RecordsAffected <= 0 Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception("RecordsAffected=" & .RecordsAffected)
                End If
            End With
            strStepName = "MyBase.SavePigKeyValue"
            MyBase.SavePigKeyValue(NewItem, IsOverwrite)
            If MyBase.LastErr <> "" Then Throw New Exception(MyBase.LastErr)
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
        End Try
    End Sub


    Private Function mCreateTableKeyValueInf() As String
        Const SUB_NAME As String = "mCreateTableKeyValueInf"
        Dim strStepName As String = "", strRet As String = ""
        Try
            Dim strTabName As String = ""
            Dim strSQL As String = ""
            moPigFunc.AddMultiLineText(strSQL, "CREATE TABLE dbo._ptKeyValueInf(")
            moPigFunc.AddMultiLineText(strSQL, "KeyName varchar(128) NOT NULL,", 1)
            moPigFunc.AddMultiLineText(strSQL, "ValueType int NOT NULL DEFAULT((0)),", 1)
            moPigFunc.AddMultiLineText(strSQL, "ExpTime datetime NOT NULL,", 1)
            moPigFunc.AddMultiLineText(strSQL, "KeyValue varchar(max)NOT NULL DEFAULT (''),", 1)
            moPigFunc.AddMultiLineText(strSQL, "ValueMD5 varchar(64)NOT NULL DEFAULT (''),", 1)
            moPigFunc.AddMultiLineText(strSQL, "CONSTRAINT PK_ptKeyValueInf PRIMARY KEY CLUSTERED(KeyName)", 1)
            moPigFunc.AddMultiLineText(strSQL, ")")
            moPigFunc.AddMultiLineText(strSQL, "CREATE INDEX UI_ptKeyValueInf_ExpTime ON dbo._ptKeyValueInf(ExpTime)")
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                strStepName = "ExecuteNonQuery"
                strRet = .ExecuteNonQuery()
                If strRet <> "OK" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(strRet)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
        End Try
    End Function

End Class
