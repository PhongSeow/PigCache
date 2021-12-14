'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用 SQL Server 版|Piggy key value application for SQL Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.10
'* Create Time: 31/8/2021
'* 1.1	1/9/2021 Add mCreateTableKeyValueInf,PigBaseMini,OpenDebug,mIsDBObjExists,GetPigKeyValue
'* 1.2	2/9/2021 Modify mNew,IsPigKeyValueExists,SavePigKeyValue,mCreateTableKeyValueInf, and remove mIsDBObjExists.
'* 1.3	4/9/2021 Modify SavePigKeyValue
'* 1.4	16/9/2021 Modify SavePigKeyValue
'* 1.5	17/9/2021 Modify mCreateTableKeyValueInf,SavePigKeyValue,IsPigKeyValueExists
'* 1.6	21/9/2021 Modify GetPigKeyValue,OpenDebug,SavePigKeyValue,mNew
'* 1.7	10/10/2021 Modify New
'* 1.8	4/12/2021 Modify GetPigKeyValue,mCreateTableKeyValueInf,GetPigKeyValue
'* 1.9	6/12/2021 Modify mNew,GetPigKeyValue
'* 1.10	7/12/2021 Modify PigFunc,SavePigKeyValue, add mAddTableCol
'************************************

Imports PigKeyCacheLib
Imports PigToolsLiteLib
Imports System.Data
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
Imports System.Data.SqlClient
#Else
Imports PigSQLSrvCoreLib
Imports Microsoft.Data.SqlClient
#End If

Public Class PigKeyValueApp
    Inherits PigKeyCacheLib.PigKeyValueApp
    Private Const CLS_VERSION As String = "1.10.1"
    Private moConnSQLSrv As ConnSQLSrv
    Private moPigFunc As New PigFunc

    '-----PigBaseMini-----Begin
    Private ReadOnly moPigBaseMini As New PigBaseMini(CLS_VERSION)
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
        MyBase.New(ConnSQLSrv.Connection.ConnectionString)
        Me.mNew(ConnSQLSrv)
    End Sub



    Private Sub mNew(ConnSQLSrv As ConnSQLSrv)
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            strStepName = "Set MyClassName"
            moPigBaseMini.MyClassName = Me.GetType.Name.ToString
            strStepName = "Set moConnSQLSrv"
            moConnSQLSrv = ConnSQLSrv
            strStepName = "New SQLSrvTools"
            Dim oSQLSrvTools As New SQLSrvTools(moConnSQLSrv)
            strStepName = "IsDBObjExists"
            If oSQLSrvTools.IsDBObjExists(SQLSrvTools.enmDBObjType.UserTable, "_ptKeyValueInf") = False Then
                strStepName = "mCreateTableKeyValueInf"
                strRet = mCreateTableKeyValueInf()
                If strRet <> "" Then Throw New Exception(strRet)
            Else
                strStepName = "mTabAddCol(SaveType)"
                strRet = Me.mAddTableCol()
                If strRet <> "OK" Then
                    Me.PrintDebugLog("mNew", strStepName, strRet)
                End If
            End If
            oSQLSrvTools = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mNew", strStepName, ex)
        End Try
    End Sub

    Public Overloads Function GetPigKeyValue(KeyName As String) As PigKeyValue
        Const SUB_NAME As String = "GetPigKeyValue"
        Dim strStepName As String = ""
        Dim strRet As String = ""
        Try
            strStepName = "MyBase.GetPigKeyValue"
            Dim oPigKeyValue As PigKeyCacheLib.PigKeyValue = MyBase.GetPigKeyValue(KeyName)
            If oPigKeyValue Is Nothing Then
                Dim strSQL As String = "SELECT TOP 1 ValueType,ExpTime,KeyValue,ValueMD5,SaveType    ,TextType FROM dbo._ptKeyValueInf WITH(NOLOCK) WHERE KeyName=@KeyName AND ExpTime>GETDATE()"
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
                        Dim suSMHead As StruSMHead
                        ReDim suSMHead.SaveValueMD5(0)
                        strStepName = "Set suSMHead"
                        With suSMHead
                            .ValueType = rsAny.Fields.Item("ValueType").IntValue
                            .SaveType = rsAny.Fields.Item("SaveType").IntValue
                            .TextType = rsAny.Fields.Item("TextType").IntValue
                        End With
                        strStepName = "Check Types"
                        Select Case suSMHead.ValueType
                            Case PigKeyCacheLib.PigKeyValue.enmValueType.Bytes, PigKeyCacheLib.PigKeyValue.enmValueType.Text
                            Case Else
                                Throw New Exception("Invalid ValueType is " & suSMHead.ValueType.ToString)
                        End Select
                        Select Case suSMHead.SaveType
                            Case PigKeyCacheLib.PigKeyValue.enmSaveType.Original, PigKeyCacheLib.PigKeyValue.enmSaveType.SaveSpace, PigKeyCacheLib.PigKeyValue.enmSaveType.EncSaveSpace
                            Case Else
                                Throw New Exception("Invalid SaveType is " & suSMHead.SaveType.ToString)
                        End Select
                        Select Case suSMHead.TextType
                            Case PigText.enmTextType.Ascii, PigText.enmTextType.Unicode, PigText.enmTextType.UnknowOrBin, PigText.enmTextType.UTF8
                            Case Else
                                Throw New Exception("Invalid TextType is " & suSMHead.TextType.ToString)
                        End Select
                        strStepName = "New PigBytes(ValueMD5)"
                        Dim pbMD5 As New PigBytes(rsAny.Fields.Item("ValueMD5").StrValue)
                        If pbMD5.LastErr <> "" Then Throw New Exception(pbMD5.LastErr)
                        strStepName = "New PigBytes(KeyValue)"
                        Dim pbValue As New PigBytes(rsAny.Fields.Item("KeyValue").StrValue)
                        If pbValue.LastErr <> "" Then Throw New Exception(pbValue.LastErr)
                        If pbMD5.IsMatchBytes(pbValue.PigMD5Bytes) = False Then
                            strStepName = "Check saving data PigMD5"
                            Throw New Exception("Mismatch")
                        End If
                        strStepName = "New PigKeyValue"
                        oPigKeyValue = New PigKeyCacheLib.PigKeyValue(KeyName)
                        If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
                        strStepName = "New PigKeyValue"
                        strRet = oPigKeyValue.InitBytesBySave(suSMHead, pbValue.Main)
                        If strRet <> "OK" Then Throw New Exception(strRet)
                        strStepName = "MyBase.SavePigKeyValue"
                        MyBase.SavePigKeyValue(oPigKeyValue, True)
                        If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
                    End If
                    rsAny.Close()
                    rsAny = Nothing
                    oCmdSQLSrvText = Nothing
                End With
            End If
            Return oPigKeyValue
        Catch ex As Exception
            strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
            Me.PrintDebugLog(SUB_NAME, "As Exception", strRet)
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
            moPigFunc.AddMultiLineText(strSQL, "INSERT INTO dbo._ptKeyValueInf(KeyName,ExpTime,KeyValue,KeyHead)VALUES(@KeyName,@ExpTime,@KeyValue,@KeyHead)", 1)
            moPigFunc.AddMultiLineText(strSQL, "ELSE")
            moPigFunc.AddMultiLineText(strSQL, "UPDATE dbo._ptKeyValueInf SET KeyHead=@KeyHead,ExpTime=@ExpTime,KeyValue=@KeyValue", 1)
            moPigFunc.AddMultiLineText(strSQL, "WHERE KeyName=@KeyName", 1)
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@KeyName", SqlDbType.VarChar, 128)
                .AddPara("@ExpTime", SqlDbType.DateTime)
                .AddPara("@KeyValue", SqlDbType.VarChar, -1)
                .AddPara("@KeyHead", SqlDbType.VarChar, 128)
                .ParaValue("@KeyName") = strKeyName
                .ParaValue("@ExpTime") = NewItem.ExpTime
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


    'Private Function mTabAddCol(ColName As String) As String
    '    Const SUB_NAME As String = "mCreateTableKeyValueInf"
    '    Dim strStepName As String = "", strRet As String
    '    Try
    '        Dim strTabName As String = ""
    '        Dim strSQL As String = ""
    '        strSQL = "ALTER TABLE dbo._ptKeyValueInf ADD "
    '        moPigFunc.AddMultiLineText(strSQL, "ALTER TABLE dbo._ptKeyValueInf")
    '        Select Case ColName
    '            Case "SaveType"
    '                strSQL &= ColName & " int DEFAULT(" & PigKeyValue.enmSaveType.Original & ")"
    '            Case "TextType"
    '                strSQL &= ColName & " int DEFAULT(" & PigText.enmTextType.UTF8 & ")"
    '            Case Else
    '                Throw New Exception("Invalid column name " & ColName)
    '        End Select
    '        strStepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.moConnSQLSrv.Connection
    '            strStepName = "ExecuteNonQuery"
    '            strRet = .ExecuteNonQuery()
    '            If strRet <> "OK" Then
    '                Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
    '                Throw New Exception(strRet)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
    '    End Try
    'End Function

    Private Function mCreateTableKeyValueInf() As String
        Const SUB_NAME As String = "mCreateTableKeyValueInf"
        Dim strStepName As String = "", strRet As String = ""
        Try
            Dim strTabName As String = ""
            Dim strSQL As String = ""
            moPigFunc.AddMultiLineText(strSQL, "CREATE TABLE dbo._ptKeyValueInf(")
            moPigFunc.AddMultiLineText(strSQL, "KeyName varchar(128) NOT NULL,", 1)
            moPigFunc.AddMultiLineText(strSQL, "ExpTime datetime NOT NULL,", 1)
            moPigFunc.AddMultiLineText(strSQL, "KeyHead varchar(256)NOT NULL DEFAULT (''),", 1)
            moPigFunc.AddMultiLineText(strSQL, "KeyValue varchar(max)NOT NULL DEFAULT (''),", 1)
            moPigFunc.AddMultiLineText(strSQL, "CreateTime datetime NOT NULL DEFAULT(GetDate()),", 1)
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

    Private Function mAddTableCol() As String
        Const SUB_NAME As String = "mAddTableCol"
        Dim strStepName As String = "", strRet As String = ""
        Try
            Dim strTabName As String = ""
            Dim strSQL As String = ""
            moPigFunc.AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT 1 FROM syscolumns c JOIN sysobjects o ON c.id=o.id AND o.xtype='U' AND o.uid=1 WHERE o.name='_ptKeyValueInf' AND c.name='KeyHead')")
            moPigFunc.AddMultiLineText(strSQL, "BEGIN")
            moPigFunc.AddMultiLineText(strSQL, "ALTER TABLE dbo._ptKeyValueInf ADD KeyHead varchar(256) NOT NULL DEFAULT ('')", 1)
            moPigFunc.AddMultiLineText(strSQL, "END")
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
