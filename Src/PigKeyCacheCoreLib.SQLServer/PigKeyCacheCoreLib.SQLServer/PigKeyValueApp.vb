'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用 SQL Server 版|Piggy key value application for SQL Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 3.3
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
'* 2.0	15/12/2021 Modify SavePigKeyValue,GetPigKeyValue
'* 2.1	17/12/2021 Modify GetPigKeyValue, add Shadows
'* 3.0	28/12/2021 Code rewriting
'* 3.1	29/12/2021 Modify mNew,IsPigKeyValueExists,GetPigKeyValue
'* 3.2	31/12/2021 Modify mAddTableCol
'* 3.3	31/12/2021 Modify mAddTableCol
'* 3.5	26/7/2022 Modify Imports and Obj,GetPigKeyValue
'************************************

Imports System.Data
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
Imports System.Data.SqlClient
Imports PigKeyCacheFwkLib
Imports PigToolsWinLib
#Else
Imports PigSQLSrvCoreLib
Imports Microsoft.Data.SqlClient
Imports PigKeyCacheLib
Imports PigToolsLiteLib
#End If

Public Class PigKeyValueApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "3.5.8"
#If NETFRAMEWORK Then
    Friend Property Obj As PigKeyCacheFwkLib.PigKeyValueApp
#Else
    Friend Property Obj As PigKeyCacheLib.PigKeyValueApp
#End If
    Private moConnSQLSrv As ConnSQLSrv
    Private moPigFunc As New PigFunc


    Public Overloads Function OpenDebug() As String
        Dim LOG As New PigStepLog("OpenDebug")
        Try
            LOG.StepName = "moConnSQLSrv.OpenDebug"
            moConnSQLSrv.OpenDebug()
            If moConnSQLSrv.LastErr <> "" Then Throw New Exception(moConnSQLSrv.LastErr)
            LOG.StepName = "Obj.OpenDebug"
            Me.Obj.OpenDebug()
            If Me.Obj.LastErr <> "" Then Throw New Exception(Me.Obj.LastErr)
            LOG.StepName = "Obj.PigKeyValues.OpenDebug"
            Me.Obj.PigKeyValues.OpenDebug()
            If Me.Obj.PigKeyValues.LastErr <> "" Then Throw New Exception(Me.Obj.PigKeyValues.LastErr)
            LOG.StepName = "moPigBaseMini.OpenDebug"
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Sub New(ConnSQLSrv As ConnSQLSrv)
        MyBase.New(ConnSQLSrv.Connection.ConnectionString)
        Me.mNew(ConnSQLSrv)
    End Sub



    Private Sub mNew(ConnSQLSrv As ConnSQLSrv)
        Dim LOG As New PigStepLog("mNew")
        Try
            LOG.StepName = "Set moConnSQLSrv"
            moConnSQLSrv = ConnSQLSrv
            Dim strShareMemRoot As String = ""
            With moConnSQLSrv
                strShareMemRoot = "<" & .PrincipalSQLServer & ">"
                If .IsTrustedConnection = False Then
                    strShareMemRoot &= "<" & .DBUser & ">"
                End If
                strShareMemRoot &= "<" & .CurrDatabase & ">"
            End With
            LOG.StepName = "New PigKeyCacheLib.PigKeyValueApp"
            If Me.IsWindows = True Then
#If NETFRAMEWORK Then
                Me.Obj = New PigKeyCacheFwkLib.PigKeyValueApp(strShareMemRoot)
#Else
                Me.Obj = New PigKeyCacheLib.PigKeyValueApp(strShareMemRoot)
#End If
            Else
#If NETFRAMEWORK Then
                Me.Obj = New PigKeyCacheFwkLib.PigKeyValueApp()
#Else
                Me.Obj = New PigKeyCacheLib.PigKeyValueApp()
#End If
            End If
            If Me.Obj.LastErr <> "" Then Throw New Exception(Me.Obj.LastErr)
            LOG.StepName = "New SQLSrvTools"
            Dim oSQLSrvTools As New SQLSrvTools(moConnSQLSrv)
            LOG.StepName = "IsDBObjExists"
            If oSQLSrvTools.IsDBObjExists(SQLSrvTools.enmDBObjType.UserTable, "_ptKeyValueInf") = False Then
                LOG.StepName = "mCreateTableKeyValueInf"
                LOG.Ret = mCreateTableKeyValueInf()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Else
                LOG.StepName = "mTabAddCol(SaveType)"
                LOG.Ret = Me.mAddTableCol()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog("mNew", LOG.StepName, LOG.Ret)
                End If
            End If
            oSQLSrvTools = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Public Overloads Function GetPigKeyValue(KeyName As String) As PigKeyValue
        Dim LOG As New PigStepLog("GetPigKeyValue")
        Try
            If Me.Obj Is Nothing Then Throw New Exception("Obj not instantiated")
            LOG.StepName = "GetPigKeyValue"
#If NETFRAMEWORK Then
            Dim oPigKeyValue As PigKeyCacheFwkLib.PigKeyValue = Me.Obj.GetPigKeyValue(KeyName)
#Else
            Dim oPigKeyValue As PigKeyCacheLib.PigKeyValue = Me.Obj.GetPigKeyValue(KeyName)
#End If
            If oPigKeyValue Is Nothing Then
                Dim strSQL As String = "SELECT TOP 1 ExpTime,HeadData,BodyData FROM dbo._ptKeyValueInf WITH(NOLOCK) WHERE KeyName=@KeyName AND ExpTime>GETDATE()"
                LOG.StepName = "New CmdSQLSrvText"
                Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
                With oCmdSQLSrvText
                    .ActiveConnection = Me.moConnSQLSrv.Connection
                    .AddPara("@KeyName", SqlDbType.VarChar, 128)
                    .ParaValue("@KeyName") = KeyName
                    LOG.StepName = "Execute"
                    Dim rsAny As Recordset = .Execute()
                    If .LastErr <> "" Then
                        LOG.AddStepNameInf(KeyName)
                        LOG.AddStepNameInf(.DebugStr)
                        Throw New Exception(.LastErr)
                    End If
                    If rsAny.EOF = False Then
                        Dim strHeadData As String = rsAny.Fields.Item("HeadData").StrValue
                        LOG.StepName = "New PigBytes(HeadData)"
                        Dim pbHead As New PigBytes(strHeadData)
                        If pbHead.LastErr <> "" Then
                            LOG.AddStepNameInf(KeyName)
                            LOG.AddStepNameInf(strHeadData)
                            Throw New Exception(.LastErr)
                        End If
                        LOG.StepName = "New PigKeyValue"
#If NETFRAMEWORK Then
                        oPigKeyValue = New PigKeyCacheFwkLib.PigKeyValue(KeyName)
#Else
                        oPigKeyValue = New PigKeyCacheLib.PigKeyValue(KeyName)
#End If
                        If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
                        LOG.StepName = "LoadHead"
                        LOG.Ret = oPigKeyValue.LoadHead(pbHead)
                        If LOG.Ret <> "OK" Then
                            LOG.AddStepNameInf(KeyName)
                            Throw New Exception(.LastErr)
                        End If
                        Dim strBodyData As String = rsAny.Fields.Item("BodyData").StrValue
                        LOG.StepName = "New PigBytes(BodyData)"
                        Dim pbBody As New PigBytes(strBodyData)
                        If pbBody.LastErr <> "" Then
                            LOG.AddStepNameInf(KeyName)
                            LOG.AddStepNameInf(strBodyData.Length)
                            Throw New Exception(.LastErr)
                        End If
                        LOG.StepName = "LoadBody"
                        LOG.Ret = oPigKeyValue.LoadBody(pbBody.Main)
                        If LOG.Ret <> "OK" Then
                            LOG.AddStepNameInf(KeyName)
                            Throw New Exception(.LastErr)
                        End If
                        LOG.StepName = "MyBase.SavePigKeyValue"
                        Me.Obj.SavePigKeyValue(oPigKeyValue, True)
                        If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
                    End If
                    rsAny.Close()
                    rsAny = Nothing
                    oCmdSQLSrvText = Nothing
                End With
            End If
            With oPigKeyValue
                If .ValueType = PigKeyValue.EnmValueType.Text Then
                    LOG.StepName = "New PigKeyValue(Text)"
                    GetPigKeyValue = New PigKeyValue(KeyName, .ExpTime, .StrValue, .TextType, .SaveType)
                Else
                    LOG.StepName = "New PigKeyValue(Bytes)"
                    GetPigKeyValue = New PigKeyValue(KeyName, .ExpTime, .BytesValue, .SaveType)
                End If
            End With
            If GetPigKeyValue.LastErr <> "" Then
                LOG.AddStepNameInf(KeyName)
                Throw New Exception(GetPigKeyValue.LastErr)
            End If
            oPigKeyValue = Nothing
        Catch ex As Exception
            LOG.Ret = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex, True)
            Me.PrintDebugLog(LOG.SubName, LOG.Ret)
            Return Nothing
        End Try
    End Function

    Public Overloads Function IsPigKeyValueExists(KeyName As String) As Boolean
        Dim LOG As New PigStepLog("IsPigKeyValueExists")
        Try
            Dim strSQL As String = "SELECT TOP 1 1 FROM dbo._ptKeyValueInf WITH(NOLOCK) WHERE KeyName=@KeyName AND ExpTime>GETDATE()"
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            If oCmdSQLSrvText.LastErr <> "" Then
                LOG.AddStepNameInf(KeyName)
                Throw New Exception(oCmdSQLSrvText.LastErr)
            End If
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@KeyName", SqlDbType.VarChar, 128)
                .ParaValue("@KeyName") = KeyName
                LOG.StepName = "Execute"
                Dim rsAny As Recordset = .Execute()
                If .LastErr <> "" Then
                    LOG.AddStepNameInf(KeyName)
                    Me.PrintDebugLog(LOG.SubName, LOG.StepName, .DebugStr)
                    Throw New Exception(.LastErr)
                End If
                LOG.StepName = "Execute"
                If rsAny.EOF = False Then
                    IsPigKeyValueExists = True
                Else
                    IsPigKeyValueExists = False
                End If
                LOG.StepName = "Close"
                rsAny.Close()
                rsAny = Nothing
            End With
            oCmdSQLSrvText = Nothing
        Catch ex As Exception
            Me.PrintDebugLog(LOG.SubName, LOG.StepName, ex.Message.ToString)
            Return False
        End Try
    End Function



    Public Overloads Function SavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True) As String
        Dim LOG As New PigStepLog("SavePigKeyValue")
        Try
            Dim strKeyName As String = NewItem.KeyName
            LOG.StepName = "Check NewItem"
            LOG.Ret = NewItem.fCheck
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strKeyName)
                Throw New Exception(LOG.Ret)
            End If
            If IsOverwrite = False Then
                If Me.IsPigKeyValueExists(strKeyName) = True Then
                    LOG.AddStepNameInf(strKeyName)
                    Throw New Exception("PigKeyValue Exists")
                End If
            End If
            Dim strSQL As String = ""
            moPigFunc.AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT TOP 1 1 FROM dbo._ptKeyValueInf WHERE KeyName=@KeyName)")
            moPigFunc.AddMultiLineText(strSQL, "INSERT INTO dbo._ptKeyValueInf(KeyName,ExpTime,HeadData,BodyData)VALUES(@KeyName,@ExpTime,@HeadData,@BodyData)", 1)
            moPigFunc.AddMultiLineText(strSQL, "ELSE")
            moPigFunc.AddMultiLineText(strSQL, "UPDATE dbo._ptKeyValueInf SET HeadData=@HeadData,ExpTime=@ExpTime,BodyData=@BodyData", 1)
            moPigFunc.AddMultiLineText(strSQL, "WHERE KeyName=@KeyName", 1)
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.moConnSQLSrv.Connection
                .AddPara("@KeyName", SqlDbType.VarChar, 128)
                .AddPara("@ExpTime", SqlDbType.DateTime)
                .AddPara("@HeadData", SqlDbType.VarChar, 256)
                .AddPara("@BodyData", SqlDbType.VarChar, -1)
                .ParaValue("@KeyName") = NewItem.KeyName
                .ParaValue("@ExpTime") = NewItem.ExpTime
                .ParaValue("@HeadData") = NewItem.fHeadData.Base64Str
                .ParaValue("@BodyData") = NewItem.fBodyData.Base64Str
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepName, .DebugStr)
                    Throw New Exception(LOG.Ret)
                ElseIf .RecordsAffected <= 0 Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepName, .DebugStr)
                    Throw New Exception("RecordsAffected=" & .RecordsAffected)
                End If
            End With
            LOG.StepName = "Obj.SavePigKeyValue"
            LOG.Ret = Me.Obj.SavePigKeyValue(NewItem.Obj, IsOverwrite)
            If LOG.Ret <> "OK" Then Throw New Exception(MyBase.LastErr)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


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
            moPigFunc.AddMultiLineText(strSQL, "HeadData varchar(256)NOT NULL DEFAULT (''),", 1)
            moPigFunc.AddMultiLineText(strSQL, "BodyData varchar(max)NOT NULL DEFAULT (''),", 1)
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
            moPigFunc.AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT 1 FROM syscolumns c JOIN sysobjects o ON c.id=o.id AND o.xtype='U' AND o.uid=1 WHERE o.name='_ptKeyValueInf' AND c.name='HeadData')")
            moPigFunc.AddMultiLineText(strSQL, "BEGIN")
            moPigFunc.AddMultiLineText(strSQL, "ALTER TABLE dbo._ptKeyValueInf ADD HeadData varchar(256) NOT NULL DEFAULT ('')", 1)
            moPigFunc.AddMultiLineText(strSQL, "END")
            moPigFunc.AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT 1 FROM syscolumns c JOIN sysobjects o ON c.id=o.id AND o.xtype='U' AND o.uid=1 WHERE o.name='_ptKeyValueInf' AND c.name='BodyData')")
            moPigFunc.AddMultiLineText(strSQL, "BEGIN")
            moPigFunc.AddMultiLineText(strSQL, "ALTER TABLE dbo._ptKeyValueInf ADD BodyData varchar(max) NOT NULL DEFAULT ('')", 1)
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

    Public Function GetStatisticsXml() As String
        Try
            Return Me.Obj.GetStatisticsXml
        Catch ex As Exception
            Me.PrintDebugLog("GetStatisticsXml", ex.Message.ToString)
            Return ""
        End Try
    End Function

    Public Function PigKeyValues() As PigKeyValues
        Try
            Return Me.Obj.PigKeyValues
        Catch ex As Exception
            Me.PrintDebugLog("GetStatisticsXml", ex.Message.ToString)
            Return Nothing
        End Try
    End Function

End Class
