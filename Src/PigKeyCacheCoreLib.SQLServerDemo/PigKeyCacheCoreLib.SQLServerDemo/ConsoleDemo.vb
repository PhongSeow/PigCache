'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: ConsoleDemo for PigKeyCacheLib.SQLServer
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 2.3
'* Create Time: 18/8/2021
'* 1.1	5/10/2021	Modify RemovePigKeyValue
'* 2.0	15/12/2021	Supports PigKeyCacheLib.SQLServer 2.0
'* 2.1	28/12/2021	Supports PigKeyCacheLib.SQLServer 3.0
'* 2.2	26/1/2022	Refer to PigConsole.Getpwdstr of PigCmdLib  is used to hide the entered password.
'* 2.3	1/10/2022	Reference PigSQLSrvLib or PigSQLSrvCoreLib instead
'**********************************

Imports System.Data
Imports PigCmdLib
Imports PigToolsLiteLib
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If


Public Class ConsoleDemo
    Public ConnSQLSrv As ConnSQLSrv
    Public CmdSQLSrvSp As CmdSQLSrvSp
    Public CmdSQLSrvText As CmdSQLSrvText
    Public ConnStr As String
    Public SQL As String
    Public DBSrv As String = "localhost"
    Public DBUser As String = "sa"
    Public DBPwd As String = ""
    Public CurrDB As String = "TestDB"
    Public InpStr As String
    Public SQLSrvKeyValue As SQLSrvKeyValue

    Public CacheWorkDir As String = "C:\Temp"
    Public KeyName As String = "Key1"
    Public KeyValue As String = "Value1"
    Public ExpTime As DateTime = Now.AddMinutes(10)
    Public PigFunc As New PigFunc
    Public TextType As PigText.enmTextType = PigText.enmTextType.UTF8
    Public Ret As String
    Public PigConsole As New PigConsole
    Public CacheTimeSec As Integer
    Public HitCache As PigKeyValue.HitCacheEnum
    Public IsCompress As Boolean
    Public Sub Main()
        Console.WriteLine("*******************")
        Console.WriteLine("Init Setting")
        Console.WriteLine("*******************")
        Console.CursorVisible = True
        Me.PigConsole.GetLine("Input SQL Server", Me.DBSrv)
        If Me.DBSrv = "" Then Me.DBSrv = "localhost"
        Console.WriteLine("SQL Server=" & Me.DBSrv)
        Me.PigConsole.GetLine("Input Default DB", Me.CurrDB)
        If Me.CurrDB = "" Then Me.CurrDB = "TestDB"
        Console.WriteLine("Default DB=" & Me.CurrDB)
        If Me.PigConsole.IsYesOrNo("Is Trusted Connection ?") = True Then
            Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB)
        Else
            Console.WriteLine("Input DB User:" & Me.DBUser)
            Me.PigConsole.GetLine("Input DB User", Me.DBUser)
            If Me.DBUser = "" Then Me.DBUser = "sa"
            Console.WriteLine("DB User=" & Me.DBUser)
            Console.WriteLine("Input DB Password" & Me.DBUser)
            Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB, Me.DBUser, Me.DBPwd)
        End If
        Me.PigConsole.GetLine("Input buffer working directory:" & Me.CacheWorkDir)
        Me.IsCompress = Me.PigConsole.IsYesOrNo("Whether to save after compression?")
        Me.ConnSQLSrv.ConnectionTimeout = 5
        Me.ConnSQLSrv.OpenOrKeepActive()
        Me.SQLSrvKeyValue = New SQLSrvKeyValue(Me.ConnSQLSrv, Me.CacheWorkDir, IsCompress)
        If Me.SQLSrvKeyValue.LastErr <> "" Then Console.WriteLine(Me.SQLSrvKeyValue.LastErr)
        Me.SQLSrvKeyValue.OpenDebug()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to SavePigKeyValue")
            Console.WriteLine("Press B to GetPigKeyValue")
            Console.WriteLine("*******************")
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("SavePigKeyValue")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input KeyName", Me.KeyName)
                    Me.PigConsole.GetLine("Input key value", Me.KeyValue)
                    Dim strDisp As String = Me.PigFunc.GetEnmDispStr(PigText.enmTextType.UTF8, True)
                    strDisp &= Me.PigFunc.GetEnmDispStr(PigText.enmTextType.Ascii)
                    strDisp &= Me.PigFunc.GetEnmDispStr(PigText.enmTextType.Unicode)
                    Me.PigConsole.GetLine("Select TextType" & strDisp, Me.TextType)
                    Console.WriteLine("SaveKeyValue...")
                    Me.Ret = Me.SQLSrvKeyValue.SaveKeyValue(Me.KeyName, Me.KeyValue, Me.TextType)
                    Console.WriteLine(Me.Ret)
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetPigKeyValue")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input KeyName", Me.KeyName)
                    Me.PigConsole.GetLine("Input Seconds cached", Me.CacheTimeSec)
                    Console.WriteLine("GetKeyValue...")
                    Me.Ret = Me.SQLSrvKeyValue.GetKeyValue(Me.KeyName, Me.KeyValue,, Me.CacheTimeSec, Me.HitCache)
                    Console.WriteLine(Me.Ret)
                    Console.WriteLine("KeyValue=" & Me.KeyValue)
                    Console.WriteLine("CacheTimeSec=" & Me.CacheTimeSec)
                    Console.WriteLine("HitCache=" & Me.HitCache.ToString)
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
