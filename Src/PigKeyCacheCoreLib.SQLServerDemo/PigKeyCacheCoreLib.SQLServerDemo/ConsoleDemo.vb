'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: ConsoleDemo for PigKeyCacheLib.SQLServer
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 18/8/2021
'* 1.1	5/10/2021	Modify RemovePigKeyValue
'**********************************

Imports System.Data
Imports PigToolsLiteLib
#If NETFRAMEWORK Then
Imports PigKeyCacheLib.SQLServer
Imports PigSQLSrvLib
#Else
Imports PigKeyCacheCoreLib.SQLServer
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
    Public PigKeyValueApp As PigKeyValueApp
    Public ShareMemRoot As String = "Test"
    Public KeyName As String = "Key1"
    Public KeyValue As String = "Value1"
    Public ExpTime As DateTime = Now.AddMinutes(10)
    Public PigFunc As New PigFunc
    Public Sub Main()

        Dim strLine As String
        Console.WriteLine("*******************")
        Console.WriteLine("Init Setting")
        Console.WriteLine("*******************")
        Console.WriteLine("Input SQL Server:" & Me.DBSrv)
        Me.DBSrv = Console.ReadLine()
        If Me.DBSrv = "" Then Me.DBSrv = "localhost"
        Console.WriteLine("SQL Server=" & Me.DBSrv)
        Console.WriteLine("Input Default DB:" & Me.CurrDB)
        Me.CurrDB = Console.ReadLine()
        If Me.CurrDB = "" Then Me.CurrDB = "TestDB"
        Console.WriteLine("Default DB=" & Me.CurrDB)
        Console.WriteLine("Is Trusted Connection ? (Y/n)")
        Me.InpStr = Console.ReadLine()
        Select Case Me.InpStr
            Case "Y", "y", ""
                Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB)
            Case Else
                Console.WriteLine("Input DB User:" & Me.DBUser)
                Me.DBUser = Console.ReadLine()
                If Me.DBUser = "" Then Me.DBUser = "sa"
                Console.WriteLine("DB User=" & Me.DBUser)
                Console.WriteLine("Input DB Password:")
                Me.DBPwd = Console.ReadLine()
                Console.WriteLine("DB Password=" & Me.DBPwd)
                Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB, Me.DBUser, Me.DBPwd)
        End Select
        Me.ConnSQLSrv.ConnectionTimeout = 5
        Me.ConnSQLSrv.OpenOrKeepActive()
#If NETFRAMEWORK Then
        Me.PigKeyValueApp = New PigKeyCacheLib.SQLServer.PigKeyValueApp(Me.ConnSQLSrv)
#Else
        Me.PigKeyValueApp = New PigKeyCacheCoreLib.SQLServer.PigKeyValueApp(Me.ConnSQLSrv)
#End If

        Me.PigKeyValueApp.OpenDebug()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to SavePigKeyValue")
            Console.WriteLine("Press B to GetPigKeyValue")
            Console.WriteLine("Press C to Show all KeyValues")
            Console.WriteLine("Press D to RemoveExpItems")
            Console.WriteLine("Press E to RemovePigKeyValue")
            Console.WriteLine("Press F to GetStatisticsXml")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("SavePigKeyValue")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Input KeyName:" & Me.KeyName)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyName = strLine
                    Console.WriteLine("Input ExpTime:" & Format(Me.ExpTime, "yyyy-MM-dd HH:mm:ss.fff"))
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.ExpTime = Me.PigFunc.GECDate(strLine)
                    Console.WriteLine("Input KeyValue:" & Me.KeyValue)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyValue = strLine
                    Console.WriteLine("New PigKeyValue")
                    Dim oPigKeyValue As New PigKeyValue(Me.KeyName, Me.ExpTime, Me.KeyValue)
                    If oPigKeyValue.LastErr <> "" Then
                        Console.WriteLine(oPigKeyValue.LastErr)
                    Else
                        Console.WriteLine("OK")
                    End If
                    With Me.PigKeyValueApp
                        Console.WriteLine("SavePigKeyValue")
                        .SavePigKeyValue(oPigKeyValue)
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                    End With
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetPigKeyValue")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Input KeyName:" & Me.KeyName)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyName = strLine
                    With Me.PigKeyValueApp
                        Console.WriteLine("GetPigKeyValue")
                        Dim oPigKeyValue As PigKeyValue = .GetPigKeyValue(Me.KeyName)
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                            If Not oPigKeyValue Is Nothing Then
                                With oPigKeyValue
                                    Console.WriteLine("KeyName=" & .KeyName)
                                    Console.WriteLine("IsExpired=" & .IsExpired)
                                    Console.WriteLine("ExpTime=" & .ExpTime)
                                    Console.WriteLine("ValueType=" & .ValueType.ToString)
                                    Console.WriteLine("StrValue=" & .StrValue)
                                End With
                            End If
                        End If
                    End With
                Case ConsoleKey.C
                    Console.WriteLine("*******************")
                    Console.WriteLine("Show all KeyValues")
                    Console.WriteLine("*******************")
                    Dim i As Integer = 1
                    For Each oPigKeyValue As PigKeyValue In Me.PigKeyValueApp.PigKeyValues
                        With oPigKeyValue
                            Console.WriteLine("*********" & i.ToString & "*********")
                            Console.WriteLine("KeyName=" & .KeyName)
                            Console.WriteLine("IsExpired=" & .IsExpired)
                            Console.WriteLine("ExpTime=" & .ExpTime)
                            Console.WriteLine("ValueType=" & .ValueType.ToString)
                            Console.WriteLine("ValueLen=" & Len(.StrValue))
                            i += 1
                        End With
                    Next
                Case ConsoleKey.D
                    Console.WriteLine("*******************")
                    Console.WriteLine("RemoveExpItems")
                    Console.WriteLine("*******************")
                    With Me.PigKeyValueApp
                        .RemoveExpItems()
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                    End With
                Case ConsoleKey.E
                    Console.WriteLine("*******************")
                    Console.WriteLine("RemovePigKeyValue")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Input KeyName:" & Me.KeyName)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyName = strLine
                    With Me.PigKeyValueApp
                        Dim strRet As String = .RemovePigKeyValue(Me.KeyName, PigKeyCacheLib.PigKeyValueApp.enmCacheLevel.ToShareMem)
                        If strRet <> "OK" Then
                            Console.WriteLine(strRet)
                        Else
                            Console.WriteLine("OK")
                        End If
                    End With
                Case ConsoleKey.F
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetStatisticsXml")
                    Console.WriteLine("*******************")
                    Console.WriteLine(Me.PigKeyValueApp.GetStatisticsXml)
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
