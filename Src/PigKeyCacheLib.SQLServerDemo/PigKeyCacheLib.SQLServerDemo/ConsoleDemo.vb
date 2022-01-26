'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: ConsoleDemo for PigKeyCacheLib.SQLServer
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 2.2
'* Create Time: 18/8/2021
'* 1.1	5/10/2021	Modify RemovePigKeyValue
'* 2.0	15/12/2021	Supports PigKeyCacheLib.SQLServer 2.0
'* 2.1	28/12/2021	Supports PigKeyCacheLib.SQLServer 3.0
'* 2.2	26/1/2022	Refer to PigConsole.Getpwdstr of PigCmdLib  is used to hide the entered password.
'**********************************

Imports System.Data
Imports PigCmdLib
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
#If NETFRAMEWORK Then
    Public PigKeyValueApp As PigKeyCacheLib.SQLServer.PigKeyValueApp
#Else
    Public PigKeyValueApp As PigKeyCacheCoreLib.SQLServer.PigKeyValueApp
#End If

    Public ShareMemRoot As String = "Test"
    Public KeyName As String = "Key1"
    Public KeyValue As String = "Value1"
    Public ExpTime As DateTime = Now.AddMinutes(10)
    Public PigFunc As New PigFunc
    Public ValueType As PigKeyValue.EnmValueType = PigKeyValue.enmValueType.Text
    Public TextType As PigText.enmTextType = PigText.enmTextType.UTF8
    Public SaveType As PigKeyValue.enmSaveType = PigKeyValue.enmSaveType.Original
    Public Ret As String
    Public PigConsole As New PigConsole
    Public Sub Main()
        Dim strLine As String
        Console.WriteLine("*******************")
        Console.WriteLine("Init Setting")
        Console.WriteLine("*******************")
        Console.CursorVisible = True
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
                Me.DBPwd = Me.PigConsole.GetPwdStr
                'Console.WriteLine("DB Password=" & Me.DBPwd)
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
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("SavePigKeyValue")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Console.WriteLine("Input KeyName:" & Me.KeyName)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyName = strLine
                    Console.WriteLine("Input ValueType(" & PigKeyValue.EnmValueType.Text & "-Text," & PigKeyValue.EnmValueType.Bytes & "-Bytes):" & Me.ValueType.ToString)
                    strLine = Console.ReadLine
                    Me.ValueType = CInt(strLine)
                    Select Case Me.ValueType
                        Case PigKeyValue.EnmValueType.Bytes, PigKeyValue.EnmValueType.Text
                            Console.WriteLine("Input ExpTime:" & Format(Me.ExpTime, "yyyy-MM-dd HH:mm:ss.fff"))
                            strLine = Console.ReadLine
                            If strLine <> "" Then Me.ExpTime = Me.PigFunc.GECDate(strLine)
                            Console.WriteLine("Input KeyValue:" & Me.KeyValue)
                            strLine = Console.ReadLine
                            If strLine <> "" Then Me.KeyValue = strLine
                            Console.WriteLine("New PigKeyValue")
#If NETFRAMEWORK Then
                            Dim oPigKeyValue As PigKeyCacheLib.SQLServer.PigKeyValue
#Else
                            Dim oPigKeyValue As PigKeyCacheCoreLib.SQLServer.PigKeyValue
#End If
                            oPigKeyValue = Nothing
                            Dim bolIsAdd As Boolean = False
                            Select Case Me.ValueType
                                Case PigKeyValue.EnmValueType.Text
                                    Console.WriteLine("Input TextType(" & PigText.enmTextType.Unicode & "-Unicode," & PigText.enmTextType.UTF8 & "-UTF8," & PigText.enmTextType.Ascii & "-Ascii):" & Me.TextType.ToString)
                                    strLine = Console.ReadLine
                                    Me.TextType = CInt(strLine)
                                    Select Case Me.TextType
                                        Case PigText.enmTextType.Unicode, PigText.enmTextType.UTF8, PigText.enmTextType.Ascii
                                        Case Else
                                            Console.WriteLine("Invalid TextType")
                                    End Select
                                Case PigKeyValue.EnmValueType.Bytes
                                    Me.TextType = PigText.enmTextType.UTF8
                                Case Else
                                    Console.WriteLine("Invalid SaveType")
                            End Select
                            Console.WriteLine("Input SaveType(" & PigKeyValue.EnmSaveType.Original & "-Original," & PigKeyValue.EnmSaveType.SaveSpace & "-SaveSpace," & PigKeyValue.EnmSaveType.EncSaveSpace & "-EncSaveSpace):" & Me.SaveType.ToString)
                            strLine = Console.ReadLine
                            Me.SaveType = CInt(strLine)
                            Select Case Me.SaveType
                                Case PigKeyValue.EnmSaveType.EncSaveSpace, PigKeyValue.EnmSaveType.Original, PigKeyValue.EnmSaveType.SaveSpace
                                    Select Case Me.ValueType
                                        Case PigKeyValue.EnmValueType.Text
#If NETFRAMEWORK Then
                                            oPigKeyValue = New PigKeyCacheLib.SQLServer.PigKeyValue(Me.KeyName, Me.ExpTime, Me.KeyValue, Me.TextType, Me.SaveType)
#Else
                                            oPigKeyValue = New PigKeyCacheCoreLib.SQLServer.PigKeyValue(Me.KeyName, Me.ExpTime, Me.KeyValue, Me.TextType, Me.SaveType)
#End If
                                        Case PigKeyValue.EnmValueType.Bytes
                                            Dim oPigText As New PigText(Me.KeyValue, Me.TextType)
#If NETFRAMEWORK Then
                                            oPigKeyValue = New PigKeyCacheLib.SQLServer.PigKeyValue(Me.KeyName, Me.ExpTime, oPigText.TextBytes, Me.SaveType)
#Else
                                            oPigKeyValue = New PigKeyCacheCoreLib.SQLServer.PigKeyValue(Me.KeyName, Me.ExpTime, oPigText.TextBytes, Me.SaveType)
#End If
                                    End Select
                                    If oPigKeyValue.LastErr <> "" Then
                                        Console.WriteLine(oPigKeyValue.LastErr)
                                    Else
                                        Console.WriteLine("OK")
                                        bolIsAdd = True
                                    End If
                                Case Else
                                    Console.WriteLine("Invalid SaveType")
                            End Select
                            If bolIsAdd = True Then
                                With Me.PigKeyValueApp
                                    Console.WriteLine("SavePigKeyValue")
                                    Me.Ret = .SavePigKeyValue(oPigKeyValue)
                                    Console.WriteLine(Me.Ret)
                                End With
                            End If
                        Case Else
                            Console.WriteLine("Invalid ValueType")
                    End Select
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetPigKeyValue")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Console.WriteLine("Input KeyName:" & Me.KeyName)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyName = strLine
                    With Me.PigKeyValueApp
                        Console.WriteLine("GetPigKeyValue")
#If NETFRAMEWORK Then
                        Dim oPigKeyValue As PigKeyCacheLib.SQLServer.PigKeyValue = .GetPigKeyValue(Me.KeyName)
#Else
                        Dim oPigKeyValue As PigKeyCacheCoreLib.SQLServer.PigKeyValue = .GetPigKeyValue(Me.KeyName)
#End If
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                            If oPigKeyValue IsNot Nothing Then
                                With oPigKeyValue
                                    Me.ValueType = .ValueType
                                    Console.WriteLine("KeyName=" & .KeyName)
                                    Console.WriteLine("IsExpired=" & .IsExpired)
                                    Console.WriteLine("ExpTime=" & .ExpTime)
                                    Console.WriteLine("ValueType=" & .ValueType.ToString)
                                    Console.WriteLine("ValueLen=" & .ValueLen.ToString)
                                    Console.WriteLine("SaveType=" & .SaveType.ToString)
                                    'Console.WriteLine("ChkMD5Type=" & .ChkMD5Type.ToString)
                                    'Console.WriteLine("BodyLen=" & .BodyLen.ToString)
                                    'Console.WriteLine("BodyMD5.Length=" & .BodyMD5.Length.ToString)
                                    'Console.WriteLine("BodyData.Main.Length=" & .BodyData.Main.Length.ToString)
                                    If .ValueType = PigKeyValue.EnmValueType.Text Then
                                        Console.WriteLine("TextType=" & .TextType.ToString)
                                    End If
                                    Select Case Me.ValueType
                                        Case PigKeyValue.EnmValueType.Text
                                            Console.WriteLine("StrValue=" & .StrValue)
                                        Case PigKeyValue.EnmValueType.Bytes
                                            Console.WriteLine("StrValue(Base64)=" & .StrValue)
                                            Dim oPigText As New PigText(.BytesValue, PigText.enmTextType.UTF8)
                                            Console.WriteLine("BytesValue(Text)=" & oPigText.Text)
                                    End Select
                                End With
                            End If
                        End If
                    End With
                Case ConsoleKey.C
                    Console.WriteLine("*******************")
                    Console.WriteLine("Show all KeyValues")
                    Console.WriteLine("*******************")
                    Dim i As Integer = 1
                    For Each oPigKeyValue As PigKeyCacheLib.PigKeyValue In Me.PigKeyValueApp.PigKeyValues
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
                    'Console.WriteLine("*******************")
                    'Console.WriteLine("RemoveExpItems")
                    'Console.WriteLine("*******************")
                    'With Me.PigKeyValueApp
                    '    .RemoveExpItems()
                    '    If .LastErr <> "" Then
                    '        Console.WriteLine(.LastErr)
                    '    Else
                    '        Console.WriteLine("OK")
                    '    End If
                    'End With
                Case ConsoleKey.E
                    'Console.WriteLine("*******************")
                    'Console.WriteLine("RemovePigKeyValue")
                    'Console.WriteLine("*******************")
                    'Console.WriteLine("Input KeyName:" & Me.KeyName)
                    'strLine = Console.ReadLine
                    'If strLine <> "" Then Me.KeyName = strLine
                    'With Me.PigKeyValueApp
                    '    Dim strRet As String = .RemovePigKeyValue(Me.KeyName, PigKeyCacheLib.PigKeyValueApp.enmCacheLevel.ToShareMem)
                    '    If strRet <> "OK" Then
                    '        Console.WriteLine(strRet)
                    '    Else
                    '        Console.WriteLine("OK")
                    '    End If
                    'End With
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
