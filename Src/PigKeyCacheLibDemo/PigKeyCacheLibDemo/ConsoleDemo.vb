'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: ConsoleDemo for PigKeyCacheLib
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 3.1.2
'* Create Time: 28/8/2021
'* 1.1	13/11/2021	Add ValueType
'* 1.2	14/11/2021	Modify SavePigKeyValue,GetPigKeyValue
'* 1.3	21/11/2021	Modify SavePigKeyValue,GetPigKeyValue
'* 1.4	1/12/2021	Add TextType,SaveType
'* 1.5	2/12/2021	Modify TextType,SaveType
'* 1.6	5/12/2021	Modify TextType,SaveType
'* 3.0	10/12/2021	Pigkeycachelib version 3.0 is supported, and the following versions of interfaces are no longer supported.
'* 3.1	13/12/2021	Modify GetPigKeyValue,SavePigKeyValue
'**********************************
Imports System.Data
Imports PigKeyCacheLib
Imports PigToolsLiteLib

Public Class ConsoleDemo
    Public PigKeyValueApp As PigKeyValueApp
    Public ShareMemRoot As String = "Test"
    Public CacheWorkDir As String = "C:\Temp"
    Public CacheLevel As PigKeyValueApp.enmCacheLevel = PigKeyValueApp.enmCacheLevel.ToShareMem
    Public KeyName As String = "Key1"
    Public KeyValue As String = "Value1"
    Public ExpTime As DateTime = Now.AddMinutes(10)
    Public ExpTimeSec As Integer = 60
    Public PigFunc As New PigFunc
    Public ValueType As PigKeyValue.enmValueType = PigKeyValue.enmValueType.Text
    Public TextType As PigText.enmTextType = PigText.enmTextType.UTF8
    Public SaveType As PigKeyValue.enmSaveType = PigKeyValue.enmSaveType.Original
    Public Ret As String

    Public Sub Main()
        Dim strLine As String
        Console.WriteLine("*******************")
        Console.WriteLine("Init Setting")
        Console.WriteLine("*******************")
        Console.WriteLine("Input CacheLevel")
        Console.WriteLine(PigKeyValueApp.EnmCacheLevel.ToList & " = ToList (Program for single process multithreading)")
        Console.WriteLine(PigKeyValueApp.EnmCacheLevel.ToShareMem & " = ToShareMem (It is applicable to multi-process and multi-threaded programs under the same user session or IIS application pools.)")
        Console.WriteLine(PigKeyValueApp.EnmCacheLevel.ToFile & " = ToFile (It is suitable for any multi process and multi thread program on the same host.)")
        Console.WriteLine("Now is " & Me.CacheLevel)
        strLine = Console.ReadLine
        If strLine <> "" Then Me.CacheLevel = strLine
        Select Case Me.CacheLevel
            Case PigKeyValueApp.enmCacheLevel.ToList
                Me.PigKeyValueApp = New PigKeyValueApp()
            Case PigKeyValueApp.enmCacheLevel.ToShareMem
                Console.WriteLine("Input ShareMemRoot:" & Me.ShareMemRoot)
                strLine = Console.ReadLine
                If strLine <> "" Then Me.ShareMemRoot = strLine
                Me.PigKeyValueApp = New PigKeyValueApp(Me.ShareMemRoot)
            Case PigKeyValueApp.enmCacheLevel.ToFile
                Console.WriteLine("Input CacheWorkDir:" & Me.CacheWorkDir)
                strLine = Console.ReadLine
                If strLine <> "" Then Me.CacheWorkDir = strLine
                Me.PigKeyValueApp = New PigKeyValueApp(Me.CacheWorkDir, Me.CacheLevel)
            Case Else
                Console.WriteLine("Unsupported CacheLevel")
                Exit Sub
        End Select
        Me.PigKeyValueApp.OpenDebug()
        Me.PigKeyValueApp.PigKeyValues.OpenDebug()
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
            Console.WriteLine("Press G to SavePigKeyValue(Simplified interface Text)")
            Console.WriteLine("Press H to SavePigKeyValue(Simplified interface Bytes)")
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
                            Dim oPigKeyValue As PigKeyValue
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
                                            oPigKeyValue = New PigKeyValue(Me.KeyName, Me.ExpTime, Me.KeyValue, Me.TextType, Me.SaveType)
                                        Case PigKeyValue.EnmValueType.Bytes
                                            Dim oPigText As New PigText(Me.KeyValue, Me.TextType)
                                            oPigKeyValue = New PigKeyValue(Me.KeyName, Me.ExpTime, oPigText.TextBytes, Me.SaveType)
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
                            If oPigKeyValue IsNot Nothing Then
                                With oPigKeyValue
                                    Me.ValueType = .ValueType
                                    Console.WriteLine("KeyName=" & .KeyName)
                                    Console.WriteLine("IsExpired=" & .IsExpired)
                                    Console.WriteLine("ExpTime=" & .ExpTime)
                                    Console.WriteLine("ValueType=" & .ValueType.ToString)
                                    Console.WriteLine("ValueLen=" & .ValueLen.ToString)
                                    Console.WriteLine("SaveType=" & .SaveType.ToString)
                                    Console.WriteLine("ChkMD5Type=" & .ChkMD5Type.ToString)
                                    Console.WriteLine("BodyLen=" & .BodyLen.ToString)
                                    Console.WriteLine("BodyMD5.Length=" & .BodyMD5.Length.ToString)
                                    Console.WriteLine("BodyData.Main.Length=" & .BodyData.Main.Length.ToString)
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
                        Dim oPigKeyValue As PigKeyValue = Me.PigKeyValueApp.GetPigKeyValue(KeyName)
                        Dim strRet As String = .RemovePigKeyValue(oPigKeyValue, Me.CacheLevel)
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
                Case ConsoleKey.G
                    Console.WriteLine("*******************")
                    Console.WriteLine("SavePigKeyValue(Simplified interface Text)")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Input KeyName:" & Me.KeyName)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyName = strLine
                    Console.WriteLine("Input KeyValue:" & Me.KeyValue)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyValue = strLine
                    Console.WriteLine("Input ExpTimeSec:" & Me.ExpTimeSec)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.ExpTimeSec = 60
                    With Me.PigKeyValueApp
                        Console.WriteLine("SavePigKeyValue")
                        .SavePigKeyValue(Me.KeyName, Me.KeyValue, Me.ExpTimeSec)
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                    End With
                Case ConsoleKey.H
                    Console.WriteLine("*******************")
                    Console.WriteLine("SavePigKeyValue(Simplified interface Bytes)")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Input KeyName:" & Me.KeyName)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyName = strLine
                    Console.WriteLine("Input KeyValue:" & Me.KeyValue)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.KeyValue = strLine
                    Console.WriteLine("Input ExpTimeSec:" & Me.ExpTimeSec)
                    strLine = Console.ReadLine
                    If strLine <> "" Then Me.ExpTimeSec = 60
                    With Me.PigKeyValueApp
                        Dim oPigText As New PigText(Me.KeyValue, PigText.enmTextType.UTF8)
                        Console.WriteLine("SavePigKeyValue")
                        .SavePigKeyValue(Me.KeyName, oPigText.TextBytes, Me.ExpTimeSec)
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                    End With
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
