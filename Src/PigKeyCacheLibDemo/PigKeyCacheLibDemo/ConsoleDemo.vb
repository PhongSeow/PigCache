'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: ConsoleDemo for PigKeyCacheLib
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 3.2.2
'* Create Time: 28/8/2021
'* 1.1	13/11/2021	Add ValueType
'* 1.2	14/11/2021	Modify SavePigKeyValue,GetPigKeyValue
'* 1.3	21/11/2021	Modify SavePigKeyValue,GetPigKeyValue
'* 1.4	1/12/2021	Add TextType,SaveType
'* 1.5	2/12/2021	Modify TextType,SaveType
'* 1.6	5/12/2021	Modify TextType,SaveType
'* 3.0	10/12/2021	Pigkeycachelib version 3.0 is supported, and the following versions of interfaces are no longer supported.
'* 3.1	13/12/2021	Modify GetPigKeyValue,SavePigKeyValue
'* 3.2	2/1/2022	Modify New 
'* 3.3	18/9/2022	Modify ConsoleDemo
'**********************************
Imports System.Data
Imports PigToolsLiteLib
Imports PigCmdLib

Public Class ConsoleDemo
    Public PigKeyValue As PigKeyValue
    Public CacheWorkDir As String = "C:\Temp"
    Public KeyName As String = "Key1"
    Public KeyValue As String = "Value1"
    Public PigFunc As New PigFunc
    Public TextType As PigText.enmTextType = PigText.enmTextType.UTF8
    Public CacheTimeSec As Integer
    Public HitCache As PigKeyValue.HitCacheEnum
    Public PigConsole As New PigConsole
    Public Ret As String
    Public IsCompress As Boolean

    Public Sub Main()
        Console.WriteLine("*******************")
        Console.WriteLine("Init Setting")
        Console.WriteLine("*******************")
        Me.PigConsole.GetLine("Input buffer working directory:" & Me.CacheWorkDir)
        Me.IsCompress = Me.PigConsole.IsYesOrNo("Whether to save after compression?")
        Me.PigKeyValue = New PigKeyValue(Me.CacheWorkDir, Me.IsCompress)
        If Me.PigKeyValue.LastErr <> "" Then
            Console.WriteLine(Me.PigKeyValue.LastErr)
            Exit Sub
        End If
        Me.PigKeyValue.OpenDebug()
        Console.WriteLine("")
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to SavePigKeyValue")
            Console.WriteLine("Press B to GetPigKeyValue")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
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
                    Me.Ret = Me.PigKeyValue.SaveKeyValue(Me.KeyName, Me.KeyValue, Me.TextType)
                    Console.WriteLine(Me.Ret)
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetPigKeyValue")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input KeyName", Me.KeyName)
                    Me.PigConsole.GetLine("Input Seconds cached", Me.CacheTimeSec)
                    Console.WriteLine("GetKeyValue...")
                    Me.Ret = Me.PigKeyValue.GetKeyValue(Me.KeyName, Me.KeyValue,, Me.CacheTimeSec, Me.HitCache)
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
