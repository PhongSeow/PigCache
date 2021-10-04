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
    Public PigFunc As New PigFunc
    Public Sub Main()
        Dim strLine As String
        Console.WriteLine("*******************")
        Console.WriteLine("Init Setting")
        Console.WriteLine("*******************")
        Console.WriteLine("Input CacheLevel")
        Console.WriteLine("10 = ToList (Program for single process multithreading)")
        Console.WriteLine("20 = ToShareMem (It is applicable to multi-process and multi-threaded programs under the same user session or IIS application pools.)")
        Console.WriteLine("30 = ToFile (It is suitable for any multi process and multi thread program on the same host.)")
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
                        Dim strRet As String = .RemovePigKeyValue(Me.KeyName, Me.CacheLevel)
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
