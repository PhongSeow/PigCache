Imports System.Data
Imports PigKeyCacheLib
Imports PigToolsLib

Public Class ConsoleDemo
    Public PigKeyValueApp As New PigKeyValueApp
    Public KeyName As String = "Key1"
    Public KeyValue As String = "Value1"
    Public ExpTime As DateTime = Now.AddMinutes(10)
    Public PigFunc As New PigFunc
    Public Sub Main()
        Me.PigKeyValueApp.OpenDebug()
        Me.PigKeyValueApp.PigKeyValues.OpenDebug()
        Dim strLine As String
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to SavePigKeyValue")
            Console.WriteLine("Press B to GetPigKeyValue")
            Console.WriteLine("Press C to Show all KeyValues")
            Console.WriteLine("Press D to RemoveExpItems")
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
                            With oPigKeyValue
                                Console.WriteLine("KeyName=" & .KeyName)
                                Console.WriteLine("IsExpired=" & .IsExpired)
                                Console.WriteLine("ExpTime=" & .ExpTime)
                                Console.WriteLine("ValueType=" & .ValueType.ToString)
                                Console.WriteLine("StrValue=" & .StrValue)
                            End With
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

            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
