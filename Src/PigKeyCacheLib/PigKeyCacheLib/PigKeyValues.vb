'**********************************
'* Name: PigKeyValues
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigKeyValue 的 集合类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 8/5/2021
'* 1.0.2	5/8/2021 Add mAdd,IsItemExists, and modify Add,Remove
'************************************
Imports PigToolsLib

Public Class PigKeyValues
    Inherits PigBaseMini
    Implements IEnumerable(Of PigKeyValue)
    Private Const CLS_VERSION As String = "1.0.2.5"

    Private moList As New List(Of PigKeyValue)

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Try
                Return moList.Count
            Catch ex As Exception
                Me.SetSubErrInf("Count", ex)
                Return -1
            End Try
        End Get
    End Property
    Public Function GetEnumerator() As IEnumerator(Of PigKeyValue) Implements IEnumerable(Of PigKeyValue).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As PigKeyValue
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(KeyName As String) As PigKeyValue
        Get
            Try
                Item = Nothing
                For Each oPigKeyValue As PigKeyValue In moList
                    If oPigKeyValue.KeyName = KeyName Then
                        Item = oPigKeyValue
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.KeyName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(KeyName) As Boolean
        Try
            IsItemExists = False
            For Each oPigKeyValue As PigKeyValue In moList
                If oPigKeyValue.KeyName = KeyName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Sub mAdd(NewItem As PigKeyValue)
        Try
            If Me.IsItemExists(NewItem.KeyName) = True Then Throw New Exception(NewItem.KeyName & " already exists.")
            moList.Add(NewItem)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mAdd", ex)
        End Try
    End Sub

    Public Sub Add(NewItem As PigKeyValue)
        Me.mAdd(NewItem)
    End Sub

    Public Function Add(KeyName As String, ExpTime As DateTime, KeyValue As String) As PigKeyValue
        Dim strStepName As String = ""
        Try
            strStepName = "New PigKeyValue"
            Dim oPigKeyValue As New PigKeyValue(KeyName, ExpTime, KeyValue)
            If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
            strStepName = "Add"
            Me.mAdd(oPigKeyValue)
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            Add = oPigKeyValue
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Add.Text", strStepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Add(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), ValueType As PigKeyValue.enmValueType) As PigKeyValue
        Dim strStepName As String = ""
        Try
            strStepName = "New PigKeyValue"
            Dim oPigKeyValue As New PigKeyValue(KeyName, ExpTime, KeyValue, ValueType)
            If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
            strStepName = "Add"
            Me.mAdd(oPigKeyValue)
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            Add = oPigKeyValue
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Add.ValueType", ex)
            Return Nothing
        End Try
    End Function

    Public Sub Remove(KeyName As String)
        Dim strStepName As String = ""
        Try
            strStepName = "For Each"
            For Each oPigKeyValue As PigKeyValue In moList
                If oPigKeyValue.KeyName = KeyName Then
                    strStepName = "Remove " & KeyName
                    moList.Remove(oPigKeyValue)
                    Exit For
                End If
            Next
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.KeyName", strStepName, ex)
        End Try
    End Sub

    Public Sub Remove(Index As Integer)
        Dim strStepName As String = ""
        Try
            strStepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.Index", strStepName, ex)
        End Try
    End Sub

End Class
