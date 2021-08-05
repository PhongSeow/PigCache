'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用|Piggy key value application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.6
'* Create Time: 8/5/2021
'* 1.0.2	13/5/2021 Modify New
'* 1.0.3	22/7/2021 Modify GetPigKeyValue
'* 1.0.4	23/7/2021 remove ObjAdoDBLib
'* 1.0.5	4/8/2021 Remove PigSQLSrvLib
'* 1.0.6	5/8/2021 Modify GetPigKeyValue,SavePigKeyValue
'************************************

Public Class PigKeyValueApp
    Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.6.3"

	Public ReadOnly Property PigKeyValues As New PigKeyValues

	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub

	Public Function GetPigKeyValue(KeyName As String) As PigKeyValue
		Const SUB_NAME As String = "GetKeyValue"
		Dim strStepName As String = ""
		Try
			strStepName = "GetItem"
			GetPigKeyValue = Me.PigKeyValues.Item(KeyName)
			If Me.PigKeyValues.LastErr <> "" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(Me.PigKeyValues.LastErr)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			Me.PrintDebugLog(SUB_NAME, "Catch Exception", Me.LastErr)
			Return Nothing
		End Try
	End Function

	Public Sub SavePigKeyValue(NewItem As PigKeyValue)
		Const SUB_NAME As String = "SavePigKeyValue"
		Dim strStepName As String = ""
		Try
			strStepName = "Add(NewItem)"
			Me.PigKeyValues.Add(NewItem)
			If Me.PigKeyValues.LastErr <> "" Then
				strStepName &= "(" & NewItem.KeyName & ")"
				Throw New Exception(Me.PigKeyValues.LastErr)
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			Me.PrintDebugLog(SUB_NAME, "Catch Exception", Me.LastErr)
		End Try
	End Sub

	Public Sub RemoveExpItems()
		Const SUB_NAME As String = "RemoveExpItems"
		Dim strStepName As String = ""
		Try
			Dim intItems As Integer = 0
			Dim astrKeyName(intItems) As String
			strStepName = "For Each"
			For Each oPigKeyValue As PigKeyValue In Me.PigKeyValues
				Dim strKeyName As String = oPigKeyValue.KeyName
				If oPigKeyValue.IsExpired = True Then
					intItems += 1
					ReDim Preserve astrKeyName(intItems)
					astrKeyName(intItems) = strKeyName
				End If
			Next
			If intItems > 0 Then
				strStepName = "For i"
				For i = 1 To intItems
					Me.PigKeyValues.Remove(astrKeyName(i))
					If Me.PigKeyValues.LastErr <> "" Then
						strStepName = "Remove " & astrKeyName(i)
						Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
					End If
				Next
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Sub


End Class
