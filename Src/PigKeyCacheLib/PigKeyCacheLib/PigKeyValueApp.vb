'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用|Piggy key value application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.12
'* Create Time: 8/5/2021
'* 1.0.2	13/5/2021 Modify New
'* 1.0.3	22/7/2021 Modify GetPigKeyValue
'* 1.0.4	23/7/2021 remove ObjAdoDBLib
'* 1.0.5	4/8/2021 Remove PigSQLSrvLib
'* 1.0.6	5/8/2021 Modify GetPigKeyValue,SavePigKeyValue
'* 1.0.7	7/8/2021 Modify New and add IsUseMemCache
'* 1.0.8	11/8/2021 Add mSavePigKeyValue2SM,mGetBytesSMBody
'* 1.0.9	13/8/2021 Modify mGetBytesSMBody
'* 1.0.10	13/8/2021 Modify mSaveSMHead,IsUseMemCache
'* 1.0.11	16/8/2021 Modify mSavePigKeyValue2SM,mSavePigKeyValue2SM,ShareMemRoot,GetPigKeyValue,mGetStruSMHead,mGetBytesSMBody, and add mSaveSMBody
'* 1.0.12	17/8/2021 Add PrintDebugLog,IsPigKeyValueExists,RemovePigKeyValue and modify GetPigKeyValue,SavePigKeyValue
'************************************

Imports PigToolsLib

Public Class PigKeyValueApp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.12.16"

	Public ReadOnly Property PigKeyValues As New PigKeyValues
	Public Property IsUseMemCache As Boolean = False
	Friend Property ShareMemRoot As String = ""

	''' <summary>
	''' 共享内存头结构
	''' </summary>
	Private Structure StruSMHead
		Dim ValueType As PigKeyValue.enmValueType
		Dim ExpTime As DateTime
		Dim ValueLen As Long
		Dim ValueMD5 As Byte()
	End Structure

	Public Sub New()
		MyBase.New(CLS_VERSION)
		mNew("")
	End Sub

	Private Sub mNew(Optional ShareMemRoot As String = "")
		Try
			If ShareMemRoot = "" Then ShareMemRoot = Me.AppTitle
			Dim oPigMD5 As New PigMD5(ShareMemRoot, PigMD5.enmTextType.UTF8)
			Me.ShareMemRoot = oPigMD5.PigMD5()
			Me.IsUseMemCache = True
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mNew", ex)
		End Try
	End Sub
	Public Sub New(ShareMemRoot As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(ShareMemRoot)
	End Sub

	Public Function IsPigKeyValueExists(KeyName As String) As Boolean
		Dim strStepName As String = ""
		Try
			If Me.PigKeyValues.IsItemExists(KeyName) = True Then
				Return True
			ElseIf Me.IsUseMemCache = False Then
				Return False
			Else
				Dim strRet As String
				strStepName = "New PigKeyValue"
				Dim pkvNew As New PigKeyValue(KeyName, Now.AddMinutes(1), "")
				pkvNew.Parent = Me
				Dim suSMHead As StruSMHead
				ReDim suSMHead.ValueMD5(0)
				strStepName = "mGetStruSMHead"
				strRet = Me.mGetStruSMHead(suSMHead, pkvNew.SMNameHead)
				If strRet <> "OK" Then
					Return False
				Else
					Dim abBody As Byte()
					ReDim abBody(0)
					strStepName = "mGetBytesSMBody"
					strRet = Me.mGetBytesSMBody(abBody, suSMHead, pkvNew.SMNameBody)
					If strRet <> "OK" Then
						Return False
					Else
						Return True
					End If
				End If
			End If
		Catch ex As Exception
			Me.SetSubErrInf("IsPigKeyValueExists", ex)
			Return False
		End Try
	End Function

	Public Function GetPigKeyValue(KeyName As String) As PigKeyValue
		Const SUB_NAME As String = "GetKeyValue"
		Dim strStepName As String = ""
		Dim bolIsNotLog As Boolean = False
		Try
			strStepName = "GetItem"
			Me.PigKeyValues.ClearErr()
			GetPigKeyValue = Me.PigKeyValues.Item(KeyName)
			If Me.PigKeyValues.LastErr <> "" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(Me.PigKeyValues.LastErr)
			End If
			If GetPigKeyValue Is Nothing Then
				Dim strRet As String
				If Me.IsUseMemCache = True Then
					strStepName = "New PigKeyValue"
					Dim pkvNew As New PigKeyValue(KeyName, Now.AddMinutes(1), "")
					pkvNew.Parent = Me
					Dim suSMHead As StruSMHead
					ReDim suSMHead.ValueMD5(0)
					strStepName = "mGetStruSMHead"
					strRet = Me.mGetStruSMHead(suSMHead, pkvNew.SMNameHead)
					If strRet <> "OK" Then
						strStepName &= strStepName & "(" & KeyName & "." & pkvNew.SMNameHead & ")"
						bolIsNotLog = True
						Throw New Exception(strRet)
					End If
					Dim abBody As Byte()
					ReDim abBody(0)
					strStepName = "mGetBytesSMBody"
					strRet = Me.mGetBytesSMBody(abBody, suSMHead, pkvNew.SMNameBody)
					If strRet <> "OK" Then
						strStepName &= strStepName & "(" & KeyName & "." & pkvNew.SMNameBody & ")"
						Throw New Exception(strRet)
					End If
					pkvNew = Nothing
					pkvNew = New PigKeyValue(KeyName, suSMHead.ExpTime, abBody, suSMHead.ValueType, suSMHead.ValueMD5)
					strStepName = "Add(pkvNew)"
					Me.PigKeyValues.Add(pkvNew)
					If Me.PigKeyValues.LastErr <> "" Then
						strStepName &= "(" & pkvNew.KeyName & ")"
						Throw New Exception(Me.PigKeyValues.LastErr)
					End If
					Return pkvNew
				End If
			End If
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			If bolIsNotLog = False Then Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", Me.LastErr)
			Return Nothing
		End Try
	End Function

	Private Function mGetBytesSMBody(ByRef BodyBytes As Byte(), SuSMHead As StruSMHead, SMNameBody As String) As String
		Const SUB_NAME As String = "mGetBytesSMBody"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim msmBody As New ShareMem
			strStepName = "Body.Init"
			strRet = msmBody.Init(SMNameBody, SuSMHead.ValueLen)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			ReDim BodyBytes(0)
			strStepName = "Body.Read"
			strRet = msmBody.Read(BodyBytes)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "Body New PigBytes"
			Dim pbBody As New PigBytes(BodyBytes)
			If pbBody.LastErr <> "" Then
				strStepName &= "(abBody.Length=" & BodyBytes.Length & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "Check Value"
			If SuSMHead.ValueLen <> BodyBytes.Length Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception("ValueLen not match " & SuSMHead.ValueLen & "," & BodyBytes.Length)
			End If
			Dim oPigMD5 As New PigMD5(BodyBytes)
			If oPigMD5.LastErr <> "" Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception(oPigMD5.LastErr)
			End If
			If pbBody.PigMD5Bytes.SequenceEqual(SuSMHead.ValueMD5) = False Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception("PigMD5 not match")
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			'Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", strRet)
			Return strRet
		End Try
	End Function

	Private Function mGetStruSMHead(ByRef SuSMHead As StruSMHead, SMNameHead As String) As String
		Const SUB_NAME As String = "mGetStruSMHead"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim msmHead As New ShareMem
			strStepName = "Head.Init"
			strRet = msmHead.Init(SMNameHead, 36)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameHead & ")"
				Throw New Exception(strRet)
			End If
			Dim abHead(0) As Byte
			strStepName = "Head.Read"
			strRet = msmHead.Read(abHead)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameHead & ")"
				Throw New Exception(strRet)
			End If
			msmHead = Nothing
			strStepName = "Head New PigBytes"
			Dim pbHead As New PigBytes(abHead)
			If pbHead.LastErr <> "" Then
				strStepName &= "(abHead.Length=" & abHead.Length & ")"
				Throw New Exception(strRet)
			End If
			ReDim SuSMHead.ValueMD5(15)
			With pbHead
				SuSMHead.ValueType = .GetInt32Value()
				SuSMHead.ExpTime = .GetDateTimeValue
				SuSMHead.ValueLen = .GetInt64Value
				SuSMHead.ValueMD5 = .GetBytesValue(16)
			End With
			strStepName = "Check StruSMHead (" & SMNameHead & ")"
			With SuSMHead
				If .ValueLen = 0 Then Throw New Exception("ValueLen is 0")
				If .ExpTime < Now Then Throw New Exception("ExpTime")
				Select Case .ValueType
					Case PigKeyValue.enmValueType.Bytes, PigKeyValue.enmValueType.EncBytes, PigKeyValue.enmValueType.Text, PigKeyValue.enmValueType.ZipBytes, PigKeyValue.enmValueType.ZipEncBytes
					Case Else
						Throw New Exception("invalid ValueType " & .ValueType)
				End Select
			End With
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			'Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", strRet)
			Return strRet
		End Try
	End Function

	Private Function mSaveSMHead(SuSMHead As StruSMHead, SMNameHead As String) As String
		Const SUB_NAME As String = "mSaveSMHead"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim msmHead As New ShareMem
			strStepName = "Head.Init"
			strRet = msmHead.Init(SMNameHead, 36)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameHead & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "Head New PigBytes"
			Dim pbHead As New PigBytes
			If pbHead.LastErr <> "" Then
				Throw New Exception(strRet)
			End If
			strStepName = "SetValue"
			With pbHead
				.SetValue(SuSMHead.ValueType)
				If .LastErr <> "" Then
					strStepName &= ".ValueType"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.ExpTime)
				If .LastErr <> "" Then
					strStepName &= ".ExpTime"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.ValueLen)
				If .LastErr <> "" Then
					strStepName &= ".ValueLen"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.ValueMD5)
				If .LastErr <> "" Then
					strStepName &= ".ValueMD5"
					Throw New Exception(.LastErr)
				End If
			End With
			strStepName = "Head.Write"
			strRet = msmHead.Write(pbHead.Main)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameHead & ")"
				Throw New Exception(strRet)
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			'Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", strRet)
			Return strRet
		End Try
	End Function

	Private Function mSaveSMBody(SuSMHead As StruSMHead, SMNameBody As String, ByRef DataBytes As Byte()) As String
		Const SUB_NAME As String = "mSaveSMBody"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim msmBody As New ShareMem
			strStepName = "Body.Init"
			strRet = msmBody.Init(SMNameBody, SuSMHead.ValueLen)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "Body.Write"
			strRet = msmBody.Write(DataBytes)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			'Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", strRet)
			Return strRet
		End Try
	End Function

	Private Sub mSavePigKeyValue2SM(ByRef NewItem As PigKeyValue)
		Const SUB_NAME As String = "mSavePigKeyValue2SM"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim suSMHead As StruSMHead
			ReDim suSMHead.ValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(suSMHead, NewItem.SMNameHead)
			If strRet <> "OK" Then
				With suSMHead
					.ValueType = NewItem.ValueType
					.ValueLen = NewItem.ValueLen
					.ValueMD5 = NewItem.ValueMD5Bytes
					.ExpTime = NewItem.ExpTime
				End With
				strStepName = "mSaveSMHead"
				strRet = Me.mSaveSMHead(suSMHead, NewItem.SMNameHead)
				If strRet <> "OK" Then
					strStepName &= "(" & NewItem.SMNameHead & ")"
					Throw New Exception(strRet)
				End If
			End If
			strStepName = "mSaveSMBody"
			strRet = Me.mSaveSMBody(suSMHead, NewItem.SMNameBody, NewItem.BytesValue)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			'Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", Me.LastErr)
		End Try
	End Sub

	Public Sub SavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True)
		Const SUB_NAME As String = "SavePigKeyValue"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			strStepName = "IsPigKeyValueExists"
			If Me.IsPigKeyValueExists(NewItem.KeyName) = True Then
				If IsOverwrite = True Then
					strStepName = "RemovePigKeyValue"
					strRet = Me.RemovePigKeyValue(NewItem.KeyName)
					If strRet <> "OK" Then
						strStepName &= "(" & NewItem.KeyName & ")"
						Throw New Exception(strRet)
					End If
				Else
					strStepName &= "(" & NewItem.KeyName & ")"
					Throw New Exception("PigKeyValue Exists")
				End If
			End If
			If Me.IsUseMemCache = True Then
				If NewItem.Parent Is Nothing Then NewItem.Parent = Me
				strStepName = "mSavePigKeyValue2SM"
				Me.mSavePigKeyValue2SM(NewItem)
				If Me.LastErr <> "" Then
					strStepName &= "(" & NewItem.KeyName & ")"
					Throw New Exception(Me.LastErr)
				End If
			End If
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

	Public Function RemovePigKeyValue(KeyName As String) As String
		Const SUB_NAME As String = "RemovePigKeyValue"
		Dim strStepName As String = ""
		Try
			strStepName = "IsPigKeyValueExists"
			If Me.IsPigKeyValueExists(KeyName) = False Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception("Not exists.")
			End If
			strStepName = "GetItem"
			Dim pkvAny As PigKeyValue = Me.PigKeyValues.Item(KeyName)
			If Me.PigKeyValues.LastErr <> "" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(Me.PigKeyValues.LastErr)
			End If
			If Me.IsUseMemCache = True Then
				If pkvAny Is Nothing Then
					strStepName = "New PigKeyValue"
					pkvAny = New PigKeyValue(KeyName, Now.AddMinutes(1), "")
					If pkvAny.LastErr <> "" Then
						strStepName &= "(" & KeyName & ")"
						Throw New Exception(pkvAny.LastErr)
					End If
				End If
				If pkvAny.Parent Is Nothing Then
					pkvAny.Parent = Me
				End If
				Dim strRet As String
				Dim SuSMHead As StruSMHead
				With SuSMHead
					.ExpTime = DateTime.MinValue
					.ValueType = PigKeyValue.enmValueType.Unknow
					.ValueLen = 0
					ReDim SuSMHead.ValueMD5(15)
				End With
				strStepName = "mSaveSMHead"
				strRet = Me.mSaveSMHead(SuSMHead, pkvAny.SMNameHead)
				If strRet <> "OK" Then
					strStepName &= "(" & KeyName & ")"
					Throw New Exception(strRet)
				End If
			End If
			If Me.PigKeyValues.IsItemExists(KeyName) = True Then
				strStepName = "PigKeyValues.Remove"
				Me.PigKeyValues.Remove(KeyName)
				If Me.PigKeyValues.LastErr <> "" Then
					strStepName &= "(" & KeyName & ")"
					Throw New Exception(Me.PigKeyValues.LastErr)
				End If
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Me.PrintDebugLog(SUB_NAME, "Catch Exception", strRet)
			Return strRet
		End Try
	End Function

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
