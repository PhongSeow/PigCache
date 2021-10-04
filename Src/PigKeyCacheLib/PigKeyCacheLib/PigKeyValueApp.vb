'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用|Piggy key value application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.7
'* Create Time: 8/5/2021
'* 1.0.2	13/5/2021 Modify New
'* 1.0.3	22/7/2021 Modify GetPigKeyValue
'* 1.0.4	23/7/2021 remove ObjAdoDBLib
'* 1.0.5	4/8/2021 Remove PigSQLSrvLib
'* 1.0.6	5/8/2021 Modify GetPigKeyValue,SavePigKeyValue
'* 1.0.7	7/8/2021 Modify New and add IsUseMemCache
'* 1.0.8	11/8/2021 Add mSavePigKeyValueToShareMem,mGetBytesSMBody
'* 1.0.9	13/8/2021 Modify mGetBytesSMBody
'* 1.0.10	13/8/2021 Modify mSaveSMHead,IsUseMemCache
'* 1.0.11	16/8/2021 Modify mSavePigKeyValueToShareMem,mSavePigKeyValueToShareMem,ShareMemRoot,GetPigKeyValue,mGetStruSMHead,mGetBytesSMBody, and add mSaveSMBody
'* 1.0.12	17/8/2021 Add PrintDebugLog,IsPigKeyValueExists,RemovePigKeyValue and modify GetPigKeyValue,SavePigKeyValue
'* 1.0.13	19/8/2021 Modify RemoveExpItems, and add GetStatisticsXml
'* 1.0.14	22/8/2021 Add CacheLevel,ForceRefCacheTime， and modify New,mNew,SavePigKeyValue,GetPigKeyValue,RemovePigKeyValue
'* 1.0.15	23/8/2021 Modify mNew,StruStatistics,New,GetStatisticsXml,IsPigKeyValueExists, and add CacheWorkDir,mIsShareMemExists
'* 1.0.16	23/8/2021 Modify GetPigKeyValue, and Add mGetPigKeyValueByShareMem
'* 1.0.17	25/8/2021 Remove Imports PigToolsLib, change to PigToolsWinLib, and add mIsBytesMatch, mSavePigKeyValueToSM rename to mSavePigKeyValueToShareMem
'* 1.0.18	26/8/2021 Modify RemovePigKeyValue,SavePigKeyValue, and add mClearShareMem
'* 1.0.19	27/8/2021 Modify mGetPigKeyValueByShareMem
'* 1.1		29/8/2021 Chanage PigToolsWinLib to PigToolsLiteLib
'* 1.2		31/8/2021 Modify ForceRefCacheTime
'* 1.3		25/9/2021 Add mSavePigKeyValueToFile,mGetStruFileHead,mSaveFileHead,mSaveFileBody
'* 1.4		26/9/2021 Modify mSavePigKeyValueToFile,SavePigKeyValue,GetPigKeyValue, and add mGetPigKeyValueByFile
'* 1.5		2/10/2021 Modify New,mNew,GetPigKeyValue
'* 1.6		3/10/2021 Add StruKeyValueCtrl,mRemoveFile,mGetPigKeyValueByList,mGetPigKeyValueByShareMem, and modify GetPigKeyValue,StruStatistics,GetStatisticsXml
'* 1.7		4/10/2021 Modify GetPigKeyValue,SavePigKeyValue,mAddPigKeyValueToList,mRemoveFile,RemovePigKeyValue, and add mGetPigKeyValueByFile
'************************************

Imports PigToolsLiteLib

Public Class PigKeyValueApp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.7.16"

	''' <summary>
	''' Value type, non text type, saved in byte array
	''' </summary>
	Public Enum enmCacheLevel
		Unknow = 0
		''' <summary>
		''' Program for single process multithreading
		''' </summary>
		ToList = 10
		''' <summary>
		''' It is applicable to multi-process and multi-threaded programs under the same user session or IIS application pools.
		''' </summary>
		ToShareMem = 20
		''' <summary>
		''' It is suitable for any multi process and multi thread program on the same host.
		''' </summary>
		ToFile = 30
		'''' <summary>
		'''' It is suitable for multi server, multi process and multi-threaded programs, and has the highest requirements for the availability of cached content, but the writing performance is poor, but the advantage is that it can share the database with the application to reduce the point of failure.
		'''' </summary>
		'ToDB = 40
		'''' <summary>
		'''' It is suitable for multi server, multi process and multi thread programs. The read and write performance is very good, but redis needs to be installed, which needs to increase the cost of managing the high availability of redis.
		'''' </summary>
		'ToRedis = 50
	End Enum

	Public ReadOnly Property PigKeyValues As New PigKeyValues
	'Public Property IsUseMemCache As Boolean = False
	Friend Property ShareMemRoot As String = ""

	Private msuStatistics As StruStatistics

	''' <summary>
	''' 统计信息结构
	''' </summary>
	Private Structure StruStatistics
		Dim GetCount As Long
		Dim GetFailCount As Long
		Dim CacheCount As Long
		Dim CacheByListCount As Long
		Dim CacheByShareMemCount As Long
		Dim CacheByFileCount As Long
		Dim CacheByDBCount As Long
		Dim CacheByRedisCount As Long
		Dim SaveCount As Long
		Dim SaveFailCount As Long
		Dim SaveToListCount As Long
		Dim SaveToShareMemCount As Long
		Dim SaveToFileCount As Long
		Dim SaveToDBCount As Long
		Dim SaveToRedisCount As Long
		Dim RemoveCount As Long
		Dim RemoveFailCount As Long
		Dim RemoveExpiredListCount As Long
		Dim RemoveExpiredShareMemCount As Long
		Dim RemoveExpiredFileCount As Long
		Dim RemoveExpiredDBCount As Long
		Dim RemoveExpiredRedisCount As Long
	End Structure

	''' <summary>
	''' 共享内存头结构
	''' </summary>
	Private Structure StruSMHead
		Dim ValueType As PigKeyValue.enmValueType
		Dim ExpTime As DateTime
		Dim ValueLen As Long
		Dim ValueMD5 As Byte()
	End Structure

	''' <summary>
	''' 键值控制结构
	''' </summary>
	Private Structure StruKeyValueCtrl
		Dim IsGetByShareMem As Boolean
		Dim IsGetByFile As Boolean
		Dim IsRemoveList As Boolean
		Dim IsClearShareMem As Boolean
		Dim IsRemoveFile As Boolean
		Dim IsRefLastRefCacheTime As Boolean
		Dim ListValueMD5 As String
		Dim ShareMemValueMD5 As String
		Dim IsSaveList As Boolean
		Dim IsSaveShareMem As Boolean
		Dim IsSaveFile As Boolean
	End Structure

	Public Sub New()
		MyBase.New(CLS_VERSION)
		mNew("", enmCacheLevel.ToList)
	End Sub

	Private Sub mNew(Optional ShareMemRootOrCacheWorkDir As String = "", Optional CacheLevel As enmCacheLevel = enmCacheLevel.ToShareMem, Optional ForceRefCacheTime As Integer = 60)
		Try
			Me.CacheLevel = CacheLevel
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					Me.ShareMemRoot = ""
					Me.CacheWorkDir = ""
				Case enmCacheLevel.ToShareMem
					If ShareMemRootOrCacheWorkDir = "" Then ShareMemRootOrCacheWorkDir = Me.AppTitle
					Me.ShareMemRoot = ShareMemRootOrCacheWorkDir
					Me.CacheWorkDir = ""
				Case enmCacheLevel.ToFile
					If ShareMemRootOrCacheWorkDir = "" Then ShareMemRootOrCacheWorkDir = Me.AppPath
					Me.ShareMemRoot = ShareMemRootOrCacheWorkDir
					Me.CacheWorkDir = ShareMemRootOrCacheWorkDir
				Case Else
					Throw New Exception("Currently unsupported cachelevel")
			End Select
			If Me.ShareMemRoot <> "" Then
				Dim oPigMD5 As PigMD5
				oPigMD5 = New PigMD5(ShareMemRootOrCacheWorkDir, PigMD5.enmTextType.UTF8)
				Me.ShareMemRoot = oPigMD5.PigMD5()
			End If
			If ForceRefCacheTime < 30 Then
				Me.ForceRefCacheTime = 30
			Else
				Me.ForceRefCacheTime = ForceRefCacheTime
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mNew", ex)
			Me.CacheLevel = enmCacheLevel.Unknow
		End Try
	End Sub
	Public Sub New(ShareMemRoot As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(ShareMemRoot)
	End Sub

	Public Sub New(ShareMemRootOrCacheWorkDir As String, CacheLevel As enmCacheLevel)
		MyBase.New(CLS_VERSION)
		Me.mNew(ShareMemRootOrCacheWorkDir, CacheLevel)
	End Sub


	Private Function mIsCacheFileExists(KeyName As String) As Boolean
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			strStepName = "New PigKeyValue"
			Dim pkvNew As New PigKeyValue(KeyName, Now.AddMinutes(1), "")
			pkvNew.Parent = Me
			Dim SuSMHead As StruSMHead
			ReDim SuSMHead.ValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(SuSMHead, pkvNew.SMNameHead, Me.CacheWorkDir)
			If strRet <> "OK" Then
				Return False
			Else
				Dim abSMBody As Byte()
				ReDim abSMBody(0)
				strStepName = "mGetBytesSMBody"
				strRet = Me.mGetBytesFileBody(abSMBody, SuSMHead, pkvNew.SMNameBody, Me.CacheWorkDir)
				If strRet <> "OK" Then
					Return False
				ElseIf abSMBody.Length <> SuSMHead.ValueLen Then
					Return False
				Else
					Dim oPigBytes As New PigBytes(abSMBody)
					'If oPigBytes.PigMD5Bytes.SequenceEqual(SuSMHead.ValueMD5) = False Then
					If Me.mIsBytesMatch(oPigBytes.PigMD5Bytes, SuSMHead.ValueMD5) = False Then
						Return False
					Else
						Return True
					End If
				End If
			End If
		Catch ex As Exception
			Me.SetSubErrInf("mIsCacheFileExists", strStepName, ex)
			Return False
		End Try
	End Function

	Private Function mGetPigKeyValueFromShareMem(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Const SUB_NAME As String = "mGetPigKeyValueFromShareMem"
		Dim strStepName As String = ""
		Dim strRet As String
		Try
			strStepName = "New PigKeyValue Get suSMHead"
			OutPigKeyValue = New PigKeyValue(KeyName, Now.AddMinutes(1), "")
			OutPigKeyValue.Parent = Me
			Dim suSMHead As StruSMHead
			ReDim suSMHead.ValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(suSMHead, OutPigKeyValue.SMNameHead)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")" : Throw New Exception(strRet)
			End If

			Dim abSMBody As Byte()
			ReDim abSMBody(0)
			strStepName = "mGetBytesSMBody"
			strRet = Me.mGetBytesSMBody(abSMBody, suSMHead, OutPigKeyValue.SMNameBody)
			If strRet = "OK" Then
				If abSMBody.Length <> suSMHead.ValueLen Then
					strRet = "SMBody.Length<>SuSMHead.ValueLen"
				Else
					Dim oPigBytes As New PigBytes(abSMBody)
					'If oPigBytes.PigMD5Bytes.SequenceEqual(suSMHead.ValueMD5) = False Then
					If Me.mIsBytesMatch(oPigBytes.PigMD5Bytes, suSMHead.ValueMD5) = False Then
						strRet = "SMBody.PigMD5<>SuSMHead.ValueMD5"
					End If
					oPigBytes = Nothing
				End If
			End If
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")" : Throw New Exception(strRet)
			End If
			OutPigKeyValue = Nothing
			strStepName = "New PigKeyValue To Out"
			OutPigKeyValue = New PigKeyValue(KeyName, suSMHead.ExpTime, abSMBody, suSMHead.ValueType, suSMHead.ValueMD5)
			If OutPigKeyValue.LastErr <> "" Then
				strStepName &= strStepName & "(" & KeyName & ")"
				Throw New Exception(OutPigKeyValue.LastErr)
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

	Private Function mGetSMNamePart(KeyName As String, ByRef SMNameHead As String, ByRef SMNameBody As String) As String
		Dim strStepName As String = ""
		Try
			strStepName = "New PigKeyValue"
			Dim pkvNew As New PigKeyValue(KeyName, Now.AddMinutes(1), "")
			pkvNew.Parent = Me
			strStepName = "GetPartName"
			SMNameHead = pkvNew.SMNameHead
			SMNameBody = pkvNew.SMNameBody
			pkvNew = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("mGetSMNamePart", strStepName, ex)
		End Try
	End Function

	Private Function mClearShareMem(KeyName As String) As String
		Const SUB_NAME As String = "mClearShareMem"
		Dim strStepName As String = "", strRet As String = ""
		Try
			Dim strSMNameHead As String = "", strSMNameBody As String = ""
			strStepName = "mGetSMNamePart"
			strRet = Me.mGetSMNamePart(KeyName, strSMNameHead, strSMNameBody)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")" : Throw New Exception(strRet)
			End If
			Dim SuSMHead As StruSMHead
			ReDim SuSMHead.ValueMD5(15)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(SuSMHead, strSMNameHead)
			If strRet = "OK" Then
				Dim intBodyLen As Integer = SuSMHead.ValueLen
				With SuSMHead
					.ExpTime = DateTime.MinValue
					.ValueLen = 0
					.ValueType = PigKeyValue.enmValueType.Unknow
					ReDim .ValueMD5(15)
				End With
				strStepName = "mSaveSMBody"
				strRet = Me.mSaveSMHead(SuSMHead, strSMNameHead)
				If strRet <> "OK" Then Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
				Dim abBody As Byte()
				ReDim abBody(intBodyLen - 1)
				strStepName = "mSaveSMBody"
				strRet = Me.mSaveSMBody(SuSMHead, strSMNameBody, abBody)
				If strRet <> "OK" Then
					strStepName &= "(" & KeyName & ")"
					Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
				End If
			Else
				strStepName &= "(" & KeyName & ")"
				Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("mClearShareMem", strStepName, ex)
		End Try
	End Function

	Private Function mIsShareMemExists(KeyName As String) As Boolean
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			strStepName = "New PigKeyValue"
			Dim pkvNew As New PigKeyValue(KeyName, Now.AddMinutes(1), "")
			pkvNew.Parent = Me
			Dim SuSMHead As StruSMHead
			ReDim SuSMHead.ValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(SuSMHead, pkvNew.SMNameHead)
			If strRet <> "OK" Then
				Return False
			Else
				Dim abSMBody As Byte()
				ReDim abSMBody(0)
				strStepName = "mGetBytesSMBody"
				strRet = Me.mGetBytesSMBody(abSMBody, SuSMHead, pkvNew.SMNameBody)
				If strRet <> "OK" Then
					Return False
				ElseIf abSMBody.Length <> SuSMHead.ValueLen Then
					Return False
				Else
					Dim oPigBytes As New PigBytes(abSMBody)
					'If oPigBytes.PigMD5Bytes.SequenceEqual(SuSMHead.ValueMD5) = False Then
					If Me.mIsBytesMatch(oPigBytes.PigMD5Bytes, SuSMHead.ValueMD5) = False Then
						Return False
					Else
						Return True
					End If
				End If
			End If
		Catch ex As Exception
			Me.SetSubErrInf("mIsShareMemExists", strStepName, ex)
			Return False
		End Try
	End Function

	Public Function IsPigKeyValueExists(KeyName As String) As Boolean
		Dim strStepName As String = ""
		Try
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					If Me.PigKeyValues.IsItemExists(KeyName) = True Then
						Return True
					Else
						Return False
					End If
				Case enmCacheLevel.ToShareMem
					Return Me.mIsShareMemExists(KeyName)
				Case enmCacheLevel.ToFile
					Return Me.mIsCacheFileExists(KeyName)
				Case Else
					strStepName = ""
					Throw New Exception("Currently unsupported cachelevel")
			End Select
		Catch ex As Exception
			Me.SetSubErrInf("IsPigKeyValueExists", ex)
			Return False
		End Try
	End Function

	'Private Function mIsForceRefCache() As Boolean
	'	Try
	'		If Math.Abs(DateDiff(DateInterval.Second, Me.mLastRefCacheTime, Now)) > Me.ForceRefCacheTime Then
	'			Return True
	'		Else
	'			Return False
	'		End If
	'	Catch ex As Exception
	'		Me.SetSubErrInf("mIsForceRefCache", ex)
	'		Return False
	'	End Try
	'End Function

	Private Function mRemovePigKeyValueFromList(KeyName As String) As String
		Dim strStepName As String = ""
		Try
			If Me.PigKeyValues.IsItemExists(KeyName) = False Then
				Return "OK"
			Else
				strStepName = "PigKeyValues.Remove"
				Me.PigKeyValues.Remove(KeyName)
				If Me.PigKeyValues.LastErr <> "" Then
					strStepName &= "(" & KeyName & ")"
					Throw New Exception(Me.PigKeyValues.LastErr)
				End If
				Return "OK"
			End If
		Catch ex As Exception
			Return Me.GetSubErrInf("mRemovePigKeyValueFromList", ex)
		End Try
	End Function


	Private Function mAddPigKeyValueToList(NewItem As PigKeyValue) As String
		Const SUB_NAME As String = "mAddPigKeyValueToList"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			Dim strKeyName As String = NewItem.KeyName
			'If Me.PigKeyValues.IsItemExists(strKeyName) = True Then
			'	strStepName = "mRemovePigKeyValueFromList"
			'	strRet = Me.mRemovePigKeyValueFromList(strKeyName)
			'	If strRet <> "OK" Then
			'		Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
			'		If Me.PigKeyValues.IsItemExists(strKeyName) Then
			'			strStepName &= "(" & strKeyName & ")"
			'			Throw New Exception("Cannot remove exists item")
			'		End If
			'	End If
			'End If
			NewItem.LastRefCacheTime = Now
			strStepName = "PigKeyValues.Add"
			Me.PigKeyValues.Add(NewItem)
			If Me.PigKeyValues.LastErr <> "" Then
				strStepName &= "(" & strKeyName & ")"
				Throw New Exception(strKeyName)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

	Private Function mGetPigKeyValueByShareMem(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Const SUB_NAME As String = "mGetPigKeyValueByShareMem"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			msuStatistics.GetCount += 1
			strStepName = "mGetPigKeyValueFromShareMem"
			strRet = Me.mGetPigKeyValueFromShareMem(KeyName, OutPigKeyValue)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
			End If
			If Not OutPigKeyValue Is Nothing Then
				If OutPigKeyValue.IsExpired = True Then
					msuStatistics.RemoveCount += 1
					msuStatistics.RemoveExpiredShareMemCount += 1
					strStepName = "mClearShareMem"
					strRet = Me.mClearShareMem(KeyName)
					If strRet <> "OK" Then
						strStepName &= "(" & KeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
						msuStatistics.RemoveFailCount += 1
					End If
					OutPigKeyValue = Nothing
					If Me.PigKeyValues.IsItemExists(KeyName) = True Then
						msuStatistics.RemoveCount += 1
						msuStatistics.RemoveExpiredListCount += 1
						strStepName = "mRemovePigKeyValueFromList"
						strRet = Me.mRemovePigKeyValueFromList(KeyName)
						If strRet <> "OK" Then
							strStepName &= "(" & KeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							msuStatistics.RemoveFailCount += 1
						End If
					End If
				Else
					msuStatistics.CacheCount += 1
					msuStatistics.CacheByShareMemCount += 1
				End If
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

	Private Function mGetPigKeyValueByFile(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Const SUB_NAME As String = "mGetPigKeyValueByFile"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			msuStatistics.GetCount += 1
			strStepName = "mGetPigKeyValueFromFile"
			strRet = Me.mGetPigKeyValueFromFile(KeyName, OutPigKeyValue)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
			End If
			If Not OutPigKeyValue Is Nothing Then
				If OutPigKeyValue.IsExpired = True Then
					msuStatistics.RemoveCount += 1
					msuStatistics.RemoveExpiredFileCount += 1
					If OutPigKeyValue.Parent Is Nothing Then OutPigKeyValue.Parent = Me
					strStepName = "mRemoveFile"
					strRet = Me.mRemoveFile(OutPigKeyValue)
					If strRet <> "OK" Then
						strStepName &= "(" & KeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
						msuStatistics.RemoveFailCount += 1
					End If
					OutPigKeyValue = Nothing
					If Me.mIsShareMemExists(KeyName) = True Then
						msuStatistics.RemoveCount += 1
						msuStatistics.RemoveExpiredShareMemCount += 1
						strStepName = "mClearShareMem"
						strRet = Me.mClearShareMem(KeyName)
						If strRet <> "OK" Then
							strStepName &= "(" & KeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							msuStatistics.RemoveFailCount += 1
						End If
					End If
					If Me.PigKeyValues.IsItemExists(KeyName) = True Then
						msuStatistics.RemoveCount += 1
						msuStatistics.RemoveExpiredListCount += 1
						strStepName = "mRemovePigKeyValueFromList"
						strRet = Me.mRemovePigKeyValueFromList(KeyName)
						If strRet <> "OK" Then
							strStepName &= "(" & KeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							msuStatistics.RemoveFailCount += 1
						End If
					End If
				Else
					msuStatistics.CacheCount += 1
					msuStatistics.CacheByFileCount += 1
				End If
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

	Private Function mGetPigKeyValueByList(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Const SUB_NAME As String = "mGetPigKeyValueByList"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			msuStatistics.GetCount += 1
			strStepName = "GetByList"
			OutPigKeyValue = Me.PigKeyValues.Item(KeyName)
			If Me.PigKeyValues.LastErr <> "" Then
				strStepName &= "(" & KeyName & ")"
				Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
			End If
			If Not OutPigKeyValue Is Nothing Then
				If OutPigKeyValue.IsExpired = True Then
					msuStatistics.RemoveCount += 1
					msuStatistics.RemoveExpiredListCount += 1
					strStepName = "mRemovePigKeyValueFromList"
					strRet = Me.mRemovePigKeyValueFromList(KeyName)
					If strRet <> "OK" Then
						strStepName &= "(" & KeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
						msuStatistics.RemoveFailCount += 1
					End If
					OutPigKeyValue = Nothing
				Else
					msuStatistics.CacheCount += 1
					msuStatistics.CacheByListCount += 1
				End If
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function
	Public Function GetPigKeyValue(KeyName As String) As PigKeyValue
		Const SUB_NAME As String = "GetKeyValue"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			GetPigKeyValue = Nothing
			Dim pkvList As PigKeyValue = Nothing
			strStepName = "mGetPigKeyValueByList"
			strRet = Me.mGetPigKeyValueByList(KeyName, pkvList)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(strRet)
			End If
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					GetPigKeyValue = pkvList
					pkvList = Nothing
				Case enmCacheLevel.ToShareMem
					Dim bolIsGetByShareMem As Boolean = False
					If pkvList Is Nothing Then
						bolIsGetByShareMem = True
					Else
						If pkvList.Parent Is Nothing Then pkvList.Parent = Me
						If pkvList.IsForceRefCache = True Then
							bolIsGetByShareMem = True
						End If
					End If
					If bolIsGetByShareMem = True Then
						strStepName = "mGetPigKeyValueByShareMem.ToShareMem"
						strRet = Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
						If strRet <> "OK" Then
							strStepName &= "(" & KeyName & ")"
							Throw New Exception(strRet)
						End If
						Dim bolIsRemoveList As Boolean = False
						Dim bolIsAddList As Boolean = False
						If GetPigKeyValue Is Nothing Then
							If Not pkvList Is Nothing Then bolIsRemoveList = True
						ElseIf Not pkvList Is Nothing Then
							If GetPigKeyValue.CompareOther(pkvList) = False Then
								bolIsRemoveList = True
								bolIsAddList = True
							Else
								pkvList.LastRefCacheTime = Now
							End If
						Else
							bolIsAddList = True
						End If
						If bolIsRemoveList = True Then
							strStepName = "mClearShareMem"
							strRet = Me.mClearShareMem(KeyName)
							If strRet <> "OK" Then
								strStepName &= "(" & KeyName & ")"
								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
								msuStatistics.RemoveFailCount += 1
							End If
							pkvList = Nothing
						End If
						If bolIsAddList = True Then
							msuStatistics.SaveToListCount += 1
							strStepName = "mAddPigKeyValueToList"
							strRet = Me.mAddPigKeyValueToList(GetPigKeyValue)
							If strRet <> "OK" Then
								strStepName &= "(" & KeyName & ")"
								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
								msuStatistics.SaveFailCount += 1
							End If
						End If
					Else
						GetPigKeyValue = pkvList
						pkvList = Nothing
					End If
				Case enmCacheLevel.ToFile
					Dim bolIsGetByShareMem As Boolean = False
					Dim bolIsGetByFile As Boolean = False
					If pkvList Is Nothing Then
						bolIsGetByShareMem = True
					Else
						If pkvList.Parent Is Nothing Then pkvList.Parent = Me
						If pkvList.IsForceRefCache = True Then
							bolIsGetByFile = True
						End If
					End If
					Dim pkvShareMem As PigKeyValue = Nothing
					If bolIsGetByFile = True Then
						strStepName = "mGetPigKeyValueByFile.ToFile"
						strRet = Me.mGetPigKeyValueByFile(KeyName, GetPigKeyValue)
						If strRet <> "OK" Then
							strStepName &= "(" & KeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
						End If
						If GetPigKeyValue Is Nothing Then
							strStepName = "RemovePigKeyValue.ToShareMem"
							strRet = Me.RemovePigKeyValue(KeyName, enmCacheLevel.ToShareMem)
							If strRet <> "OK" Then
								strStepName &= "(" & KeyName & ")"
								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							End If
						Else
							If Me.mIsShareMemExists(KeyName) = True Then
								If GetPigKeyValue.Parent Is Nothing Then GetPigKeyValue.Parent = Me
								strStepName = "mGetPigKeyValueByShareMem.ToFile"
								strRet = Me.mGetPigKeyValueByShareMem(KeyName, pkvShareMem)
								If strRet <> "OK" Then
									strStepName &= "(" & KeyName & ")"
									Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
								End If
								Dim bolIsSaveShareMem As Boolean = False
								If Not pkvShareMem Is Nothing Then
									If pkvShareMem.CompareOther(GetPigKeyValue) = False Then
										strStepName = "mClearShareMem.ToFile"
										strRet = Me.mClearShareMem(KeyName)
										If strRet <> "OK" Then
											strStepName &= "(" & KeyName & ")"
											Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
										End If
										bolIsSaveShareMem = True
									End If
								Else
									bolIsSaveShareMem = True
								End If
								If bolIsSaveShareMem = True Then
									strStepName = "mSavePigKeyValueToShareMem.ToFile"
									strRet = Me.mSavePigKeyValueToShareMem(GetPigKeyValue)
									If strRet <> "OK" Then
										strStepName &= "(" & KeyName & ")"
										Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
									End If
								End If
							End If
						End If
					ElseIf bolIsGetByShareMem = True Then
						strStepName = "mGetPigKeyValueByShareMem.ToFile"
						strRet = Me.mGetPigKeyValueByShareMem(KeyName, pkvShareMem)
						If strRet <> "OK" Then
							strStepName &= "(" & KeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
						End If
						If pkvShareMem Is Nothing Then
							strStepName = "mGetPigKeyValueByFile.ToFile2"
							strRet = Me.mGetPigKeyValueByFile(KeyName, GetPigKeyValue)
							If strRet <> "OK" Then
								strStepName &= "(" & KeyName & ")"
								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							End If
							If Not GetPigKeyValue Is Nothing Then
								If GetPigKeyValue.Parent Is Nothing Then GetPigKeyValue.Parent = Me
								strStepName = "mSavePigKeyValueToShareMem.ToFile2"
								strRet = Me.mSavePigKeyValueToShareMem(GetPigKeyValue)
								If strRet <> "OK" Then
									strStepName &= "(" & KeyName & ")"
									Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
								End If
							End If
						Else
							GetPigKeyValue = pkvShareMem
							pkvList = Nothing
						End If
					End If
				Case Else
					strStepName = KeyName
					Throw New Exception("Unsupported CacheLevel")
			End Select
			Me.ClearErr()
		Catch ex As Exception
			msuStatistics.GetFailCount += 1
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			Return Nothing
		End Try
	End Function

	'Public Function GetPigKeyValue(KeyName As String) As PigKeyValue
	'	Const SUB_NAME As String = "GetKeyValue"
	'	Dim strStepName As String = ""
	'	Dim strRet As String = ""
	'	Try
	'		Dim suKeyValueCtrl As StruKeyValueCtrl
	'		msuStatistics.GetCount += 1
	'		strStepName = "GetByList"
	'		GetPigKeyValue = Me.PigKeyValues.Item(KeyName)
	'		If Me.PigKeyValues.LastErr <> "" Then
	'			strStepName &= "(" & KeyName & ")"
	'			Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
	'		End If
	'		With suKeyValueCtrl
	'			.IsRefLastRefCacheTime = False
	'			If GetPigKeyValue Is Nothing Then
	'				.IsRemoveList = False
	'			ElseIf GetPigKeyValue.IsExpired = True Then
	'				.IsRemoveList = True
	'			Else
	'				.IsRemoveList = False
	'			End If
	'		End With
	'		If suKeyValueCtrl.IsRemoveList = True Then
	'			msuStatistics.RemoveListCount += 1
	'			strStepName = "mRemovePigKeyValueFromList.List"
	'			strRet = Me.mRemovePigKeyValueFromList(KeyName)
	'			If strRet <> "OK" Then
	'				strStepName &= "(" & KeyName & ")"
	'				Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'				msuStatistics.RemoveFailCount += 1
	'			End If
	'			GetPigKeyValue = Nothing
	'			suKeyValueCtrl.IsRemoveList = False
	'		End If
	'		Select Case Me.CacheLevel
	'			Case enmCacheLevel.ToShareMem, enmCacheLevel.ToFile
	'				If GetPigKeyValue Is Nothing Then
	'					suKeyValueCtrl.ListValueMD5 = ""
	'				Else
	'					suKeyValueCtrl.ListValueMD5 = GetPigKeyValue.ValueMD5
	'				End If
	'				If suKeyValueCtrl.ListValueMD5 = "" Then
	'					suKeyValueCtrl.IsGetByShareMem = True
	'				ElseIf Me.mIsForceRefCache = True Then
	'					suKeyValueCtrl.IsGetByShareMem = True
	'					suKeyValueCtrl.IsRefLastRefCacheTime = True
	'				Else
	'					suKeyValueCtrl.IsGetByShareMem = False
	'				End If
	'				If suKeyValueCtrl.IsGetByShareMem = True Then
	'					strStepName = "mGetPigKeyValueByShareMem"
	'					Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
	'					If Me.LastErr <> "" Then
	'						strStepName &= "(" & KeyName & ")"
	'						Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
	'					End If
	'					If GetPigKeyValue Is Nothing Then
	'						suKeyValueCtrl.IsClearShareMem = False
	'						If suKeyValueCtrl.ListValueMD5 <> "" Then suKeyValueCtrl.IsRemoveList = True
	'					ElseIf GetPigKeyValue.IsExpired = True Then
	'						suKeyValueCtrl.IsClearShareMem = True
	'						If suKeyValueCtrl.ListValueMD5 <> "" Then suKeyValueCtrl.IsRemoveList = True
	'					ElseIf GetPigKeyValue.ValueMD5 <> suKeyValueCtrl.ListValueMD5 And suKeyValueCtrl.ListValueMD5 <> "" Then
	'						suKeyValueCtrl.IsRemoveList = True
	'					Else
	'						suKeyValueCtrl.IsClearShareMem = False
	'					End If
	'					If suKeyValueCtrl.IsRemoveList = True Then
	'						msuStatistics.RemoveListCount += 1
	'						strStepName = "mRemovePigKeyValueFromList.ShareMem"
	'						strRet = Me.mRemovePigKeyValueFromList(KeyName)
	'						If strRet <> "OK" Then
	'							strStepName &= "(" & KeyName & ")"
	'							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'							msuStatistics.RemoveFailCount += 1
	'						End If
	'						suKeyValueCtrl.IsRemoveList = False
	'						suKeyValueCtrl.ListValueMD5 = ""
	'					End If
	'					If suKeyValueCtrl.IsClearShareMem = True Then
	'						msuStatistics.ClearShareMemCount += 1
	'						strStepName = "mClearShareMem.ShareMem"
	'						strRet = Me.mClearShareMem(KeyName)
	'						If strRet <> "OK" Then
	'							strStepName &= "(" & KeyName & ")"
	'							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'							msuStatistics.RemoveFailCount += 1
	'						End If

	'						suKeyValueCtrl.IsClearShareMem = False
	'						GetPigKeyValue = Nothing
	'					End If
	'					If Me.CacheLevel = enmCacheLevel.ToFile Then
	'						If GetPigKeyValue Is Nothing Then
	'							suKeyValueCtrl.ShareMemValueMD5 = ""
	'						Else
	'							suKeyValueCtrl.ShareMemValueMD5 = GetPigKeyValue.ValueMD5
	'						End If
	'						If suKeyValueCtrl.ShareMemValueMD5 = "" Then
	'							suKeyValueCtrl.IsGetByFile = True
	'						ElseIf Me.mIsForceRefCache = True Then
	'							suKeyValueCtrl.IsGetByFile = True
	'							suKeyValueCtrl.IsRefLastRefCacheTime = True
	'						Else
	'							suKeyValueCtrl.IsGetByFile = False
	'						End If
	'						If suKeyValueCtrl.IsGetByFile = True Then
	'							strStepName = "mGetPigKeyValueByFile"
	'							Me.mGetPigKeyValueByFile(KeyName, GetPigKeyValue)
	'							If Me.LastErr <> "" Then
	'								strStepName &= "(" & KeyName & ")"
	'								Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
	'							End If
	'							If GetPigKeyValue Is Nothing Then
	'								suKeyValueCtrl.IsClearShareMem = True
	'								suKeyValueCtrl.IsRemoveList = True
	'								suKeyValueCtrl.IsSaveShareMem = False
	'							ElseIf GetPigKeyValue.IsExpired = True Then
	'								suKeyValueCtrl.IsClearShareMem = True
	'								suKeyValueCtrl.IsRemoveList = True
	'								suKeyValueCtrl.IsSaveShareMem = False
	'								If GetPigKeyValue.Parent Is Nothing Then GetPigKeyValue.Parent = Me
	'								msuStatistics.RemoveFileCount += 1
	'								strStepName = "mRemoveFile.File"
	'								strRet = Me.mRemoveFile(GetPigKeyValue)
	'								If strRet <> "OK" Then
	'									strStepName &= "(" & KeyName & ")"
	'									Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'									msuStatistics.RemoveFailCount += 1
	'								End If
	'								GetPigKeyValue = Nothing
	'							ElseIf GetPigKeyValue.ValueMD5 <> suKeyValueCtrl.ShareMemValueMD5 And suKeyValueCtrl.ShareMemValueMD5 <> "" Then
	'								suKeyValueCtrl.IsClearShareMem = True
	'								suKeyValueCtrl.IsRemoveList = True
	'								suKeyValueCtrl.IsSaveShareMem = True
	'							Else
	'								suKeyValueCtrl.IsSaveShareMem = True
	'							End If
	'							If suKeyValueCtrl.IsRemoveList = True Then
	'								msuStatistics.RemoveListCount += 1
	'								strStepName = "mRemovePigKeyValueFromList.File"
	'								strRet = Me.mRemovePigKeyValueFromList(KeyName)
	'								If strRet <> "OK" Then
	'									strStepName &= "(" & KeyName & ")"
	'									Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'									msuStatistics.RemoveFailCount += 1
	'								End If
	'								suKeyValueCtrl.IsRemoveList = False
	'							End If
	'							If suKeyValueCtrl.IsClearShareMem = True Then
	'								msuStatistics.ClearShareMemCount += 1
	'								strStepName = "mClearShareMem.File"
	'								strRet = Me.mClearShareMem(KeyName)
	'								If strRet <> "OK" Then
	'									strStepName &= "(" & KeyName & ")"
	'									Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'									msuStatistics.RemoveFailCount += 1
	'								End If
	'								suKeyValueCtrl.IsClearShareMem = False
	'							End If
	'							If suKeyValueCtrl.IsSaveShareMem = True Then
	'								If GetPigKeyValue.Parent Is Nothing Then GetPigKeyValue.Parent = Me
	'								msuStatistics.SaveToShareMemCount += 1
	'								strStepName = "mSavePigKeyValueToShareMem.File"
	'								strRet = Me.mSavePigKeyValueToShareMem(GetPigKeyValue)
	'								If strRet <> "OK" Then
	'									strStepName &= "(" & KeyName & ")"
	'									Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'									msuStatistics.SaveFailCount += 1
	'								End If
	'								suKeyValueCtrl.IsSaveShareMem = False
	'							End If
	'							If Not GetPigKeyValue Is Nothing Then
	'								msuStatistics.CacheByFileCount += 1
	'								msuStatistics.CacheCount += 1
	'							End If
	'						End If
	'					ElseIf Not GetPigKeyValue Is Nothing Then
	'						msuStatistics.CacheByShareMemCount += 1
	'						msuStatistics.CacheCount += 1
	'						If suKeyValueCtrl.ListValueMD5 = "" Then
	'							suKeyValueCtrl.IsSaveList = True
	'						ElseIf GetPigKeyValue.ValueMD5 <> suKeyValueCtrl.ListValueMD5 And suKeyValueCtrl.ListValueMD5 <> "" Then
	'							suKeyValueCtrl.IsSaveList = True
	'						Else
	'							suKeyValueCtrl.IsSaveList = False
	'						End If
	'						If suKeyValueCtrl.IsSaveList = True Then
	'							msuStatistics.SaveToListCount += 1
	'							strStepName = "mAddPigKeyValueToList.ShareMem"
	'							strRet = Me.mAddPigKeyValueToList(GetPigKeyValue)
	'							If strRet <> "OK" Then
	'								strStepName &= "(" & KeyName & ")"
	'								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'								msuStatistics.SaveFailCount += 1
	'							End If
	'							suKeyValueCtrl.IsSaveList = False
	'						End If
	'					End If
	'				Else
	'					msuStatistics.CacheByListCount += 1
	'					msuStatistics.CacheCount += 1
	'				End If
	'				If suKeyValueCtrl.IsRefLastRefCacheTime = True Then
	'					Me.mLastRefCacheTime = Now
	'					suKeyValueCtrl.IsRefLastRefCacheTime = False
	'				End If
	'			Case enmCacheLevel.ToList
	'				If Not GetPigKeyValue Is Nothing Then
	'					msuStatistics.CacheByListCount += 1
	'					msuStatistics.CacheCount += 1
	'				End If
	'			Case Else
	'				strStepName = KeyName
	'				Throw New Exception("Unsupported CacheLevel")
	'		End Select



	'		'Select Case Me.CacheLevel
	'		'	Case enmCacheLevel.ToShareMem, enmCacheLevel.ToFile
	'		'		Dim strListValueMD5 As String
	'		'		If GetPigKeyValue Is Nothing Then
	'		'			strListValueMD5 = ""
	'		'		Else
	'		'			strListValueMD5 = GetPigKeyValue.ValueMD5
	'		'		End If
	'		'		Dim bolIsGetByShareMem As Boolean = False, bolGetByFile As Boolean = False, bolIsRemoveList As Boolean = False, bolIsRemoveShareMem As Boolean = False
	'		'		If strListValueMD5 = "" Then
	'		'			bolIsGetByShareMem = True
	'		'		ElseIf Me.mIsForceRefCache = True Then
	'		'			bolIsGetByShareMem = True
	'		'		End If
	'		'		If bolIsGetByShareMem = True Then
	'		'			strStepName = "mGetPigKeyValueByShareMem"
	'		'			Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
	'		'			If Me.LastErr <> "" Then
	'		'				strStepName &= "(" & KeyName & ")"
	'		'				Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
	'		'			End If
	'		'			If GetPigKeyValue Is Nothing And Me.CacheLevel = enmCacheLevel.ToFile Then
	'		'				bolGetByFile = True
	'		'			End If
	'		'		End If




	'		'		If Me.mIsForceRefCache = True Then
	'		'			If Not GetPigKeyValue Is Nothing Then
	'		'				strStepName = "PigKeyValues.Remove"
	'		'				Me.PigKeyValues.Remove(KeyName)
	'		'				If Me.PigKeyValues.LastErr <> "" Then
	'		'					strStepName &= "(" & KeyName & ")"
	'		'					Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
	'		'				End If
	'		'				GetPigKeyValue = Nothing
	'		'				Me.mLastRefCacheTime = Now
	'		'			End If
	'		'		End If
	'		'		If Me.mIsForceRefCache = True Then
	'		'			strStepName = "mGetPigKeyValueByShareMem"
	'		'			Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
	'		'			If Me.LastErr <> "" Then
	'		'				strStepName &= "(" & KeyName & ")"
	'		'				Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
	'		'			End If

	'		'		End If



	'		'		Select Case Me.CacheLevel
	'		'			Case enmCacheLevel.ToShareMem
	'		'				If GetPigKeyValue Is Nothing Then
	'		'					strStepName = "mGetPigKeyValueByShareMem"
	'		'					Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
	'		'					If Me.LastErr <> "" Then
	'		'						strStepName &= "(" & KeyName & ")"
	'		'						Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
	'		'					End If
	'		'					If GetPigKeyValue Is Nothing Then
	'		'						strStepName = "mRemovePigKeyValueFromList"
	'		'						strRet = Me.mRemovePigKeyValueFromList(KeyName)
	'		'						If strRet <> "OK" Then
	'		'							strStepName &= "(" & KeyName & ")"
	'		'							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'		'						End If
	'		'					Else
	'		'						strStepName = "mAddPigKeyValueToList"
	'		'						strRet = Me.mAddPigKeyValueToList(GetPigKeyValue)
	'		'						If strRet <> "OK" Then
	'		'							strStepName &= "(" & KeyName & ")"
	'		'							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'		'						End If
	'		'						msuStatistics.CacheByShareMemCount += 1
	'		'						msuStatistics.CacheCount += 1
	'		'					End If
	'		'				Else
	'		'					msuStatistics.CacheByListCount += 1
	'		'					msuStatistics.CacheCount += 1
	'		'				End If
	'		'			Case enmCacheLevel.ToFile
	'		'				If GetPigKeyValue Is Nothing Then
	'		'					strStepName = "mGetPigKeyValueByShareMem"
	'		'					Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
	'		'					If Me.LastErr <> "" Then
	'		'						strStepName &= "(" & KeyName & ")"
	'		'						Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
	'		'					End If
	'		'					If GetPigKeyValue Is Nothing Or Me.mIsForceRefCache = True Then
	'		'						strStepName = "mGetPigKeyValueByFile"
	'		'						Me.mGetPigKeyValueByFile(KeyName, GetPigKeyValue)
	'		'						If Me.LastErr <> "" Then
	'		'							strStepName &= "(" & KeyName & ")"
	'		'							Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
	'		'						End If
	'		'						If GetPigKeyValue Is Nothing Then
	'		'							strStepName = "mClearShareMem"
	'		'							strRet = Me.mClearShareMem(KeyName)
	'		'							If strRet <> "OK" Then
	'		'								strStepName &= "(" & KeyName & ")"
	'		'								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'		'							End If
	'		'						Else
	'		'							If GetPigKeyValue.Parent Is Nothing Then GetPigKeyValue.Parent = Me
	'		'							strStepName = "mSavePigKeyValueToShareMem"
	'		'							strRet = Me.mSavePigKeyValueToShareMem(GetPigKeyValue)
	'		'							If strRet <> "OK" Then
	'		'								strStepName &= "(" & KeyName & ")"
	'		'								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'		'							End If
	'		'							msuStatistics.CacheByFileCount += 1
	'		'							msuStatistics.CacheCount += 1
	'		'						End If
	'		'					Else
	'		'						strStepName = "mAddPigKeyValueToList"
	'		'						strRet = Me.mAddPigKeyValueToList(GetPigKeyValue)
	'		'						If strRet <> "OK" Then
	'		'							strStepName &= "(" & KeyName & ")"
	'		'							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
	'		'						End If
	'		'						msuStatistics.CacheByShareMemCount += 1
	'		'						msuStatistics.CacheCount += 1
	'		'					End If
	'		'				Else
	'		'					msuStatistics.CacheByListCount += 1
	'		'					msuStatistics.CacheCount += 1
	'		'				End If
	'		'		End Select
	'		'	Case enmCacheLevel.ToList
	'		'	Case Else
	'		'End Select
	'		Me.ClearErr()
	'	Catch ex As Exception
	'		msuStatistics.GetFailCount += 1
	'		Me.SetSubErrInf(SUB_NAME, strStepName, ex)
	'		Return Nothing
	'	End Try
	'End Function

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
			'If pbBody.PigMD5Bytes.SequenceEqual(SuSMHead.ValueMD5) = False Then
			If Me.mIsBytesMatch(pbBody.PigMD5Bytes, SuSMHead.ValueMD5) = False Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception("PigMD5 not match")
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

	Private Function mGetBytesFileBody(ByRef BodyBytes As Byte(), SuSMHead As StruSMHead, SMNameBody As String, CacheWorkDir As String) As String
		Const SUB_NAME As String = "mGetBytesFileBody"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim strFilePath As String = CacheWorkDir & Me.OsPathSep & SMNameBody
			strStepName = "New PigFile"
			Dim pfBody As New PigFile(strFilePath)
			If pfBody.LastErr <> "" Then
				strStepName &= "(" & strFilePath & ")"
				Throw New Exception(pfBody.LastErr)
			End If
			ReDim BodyBytes(0)
			strStepName = "Body.LoadFile"
			strRet = pfBody.LoadFile
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "Body Get Bytes"
			BodyBytes = pfBody.GbMain.Main
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
			'If pbBody.PigMD5Bytes.SequenceEqual(SuSMHead.ValueMD5) = False Then
			If Me.mIsBytesMatch(pbBody.PigMD5Bytes, SuSMHead.ValueMD5) = False Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception("PigMD5 not match")
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

	Private Function mGetStruSMHead(ByRef SuSMHead As StruSMHead, SMNameHead As String, CacheWorkDir As String) As String
		Const SUB_NAME As String = "mGetStruSMHead"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim strFilePath As String = CacheWorkDir & Me.OsPathSep & SMNameHead
			strStepName = "New PigFile"
			Dim pfHead As New PigFile(strFilePath)
			If pfHead.LastErr <> "" Then
				strStepName &= "(" & strFilePath & ")"
				Throw New Exception(pfHead.LastErr)
			End If
			Dim abHead(0) As Byte
			strStepName = "Head.LoadFile"
			strRet = pfHead.LoadFile
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameHead & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "Body Get Bytes"
			abHead = pfHead.GbMain.Main
			pfHead = Nothing
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

	Private Function mSavePigKeyValueToShareMem(ByRef NewItem As PigKeyValue) As String
		Const SUB_NAME As String = "mSavePigKeyValueToShareMem"
		Dim strStepName As String = ""
		Dim strRet As String
		Try
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
					strStepName &= "(" & NewItem.KeyName & "." & NewItem.SMNameHead & ")"
					Throw New Exception(strRet)
				End If
			End If
			strStepName = "mSaveSMBody"
			strRet = Me.mSaveSMBody(suSMHead, NewItem.SMNameBody, NewItem.BytesValue)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & "." & NewItem.SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

	Private Function mSavePigKeyValueToFile(ByRef NewItem As PigKeyValue) As String
		Const SUB_NAME As String = "mSavePigKeyValueToFile"
		Dim strStepName As String = ""
		Dim strRet As String
		Try
			strStepName = "Check CacheWorkDir"
			If Me.CacheWorkDir = "" Then
				Throw New Exception("Undefined CacheWorkDir")
			End If
			Dim suSMHead As StruSMHead
			ReDim suSMHead.ValueMD5(0)
			strStepName = "mGetStruFileHead"
			strRet = Me.mGetStruFileHead(suSMHead, NewItem.SMNameHead)
			If strRet <> "OK" Then
				With suSMHead
					.ValueType = NewItem.ValueType
					.ValueLen = NewItem.ValueLen
					.ValueMD5 = NewItem.ValueMD5Bytes
					.ExpTime = NewItem.ExpTime
				End With
				strStepName = "mSaveFileHead"
				strRet = Me.mSaveFileHead(suSMHead, NewItem.SMNameHead)
				If strRet <> "OK" Then
					strStepName &= "(" & NewItem.KeyName & "." & NewItem.SMNameHead & ")"
					Throw New Exception(strRet)
				End If
			End If
			strStepName = "mSaveFileBody"
			strRet = Me.mSaveFileBody(suSMHead, NewItem.SMNameBody, NewItem.BytesValue)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & "." & NewItem.SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

	Public Sub SavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True)
		Const SUB_NAME As String = "SavePigKeyValue"
		Dim strStepName As String = ""
		Dim strRet As String
		Try
			strStepName = "Check NewItem"
			If NewItem.LastErr <> "" Then Throw New Exception(NewItem.LastErr)
			Dim strKeyName As String = NewItem.KeyName
			Dim pkvOld As PigKeyValue = Nothing
			'获取旧的成员
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					strStepName = "mGetPigKeyValueByList"
					strRet = Me.mGetPigKeyValueByList(strKeyName, pkvOld)
					If Me.PigKeyValues.LastErr <> "" Then
						strStepName &= "(" & strKeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
					End If
				Case enmCacheLevel.ToShareMem
					strStepName = "mGetPigKeyValueByShareMem"
					strRet = Me.mGetPigKeyValueByShareMem(strKeyName, pkvOld)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
					End If
				Case enmCacheLevel.ToFile
					strStepName = "mGetPigKeyValueByFile"
					strRet = Me.mGetPigKeyValueByFile(strKeyName, pkvOld)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
					End If
				Case Else
					strStepName = strKeyName
					Throw New Exception("Unsupported CacheLevel")
			End Select
			'确定新增还是更新
			Dim bolIsNew As Boolean = False, bolUpdate As Boolean = False
			If NewItem.Parent Is Nothing Then NewItem.Parent = Me
			If pkvOld Is Nothing Then
				bolIsNew = True
			ElseIf pkvOld.CompareOther(NewItem) = False Then
				If IsOverwrite = False Then
					strStepName = strKeyName
					Throw New Exception("PigKeyValue Exists")
				End If
				bolUpdate = True
			End If

			If bolIsNew = True Then
				msuStatistics.SaveCount += 1
				Select Case Me.CacheLevel
					Case enmCacheLevel.ToList
						strStepName = "mAddPigKeyValueToList"
						strRet = Me.mAddPigKeyValueToList(NewItem)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Throw New Exception(strRet)
						End If
						msuStatistics.SaveToListCount += 1
					Case enmCacheLevel.ToShareMem
						strStepName = "mSavePigKeyValueToShareMem.New"
						strRet = Me.mSavePigKeyValueToShareMem(NewItem)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Throw New Exception(strRet)
						End If
						msuStatistics.SaveToShareMemCount += 1
					Case enmCacheLevel.ToFile
						strStepName = "mSavePigKeyValueToFile.New"
						strRet = Me.mSavePigKeyValueToFile(NewItem)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Throw New Exception(strRet)
						End If
						msuStatistics.SaveToFileCount += 1
				End Select
			ElseIf bolUpdate = True Then
				msuStatistics.SaveCount += 1
				Select Case Me.CacheLevel
					Case enmCacheLevel.ToList
						strStepName = "CopyToMe.Update.ToList"
						strRet = pkvOld.CopyToMe(NewItem)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Throw New Exception(strRet)
						End If
						msuStatistics.SaveToListCount += 1
					Case enmCacheLevel.ToShareMem
						strStepName = "mClearShareMem.Update"
						strRet = Me.mClearShareMem(strKeyName)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							msuStatistics.RemoveFailCount += 1
						End If
						strStepName = "mSavePigKeyValueToShareMem.Update"
						strRet = Me.mSavePigKeyValueToShareMem(NewItem)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Throw New Exception(strRet)
						End If
						msuStatistics.SaveToShareMemCount += 1
					Case enmCacheLevel.ToFile
						If pkvOld.Parent Is Nothing Then pkvOld.Parent = Me
						strStepName = "mClearShareMem.Update"
						strRet = Me.mRemoveFile(pkvOld)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							msuStatistics.RemoveFailCount += 1
						End If
						If NewItem.Parent Is Nothing Then NewItem.Parent = Me
						strStepName = "mSavePigKeyValueToFile.Update"
						strRet = Me.mSavePigKeyValueToFile(NewItem)
						If strRet <> "OK" Then
							strStepName &= "(" & strKeyName & ")"
							Throw New Exception(strRet)
						End If
						msuStatistics.SaveToFileCount += 1
				End Select
			Else
				pkvOld.LastRefCacheTime = Now
			End If
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToShareMem, enmCacheLevel.ToFile
					If Me.CacheLevel = enmCacheLevel.ToFile Then
						If Me.mIsShareMemExists(strKeyName) = True Then
							strStepName = "mGetPigKeyValueByShareMem"
							strRet = Me.mGetPigKeyValueByShareMem(strKeyName, pkvOld)
							If strRet <> "OK" Then
								strStepName &= "(" & strKeyName & ")"
								Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
							End If
							If Not pkvOld Is Nothing Then
								If pkvOld.Parent Is Nothing Then pkvOld.Parent = Me
								If pkvOld.CompareOther(NewItem) = False Then
									strStepName = "mClearShareMem.ToFile"
									strRet = Me.mClearShareMem(strKeyName)
									If strRet <> "OK" Then
										strStepName &= "(" & strKeyName & ")"
										Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
										msuStatistics.RemoveFailCount += 1
									End If
								End If
							End If
						End If
					End If
					If Me.PigKeyValues.IsItemExists(strKeyName) = True Then
						strStepName = "mGetPigKeyValueByList.ToShareMem.ToFile"
						strRet = Me.mGetPigKeyValueByList(strKeyName, pkvOld)
						If Me.PigKeyValues.LastErr <> "" Then
							strStepName &= "(" & strKeyName & ")"
							Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
						End If
						If Not pkvOld Is Nothing Then
							If pkvOld.Parent Is Nothing Then pkvOld.Parent = Me
							If pkvOld.CompareOther(NewItem) = False Then
								strStepName = "CopyToMe.ToShareMem.ToFile"
								strRet = pkvOld.CopyToMe(NewItem)
								If strRet <> "OK" Then
									strStepName &= "(" & strKeyName & ")"
									Throw New Exception(strRet)
								End If
							End If
						End If
					End If
			End Select
			Me.ClearErr()
		Catch ex As Exception
			msuStatistics.SaveFailCount += 1
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			Me.PrintDebugLog(SUB_NAME, "Catch Exception", Me.LastErr)
		End Try
	End Sub

	Public Function RemovePigKeyValue(KeyName As String, CacheLevel As enmCacheLevel) As String
		Const SUB_NAME As String = "RemovePigKeyValue"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			Dim bolIsToList As Boolean = False, bolIsToShareMem As Boolean = False, bolIsToFile As Boolean = False
			Select Case CacheLevel
				Case enmCacheLevel.ToList
					bolIsToList = True
				Case enmCacheLevel.ToShareMem
					bolIsToList = True
					bolIsToShareMem = True
				Case enmCacheLevel.ToFile
					bolIsToList = True
					bolIsToShareMem = True
					bolIsToFile = True
				Case Else
					strStepName = KeyName
					Throw New Exception("Unsupported CacheLevel")
			End Select
			msuStatistics.RemoveCount += 1
			Dim strErr As String = ""
			If bolIsToFile = True Then
				If Me.mIsCacheFileExists(KeyName) = True Then
					Dim oPigKeyValue As New PigKeyValue(KeyName, Now.AddSeconds(1), "")
					strStepName = "mRemoveFile"
					strRet = Me.mRemoveFile(oPigKeyValue)
					If strRet <> "OK" Then strErr &= strStepName & ":" & strRet
					oPigKeyValue = Nothing
				End If
			End If
			If bolIsToShareMem = True Then
				If Me.mIsShareMemExists(KeyName) = True Then
					strStepName = "mClearShareMem"
					strRet = Me.mClearShareMem(KeyName)
					If strRet <> "OK" Then strErr &= strStepName & ":" & strRet
				End If
			End If
			If bolIsToList = True Then
				If Me.PigKeyValues.IsItemExists(KeyName) = True Then
					strStepName = "mRemovePigKeyValueFromList"
					strRet = Me.mRemovePigKeyValueFromList(KeyName)
					If strRet <> "OK" Then strErr &= strStepName & ":" & strRet
				End If
			End If
			If strErr <> "" Then
				strStepName = "Remove(" & KeyName & ")"
				Throw New Exception(strErr)
			End If
			Return "OK"
		Catch ex As Exception
			strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Me.PrintDebugLog(SUB_NAME, "Catch Exception", strRet)
			msuStatistics.RemoveFailCount += 1
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
				Dim strRet As String
				For i = 1 To intItems
					strStepName = "RemovePigKeyValue"
					strRet = Me.RemovePigKeyValue(astrKeyName(i), Me.CacheLevel)
					If strRet <> "OK" Then
						strStepName &= "(" & astrKeyName(i) & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
					End If
				Next
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Sub

	Public Function GetStatisticsXml() As String
		Try
			Dim oPigXml As New PigXml(True)
			GetStatisticsXml = ""
			oPigXml.AddEle("PID", System.Diagnostics.Process.GetCurrentProcess.Id.ToString)
			oPigXml.AddEle("StatisticsTime", Format(Now, "yyyy-MM-dd HH:mm:ss.fff"))
			With msuStatistics
				oPigXml.AddEle("GetCount", .GetCount)
				oPigXml.AddEle("GetFailCount", .GetFailCount)
				'---------
				oPigXml.AddEle("SaveCount", .SaveCount)
				oPigXml.AddEle("SaveFailCount", .SaveFailCount)
				oPigXml.AddEle("SaveToListCount", .SaveToListCount)
				Select Case Me.CacheLevel
					Case enmCacheLevel.ToShareMem
						oPigXml.AddEle("SaveToShareMemCount", .SaveToShareMemCount)
					Case enmCacheLevel.ToFile
						oPigXml.AddEle("SaveToShareMemCount", .SaveToShareMemCount)
						oPigXml.AddEle("SaveToFileCount", .SaveToFileCount)
				End Select
				'---------
				oPigXml.AddEle("CacheCount", .CacheCount)
				oPigXml.AddEle("CacheByListCount", .CacheByListCount)
				Select Case Me.CacheLevel
					Case enmCacheLevel.ToShareMem
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
					Case enmCacheLevel.ToFile
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
						oPigXml.AddEle("CacheByFileCount", .CacheByFileCount)
				End Select
				'---------
				oPigXml.AddEle("RemoveCount", .RemoveCount)
				oPigXml.AddEle("RemoveFailCount", .RemoveFailCount)
				oPigXml.AddEle("RemoveExpiredListCount", .RemoveExpiredListCount)
				Select Case Me.CacheLevel
					Case enmCacheLevel.ToShareMem
						oPigXml.AddEle("RemoveExpiredShareMemCount", .RemoveExpiredShareMemCount)
					Case enmCacheLevel.ToFile
						oPigXml.AddEle("RemoveExpiredShareMemCount", .RemoveExpiredShareMemCount)
						oPigXml.AddEle("RemoveExpiredFileCount", .RemoveExpiredFileCount)
				End Select
			End With
			GetStatisticsXml = oPigXml.MainXmlStr
			oPigXml = Nothing
		Catch ex As Exception
			Me.SetSubErrInf("GetStatisticsXml", ex)
			Return ""
		End Try
	End Function

	Private menmCacheLevel As enmCacheLevel = enmCacheLevel.ToList
	Public Property CacheLevel As enmCacheLevel
		Get
			Return menmCacheLevel
		End Get
		Friend Set(value As enmCacheLevel)
			menmCacheLevel = value
		End Set
	End Property

	Private mintForceRefCacheTime As Integer = 60
	Public Property ForceRefCacheTime As Integer
		Get
			Return mintForceRefCacheTime
		End Get
		Friend Set(value As Integer)
			mintForceRefCacheTime = value
		End Set
	End Property

	'Private Property mLastRefCacheTime As DateTime = DateTime.MinValue

	Private mstrCacheWorkDir As String
	Public Property CacheWorkDir As String
		Get
			Return mstrCacheWorkDir
		End Get
		Friend Set(value As String)
			mstrCacheWorkDir = value
		End Set
	End Property
	Private Function mIsBytesMatch(ByRef SrcBytes As Byte(), ByRef MatchBytes As Byte()) As Boolean
		Try
#If NET40_OR_GREATER Then
			Return SrcBytes.SequenceEqual(MatchBytes)
#Else
            Dim i As Long
            If SrcBytes.Length <> MatchBytes.Length Then
                Return False
            Else
                mIsBytesMatch = True
                For i = 0 To SrcBytes.Length - 1
                    If SrcBytes(i) <> MatchBytes(i) Then
                        mIsBytesMatch = False
                        Exit For
                    End If
                Next
            End If

#End If
		Catch ex As Exception
			Me.SetSubErrInf("mIsBytesMatch", ex)
			Return False
		End Try
	End Function

	Private Function mGetStruFileHead(ByRef SuSMHead As StruSMHead, SMNameHead As String) As String
		Const SUB_NAME As String = "mGetStruFileHead"
		Dim strStepName As String = ""
		Try
			Dim strSMNameHeadFilePath As String = Me.CacheWorkDir & Me.OsPathSep & SMNameHead
			Dim abHead(0) As Byte
			strStepName = "New PigFile"
			Dim oPigFile As New PigFile(strSMNameHeadFilePath)
			If oPigFile.LastErr <> "" Then
				strStepName &= "(" & strSMNameHeadFilePath & ")"
				Throw New Exception(oPigFile.LastErr)
			End If
			strStepName = "LoadFile"
			oPigFile.LoadFile()
			If oPigFile.LastErr <> "" Then
				strStepName &= "(" & strSMNameHeadFilePath & ")"
				Throw New Exception(oPigFile.LastErr)
			End If
			ReDim SuSMHead.ValueMD5(15)
			With oPigFile.GbMain
				SuSMHead.ValueType = .GetInt32Value()
				SuSMHead.ExpTime = .GetDateTimeValue
				SuSMHead.ValueLen = .GetInt64Value
				SuSMHead.ValueMD5 = .GetBytesValue(16)
			End With
			oPigFile = Nothing
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
			Return strRet
		End Try
	End Function

	Private Function mSaveFileHead(SuSMHead As StruSMHead, SMNameHead As String) As String
		Const SUB_NAME As String = "mSaveFileHead"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim strSMNameHeadFilePath As String = Me.CacheWorkDir & Me.OsPathSep & SMNameHead
			strStepName = "New PigFile"
			Dim oPigFile As New PigFile(strSMNameHeadFilePath)
			If oPigFile.LastErr <> "" Then
				strStepName &= "(" & strSMNameHeadFilePath & ")"
				Throw New Exception(oPigFile.LastErr)
			End If
			strStepName = "New GbMain"
			oPigFile.GbMain = New PigBytes
			strStepName = "SetValue"
			With oPigFile.GbMain
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
			strStepName = "SaveFile"
			strRet = oPigFile.SaveFile(False)
			If strRet <> "OK" Then
				strStepName &= "(" & strSMNameHeadFilePath & ")"
				Throw New Exception(strRet)
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

	Private Function mSaveFileBody(SuSMHead As StruSMHead, SMNameBody As String, ByRef DataBytes As Byte()) As String
		Const SUB_NAME As String = "mSaveFileBody"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim strSMNameBodyFilePath As String = Me.CacheWorkDir & Me.OsPathSep & SMNameBody
			strStepName = "New PigFile"
			Dim oPigFile As New PigFile(strSMNameBodyFilePath)
			If oPigFile.LastErr <> "" Then
				strStepName &= "(" & strSMNameBodyFilePath & ")"
				Throw New Exception(oPigFile.LastErr)
			End If
			strStepName = "New GbMain"
			oPigFile.GbMain = New PigBytes(DataBytes)
			strStepName = "SaveFile"
			strRet = oPigFile.SaveFile(False)
			If strRet <> "OK" Then
				strStepName &= "(" & strSMNameBodyFilePath & ")"
				Throw New Exception(strRet)
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			'Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", strRet)
			Return strRet
		End Try
	End Function

	Private Function mGetPigKeyValueFromFile(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Const SUB_NAME As String = "mGetPigKeyValueFromFile"
		Dim strStepName As String = ""
		Dim strRet As String
		Try
			strStepName = "New PigKeyValue Get suSMHead"
			OutPigKeyValue = New PigKeyValue(KeyName, Now.AddMinutes(1), "")
			OutPigKeyValue.Parent = Me
			Dim suSMHead As StruSMHead
			ReDim suSMHead.ValueMD5(0)
			strStepName = "mGetStruFileHead"
			strRet = Me.mGetStruFileHead(suSMHead, OutPigKeyValue.SMNameHead)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")" : Throw New Exception(strRet)
			End If
			Dim abSMBody As Byte()
			ReDim abSMBody(0)
			strStepName = "mGetBytesSMBody"
			strRet = Me.mGetBytesFileBody(abSMBody, suSMHead, OutPigKeyValue.SMNameBody, Me.CacheWorkDir)
			If strRet = "OK" Then
				If abSMBody.Length <> suSMHead.ValueLen Then
					strRet = "SMBody.Length<>SuSMHead.ValueLen"
				Else
					Dim oPigBytes As New PigBytes(abSMBody)
					If Me.mIsBytesMatch(oPigBytes.PigMD5Bytes, suSMHead.ValueMD5) = False Then
						strRet = "SMBody.PigMD5<>SuSMHead.ValueMD5"
					End If
					oPigBytes = Nothing
				End If
			End If
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")" : Throw New Exception(strRet)
			End If
			OutPigKeyValue = Nothing
			strStepName = "New PigKeyValue To Out"
			OutPigKeyValue = New PigKeyValue(KeyName, suSMHead.ExpTime, abSMBody, suSMHead.ValueType, suSMHead.ValueMD5)
			If OutPigKeyValue.LastErr <> "" Then
				strStepName &= strStepName & "(" & KeyName & ")"
				Throw New Exception(OutPigKeyValue.LastErr)
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

	Public Function mRemoveFile(ByRef SrcItem As PigKeyValue) As String
		Dim strStepName As String = ""
		Try
			Dim strDelFile As String = Me.CacheWorkDir & Me.OsPathSep & SrcItem.SMNameBody
			If System.IO.File.Exists(strDelFile) = True Then
				strStepName = "Delete" & strDelFile
				System.IO.File.Delete(strDelFile)
			End If
			strDelFile = Me.CacheWorkDir & Me.OsPathSep & SrcItem.SMNameHead
			If System.IO.File.Exists(strDelFile) = True Then
				strStepName = "Delete" & strDelFile
				System.IO.File.Delete(strDelFile)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("mRemoveFile", strStepName, ex)
		End Try
	End Function

End Class
