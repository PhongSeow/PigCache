'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用|Piggy key value application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 2.5
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
'* 1.8		21/10/2021 Modify mIsCacheFileExists,SavePigKeyValue,mIsCacheFileExists,GetPigKeyValue,mSavePigKeyValueToShareMem
'* 1.9		24/10/2021 Modify SavePigKeyValue,fSaveValueLen,mGetPigKeyValueByFile,mGetPigKeyValueFromFile,StruSMHead
'* 2.0		28/10/2021 Add fGetSMNameHeadAndBody,mIsCacheFileExists
'* 2.1		30/11/2021 Modify mGetStruSMHead,mGetPigKeyValueFromFile, remove mGetStruFileHead
'* 2.2		1/12/2021 Modify mSaveSMHead,mGetPigKeyValueFromFile,mSaveFileHead,mGetStruSMHead,mGetPigKeyValueFromShareMem, remove mGetSMNamePart
'* 2.3		2/12/2021 Imports System.IO, Add mChkSaveBodyBytes, Modify mSavePigKeyValueToFile,mGetStruSMHead
'* 2.5		5/12/2021 Add new SavePigKeyValue,DefaultSaveType,DefaultTextType
'************************************

Imports PigToolsLiteLib
Imports System.IO

Public Class PigKeyValueApp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "2.5.6"
	Private moPigFunc As New PigFunc

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
	Public Structure StruSMHead
		Dim ValueType As PigKeyValue.enmValueType
		Dim ExpTime As DateTime
		Dim ValueLen As Long
		Dim SaveValueMD5 As Byte()
		Dim TextType As PigText.enmTextType
		Dim SaveType As PigKeyValue.enmSaveType
		Dim SaveValueLen As Long
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
		Dim strRet As String = ""
		Try
			Dim oPigKeyValue As New PigKeyValue(KeyName)
			If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
			Dim strSMNameHeadPath As String = Me.CacheWorkDir & Me.OsPathSep & oPigKeyValue.fSMNameHead
			If IO.File.Exists(strSMNameHeadPath) = False Then
				strStepName = "Check " & strSMNameHeadPath
				Throw New Exception("Non-existent")
			End If
			Dim SuSMHead As StruSMHead
			ReDim SuSMHead.SaveValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(SuSMHead, oPigKeyValue.fSMNameHead, Me.CacheWorkDir)
			If strRet <> "OK" Then Throw New Exception(strRet)
			Dim strSMNameBodyPath As String = Me.CacheWorkDir & Me.OsPathSep & oPigKeyValue.fSMNameBody
			Dim lngFileSize As Long
			strStepName = "mGetFileLen"
			strRet = Me.mGetFileLen(strSMNameBodyPath, lngFileSize)
			If strRet <> "OK" Then
				strStepName &= "(" & strSMNameBodyPath & ")"
				Throw New Exception(strRet)
			End If
			If SuSMHead.SaveValueLen <> lngFileSize Then
				strStepName &= "(" & strSMNameBodyPath & ")"
				Throw New Exception("Length mismatch")
			End If
			Return "OK"
		Catch ex As Exception
			strRet = Me.GetSubErrInf("mIsCacheFileExists", strStepName, ex)
			Me.PrintDebugLog("As Exception", strRet)
			Return False
		End Try
	End Function

	Private Function mGetPigKeyValueFromShareMem(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Const SUB_NAME As String = "mGetPigKeyValueFromShareMem"
		Dim strStepName As String = ""
		Dim strRet As String
		Try

			If OutPigKeyValue Is Nothing Then
				strStepName = "New PigKeyValue"
				OutPigKeyValue = New PigKeyValue(KeyName)
				If OutPigKeyValue.LastErr <> "" Then
					strStepName &= "(" & KeyName & ")"
					Throw New Exception(OutPigKeyValue.LastErr)
				End If
			End If
			Dim suSMHead As StruSMHead
			ReDim suSMHead.SaveValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(suSMHead, OutPigKeyValue.fSMNameHead)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(strRet)
			End If

			Dim abSMBody As Byte()
			ReDim abSMBody(0)
			strStepName = "mGetBytesSMBody"
			strRet = Me.mGetBytesSMBody(abSMBody, suSMHead, OutPigKeyValue.fSMNameBody)
			If strRet = "OK" Then
				If abSMBody.Length <> suSMHead.SaveValueLen Then
					strRet = "SMBody.Length<>SuSMHead.SaveValueLen"
				Else
					strStepName &= ",mChkSaveBodyBytes"
					strRet = Me.mChkSaveBodyBytes(abSMBody, suSMHead)
					'Dim oPigBytes As New PigBytes(abSMBody)
					'If Me.mIsBytesMatch(oPigBytes.PigMD5Bytes, suSMHead.SaveValueMD5) = False Then
					'	strRet = "SMBody.PigMD5<>SuSMHead.ValueMD5"
					'End If
					'oPigBytes = Nothing
				End If
			End If
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "InitBytesBySave"
			strRet = OutPigKeyValue.InitBytesBySave(suSMHead, abSMBody)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(strRet)
			End If
			Return "OK"
		Catch ex As Exception
			OutPigKeyValue = Nothing
			strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function


	Private Function mClearShareMem(KeyName As String) As String
		Const SUB_NAME As String = "mClearShareMem"
		Dim strStepName As String = "", strRet As String = ""
		Try
			Dim oPigKeyValue As New PigKeyValue(KeyName)
			If oPigKeyValue.LastErr <> "" Then
				Throw New Exception(oPigKeyValue.LastErr)
			End If
			Dim SuSMHead As StruSMHead
			ReDim SuSMHead.SaveValueMD5(15)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(SuSMHead, oPigKeyValue.fSMNameHead)
			If strRet = "OK" Then
				Dim intBodyLen As Integer = SuSMHead.SaveValueLen
				With SuSMHead
					.ExpTime = DateTime.MinValue
					.SaveValueLen = 0
					.ValueType = PigKeyValue.enmValueType.Unknow
					ReDim .SaveValueMD5(15)
					.TextType = PigText.enmTextType.UnknowOrBin
					.SaveType = PigKeyValue.enmSaveType.Original
				End With
				strStepName = "mSaveSMBody"
				strRet = Me.mSaveSMHead(SuSMHead, oPigKeyValue.fSMNameHead)
				If strRet <> "OK" Then Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
				Dim abBody As Byte()
				ReDim abBody(intBodyLen - 1)
				strStepName = "mSaveSMBody"
				strRet = Me.mSaveSMBody(SuSMHead, oPigKeyValue.fSMNameBody, abBody)
				If strRet <> "OK" Then
					strStepName &= "(" & KeyName & ")"
					Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
				End If
			Else
				strStepName &= "(" & KeyName & ")"
				Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
			End If
			oPigKeyValue = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("mClearShareMem", strStepName, ex)
		End Try
	End Function

	Private Function mIsShareMemExists(KeyName As String) As Boolean
		Const SUB_NAME As String = "mIsShareMemExists"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			Dim oPigKeyValue As New PigKeyValue(KeyName)
			If oPigKeyValue.LastErr <> "" Then Throw New Exception(oPigKeyValue.LastErr)
			Dim SuSMHead As StruSMHead
			ReDim SuSMHead.SaveValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(SuSMHead, oPigKeyValue.fSMNameHead)
			If strRet <> "OK" Then
				Return False
			Else
				Dim abSMBody As Byte()
				ReDim abSMBody(0)
				strStepName = "mGetBytesSMBody"
				strRet = Me.mGetBytesSMBody(abSMBody, SuSMHead, oPigKeyValue.fSMNameBody)
				If strRet <> "OK" Then
					Return False
				ElseIf abSMBody.Length <> SuSMHead.SaveValueLen Then
					Return False
				Else
					strStepName &= ",mChkSaveBodyBytes"
					strRet = Me.mChkSaveBodyBytes(abSMBody, SuSMHead)
					If strRet <> "OK" Then
						mIsShareMemExists = False
					Else
						mIsShareMemExists = True
					End If
					'Dim oPigBytes As New PigBytes(abSMBody)
					'If Me.mIsBytesMatch(oPigBytes.PigMD5Bytes, SuSMHead.SaveValueMD5) = False Then
					'	mIsShareMemExists = False
					'Else
					'	mIsShareMemExists = True
					'End If
					'oPigBytes = Nothing
				End If
			End If
		Catch ex As Exception
			strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Me.PrintDebugLog(SUB_NAME, strRet)
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
						If pkvList.fIsForceRefCache = True Then
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
							If GetPigKeyValue.fCompareOther(pkvList) = False Then
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
						If pkvList.fIsForceRefCache = True Then
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
									If pkvShareMem.fCompareOther(GetPigKeyValue) = False Then
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
								msuStatistics.SaveToShareMemCount += 1
								strStepName = "mSavePigKeyValueToShareMem.ToFile2"
								strRet = Me.mSavePigKeyValueToShareMem(GetPigKeyValue)
								If strRet <> "OK" Then
									strStepName &= "(" & KeyName & ")"
									Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
								End If
							End If
						Else
							If pkvShareMem.Parent Is Nothing Then pkvShareMem.Parent = Me
							msuStatistics.SaveToListCount += 1
							strStepName = "mAddPigKeyValueToList"
							strRet = Me.mAddPigKeyValueToList(pkvShareMem)
							If strRet <> "OK" Then
								strStepName &= "(" & KeyName & ")"
								Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
								msuStatistics.SaveFailCount += 1
							End If
							GetPigKeyValue = pkvShareMem
							pkvList = Nothing
						End If
					Else
						GetPigKeyValue = pkvList
						pkvList = Nothing
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


	Private Function mGetBytesSMBody(ByRef BodyBytes As Byte(), SuSMHead As StruSMHead, SMNameBody As String) As String
		Const SUB_NAME As String = "mGetBytesSMBody"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			Dim msmBody As New ShareMem
			strStepName = "Body.Init"
			strRet = msmBody.Init(SMNameBody, SuSMHead.SaveValueLen)
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
			If SuSMHead.SaveValueLen <> BodyBytes.Length Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception("ValueLen not match " & SuSMHead.SaveValueLen & "," & BodyBytes.Length)
			End If
			strStepName = "mChkSaveBodyBytes"
			strRet = Me.mChkSaveBodyBytes(BodyBytes, SuSMHead)
			If strRet <> "OK" Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception(strRet)
			End If
			'Dim oPigMD5 As New PigMD5(BodyBytes)
			'If oPigMD5.LastErr <> "" Then
			'	strStepName &= "(" & SMNameBody & ")"
			'	Throw New Exception(oPigMD5.LastErr)
			'End If
			'If Me.mIsBytesMatch(pbBody.PigMD5Bytes, SuSMHead.SaveValueMD5) = False Then
			'	strStepName &= "(" & SMNameBody & ")"
			'	Throw New Exception("PigMD5 not match")
			'End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

	Private Function mChkSaveBodyBytes(ByRef SaveBodyBytes As Byte(), SuSMHead As StruSMHead) As String
		Const SUB_NAME As String = "mChkSaveBodyBytes"
		Dim strStepName As String = ""
		Try
			If SaveBodyBytes.Length <> SuSMHead.SaveValueLen Then
				strStepName = "Check Length"
				Throw New Exception("Mismatch")
			End If
			strStepName = "New PigMD5"
			Dim oPigMD5 As New PigMD5(SaveBodyBytes)
			If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
			If Me.mIsBytesMatch(oPigMD5.PigMD5Bytes, SuSMHead.SaveValueMD5) = False Then
				strStepName = "Check PigMD5"
				Throw New Exception("Mismatch")
			End If
			oPigMD5 = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
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
			If SuSMHead.SaveValueLen <> BodyBytes.Length Then
				strStepName &= "(" & SMNameBody & ")"
				Throw New Exception("ValueLen not match " & SuSMHead.SaveValueLen & "," & BodyBytes.Length)
			End If
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

	''' <summary>
	''' GetStruSMHead from file
	''' </summary>
	''' <param name="SuSMHead"></param>
	''' <param name="SMNameHead"></param>
	''' <param name="CacheWorkDir"></param>
	''' <returns></returns>
	Private Function mGetStruSMHead(ByRef SuSMHead As StruSMHead, SMNameHead As String, CacheWorkDir As String) As String
		Const SUB_NAME As String = "mGetStruSMHead"
		Dim strStepName As String = ""
		Try
			Dim strRet As String
			If CacheWorkDir = "" Then
				strStepName = "Check CacheWorkDir"
				Throw New Exception("Undefined")
			End If
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
			abHead = pfHead.GbMain.Main
			pfHead = Nothing
			strStepName = "New PigBytes(abHead)"
			Dim pbHead As New PigBytes(abHead)
			If pbHead.LastErr <> "" Then
				strStepName &= "(abHead.Length=" & abHead.Length & ")"
				Throw New Exception(strRet)
			End If
			ReDim SuSMHead.SaveValueMD5(15)
			With pbHead
				SuSMHead.ValueType = .GetInt32Value()
				SuSMHead.ExpTime = .GetDateTimeValue
				SuSMHead.SaveValueLen = .GetInt64Value
				SuSMHead.SaveValueMD5 = .GetBytesValue(16)
				SuSMHead.TextType = .GetInt32Value()
				SuSMHead.SaveType = .GetInt32Value()
				SuSMHead.SaveValueLen = .GetInt32Value()
			End With
			strStepName = "Check StruSMHead (" & SMNameHead & ")"
			With SuSMHead
				If .SaveValueLen < 0 Then Throw New Exception("Data length is less than or equal to 0")
				If .ExpTime < Now Then Throw New Exception("Data expired ")
				Select Case .ValueType
					Case PigKeyValue.enmValueType.Bytes
					Case PigKeyValue.enmValueType.Text
						Select Case .TextType
							Case PigText.enmTextType.Ascii, PigText.enmTextType.Unicode, PigText.enmTextType.UTF8
							Case Else
								Throw New Exception("Invalid TextType is " & .TextType)
						End Select
					Case Else
						Throw New Exception("Invalid ValueType is " & .ValueType)
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
			strRet = msmHead.Init(SMNameHead, 52)
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
			ReDim SuSMHead.SaveValueMD5(15)
			With pbHead
				SuSMHead.ValueType = .GetInt32Value()
				SuSMHead.ExpTime = .GetDateTimeValue
				SuSMHead.SaveValueLen = .GetInt64Value
				SuSMHead.SaveValueMD5 = .GetBytesValue(16)
				SuSMHead.TextType = .GetInt32Value()
				SuSMHead.SaveType = .GetInt32Value()
				SuSMHead.SaveValueLen = .GetInt32Value()
			End With
			strStepName = "Check StruSMHead (" & SMNameHead & ")"
			With SuSMHead
				If .SaveValueLen < 0 Then Throw New Exception("Data length is less than or equal to 0")
				If .ExpTime < Now Then Throw New Exception("Data expired ")
				Select Case .ValueType
					Case PigKeyValue.enmValueType.Bytes
					Case PigKeyValue.enmValueType.Text
						Select Case .TextType
							Case PigText.enmTextType.Ascii, PigText.enmTextType.Unicode, PigText.enmTextType.UTF8
							Case Else
								Throw New Exception("Invalid TextType is " & .TextType)
						End Select
					Case Else
						Throw New Exception("Invalid ValueType is " & .ValueType)
				End Select
			End With
			Return "OK"
		Catch ex As Exception
			Dim strRet As String = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
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
			strRet = msmHead.Init(SMNameHead, 52)
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
				.SetValue(SuSMHead.SaveValueMD5)
				If .LastErr <> "" Then
					strStepName &= ".SaveValueMD5"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.TextType)
				If .LastErr <> "" Then
					strStepName &= ".TextType"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.SaveType)
				If .LastErr <> "" Then
					strStepName &= ".SaveType"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.SaveValueLen)
				If .LastErr <> "" Then
					strStepName &= ".SaveValueLen"
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
			strRet = msmBody.Init(SMNameBody, SuSMHead.SaveValueLen)
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
			ReDim suSMHead.SaveValueMD5(0)
			Dim abSaveBytes As Byte(), abSavePigMD5 As Byte()
			ReDim abSaveBytes(0)
			ReDim abSavePigMD5(0)
			strStepName = "GetSaveData"
			strRet = NewItem.GetSaveData(abSaveBytes, abSavePigMD5)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & ")"
				Throw New Exception(strRet)
			End If
			With suSMHead
				.ValueType = NewItem.ValueType
				.SaveValueLen = abSaveBytes.Length
				.SaveValueMD5 = abSavePigMD5
				.ExpTime = NewItem.ExpTime
				.TextType = NewItem.TextType
				.SaveType = NewItem.SaveType
				.ValueLen = NewItem.ValueLen
			End With
			strStepName = "mSaveSMHead"
			strRet = Me.mSaveSMHead(suSMHead, NewItem.fSMNameHead)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & "." & NewItem.fSMNameHead & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "mSaveSMBody"
			strRet = Me.mSaveSMBody(suSMHead, NewItem.fSMNameBody, abSaveBytes)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & "." & NewItem.fSMNameBody & ")"
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
			Dim suSMHead As StruSMHead
			ReDim suSMHead.SaveValueMD5(0)
			Dim abSaveBytes As Byte(), abSavePigMD5 As Byte()
			ReDim abSaveBytes(0)
			ReDim abSavePigMD5(0)
			strStepName = "GetSaveData"
			strRet = NewItem.GetSaveData(abSaveBytes, abSavePigMD5)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & ")"
				Throw New Exception(strRet)
			End If
			With suSMHead
				.ValueType = NewItem.ValueType
				.SaveValueLen = abSaveBytes.Length
				.SaveValueMD5 = abSavePigMD5
				.ExpTime = NewItem.ExpTime
				.TextType = NewItem.TextType
				.SaveType = NewItem.SaveType
				.ValueLen = NewItem.ValueLen
			End With
			strStepName = "mSaveFileHead"
			strRet = Me.mSaveFileHead(suSMHead, NewItem.fSMNameHead)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & "." & NewItem.fSMNameHead & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "mSaveFileBody"
			strRet = Me.mSaveFileBody(suSMHead, NewItem.fSMNameBody, abSaveBytes)
			If strRet <> "OK" Then
				strStepName &= "(" & NewItem.KeyName & "." & NewItem.fSMNameBody & ")"
				Throw New Exception(strRet)
			End If
			'Dim suSMHead As StruSMHead
			'ReDim suSMHead.SaveValueMD5(0)
			'strStepName = "mGetStruSMHead"
			'strRet = Me.mGetStruSMHead(suSMHead, NewItem.fSMNameHead, Me.CacheWorkDir)
			'If strRet <> "OK" Then
			'	Select Case NewItem.ValueType
			'		Case PigKeyValue.enmValueType.Text, PigKeyValue.enmValueType.Bytes
			'			Dim abSaveBytes As Byte(), abSavePigMD5 As Byte()
			'			ReDim abSaveBytes(0)
			'			ReDim abSavePigMD5(0)
			'			strStepName = "GetSaveData"
			'			strRet = NewItem.GetSaveData(abSaveBytes, abSavePigMD5)
			'			If strRet <> "OK" Then Throw New Exception(strRet)
			'			With suSMHead
			'				.ValueType = NewItem.ValueType
			'				.SaveValueLen = NewItem.ValueLen
			'				.SaveValueMD5 = abSavePigMD5
			'				.ExpTime = NewItem.ExpTime
			'				.TextType = NewItem.TextType
			'				.SaveType = NewItem.SaveType
			'				.SaveValueLen = abSaveBytes.Length
			'			End With
			'			strStepName = "mSaveFileHead"
			'			strRet = Me.mSaveFileHead(suSMHead, NewItem.fSMNameHead)
			'			If strRet <> "OK" Then
			'				strStepName &= "(" & NewItem.KeyName & "." & NewItem.fSMNameHead & ")"
			'				Throw New Exception(strRet)
			'			End If
			'			strStepName = "mSaveFileBody"
			'			strRet = Me.mSaveFileBody(suSMHead, NewItem.fSMNameBody, abSaveBytes)
			'			If strRet <> "OK" Then
			'				strStepName &= "(" & NewItem.KeyName & "." & NewItem.fSMNameBody & ")"
			'				Throw New Exception(strRet)
			'			End If
			'		Case Else
			'			strStepName = "ValueType is " & NewItem.ValueType.ToString
			'			Throw New Exception("Unsupported")
			'	End Select
			'End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

	Public Sub SavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True)
		Try
			Dim strRet As String = Me.mSavePigKeyValue(NewItem, IsOverwrite)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("SavePigKeyValue", ex)
		End Try
	End Sub

	Public Function SavePigKeyValue(KeyName As String, KeyValue As Byte(), ExpTimeSec As Long, Optional IsOverwrite As Boolean = True) As String
		Dim strStepName As String = "SavePigKeyValue"
		Dim strRet As String = ""
		Try
			strStepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, Now.AddSeconds(ExpTimeSec), KeyValue, Me.DefaultSaveType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			strStepName = "mSavePigKeyValue"
			strRet = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If strRet <> "OK" Then Throw New Exception(strRet)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", strStepName, ex)
		End Try
	End Function

	Public Function SavePigKeyValue(KeyName As String, KeyValue As Byte(), ExpTime As DateTime, Optional IsOverwrite As Boolean = True) As String
		Dim strStepName As String = "SavePigKeyValue"
		Dim strRet As String = ""
		Try
			strStepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, ExpTime, KeyValue, Me.DefaultSaveType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			strStepName = "mSavePigKeyValue"
			strRet = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If strRet <> "OK" Then Throw New Exception(strRet)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", strStepName, ex)
		End Try
	End Function

	Public Function SavePigKeyValue(KeyName As String, KeyValue As String, ExpTime As DateTime, Optional IsOverwrite As Boolean = True) As String
		Dim strStepName As String = "SavePigKeyValue"
		Dim strRet As String = ""
		Try
			strStepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, ExpTime, KeyValue, Me.DefaultTextType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			strStepName = "mSavePigKeyValue"
			strRet = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If strRet <> "OK" Then Throw New Exception(strRet)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", strStepName, ex)
		End Try
	End Function

	Public Function SavePigKeyValue(KeyName As String, KeyValue As String, ExpTimeSec As Long, Optional IsOverwrite As Boolean = True) As String
		Dim strStepName As String = "SavePigKeyValue"
		Dim strRet As String = ""
		Try
			strStepName = "New PigKeyValue"
			Dim oNewItem As New PigKeyValue(KeyName, Now.AddSeconds(ExpTimeSec), KeyValue, Me.DefaultTextType)
			If oNewItem.LastErr <> "" Then Throw New Exception(oNewItem.LastErr)
			strStepName = "mSavePigKeyValue"
			strRet = Me.mSavePigKeyValue(oNewItem, IsOverwrite)
			If strRet <> "OK" Then Throw New Exception(strRet)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("SavePigKeyValue", strStepName, ex)
		End Try
	End Function

	Private Function mSavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True) As String
		Const SUB_NAME As String = "mSavePigKeyValue"
		Dim strStepName As String = ""
		Dim strRet As String
		Try
			strStepName = "NewItem.Check"
			strRet = NewItem.Check()
			If strRet <> "OK" Then Throw New Exception(strRet)
			Dim strKeyName As String = NewItem.KeyName
			Dim pkvOld As PigKeyValue = Nothing
			'获取旧的成员
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					strStepName = "mGetPigKeyValueByList"
					strRet = Me.mGetPigKeyValueByList(strKeyName, pkvOld)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
					End If
				Case enmCacheLevel.ToShareMem
					strStepName = "mGetPigKeyValueByShareMem"
					strRet = Me.mGetPigKeyValueByShareMem(strKeyName, pkvOld)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
					End If
				Case enmCacheLevel.ToFile
					strStepName = "mGetPigKeyValueByFile"
					strRet = Me.mGetPigKeyValueByFile(strKeyName, pkvOld)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
					End If
				Case Else
					strStepName = Me.CacheLevel.ToString
					Throw New Exception("Unsupported CacheLevel")
			End Select
			'确定新增还是更新
			Dim bolIsNew As Boolean = False, bolUpdate As Boolean = False
			If NewItem.Parent Is Nothing Then NewItem.Parent = Me
			If pkvOld Is Nothing Then
				bolIsNew = True
			ElseIf pkvOld.fCompareOther(NewItem) = False Then
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
						strRet = pkvOld.fCopyToMe(NewItem)
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
								If pkvOld.fCompareOther(NewItem) = False Then
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
							If pkvOld.fCompareOther(NewItem) = False Then
								strStepName = "CopyToMe.ToShareMem.ToFile"
								strRet = pkvOld.fCopyToMe(NewItem)
								If strRet <> "OK" Then
									strStepName &= "(" & strKeyName & ")"
									Throw New Exception(strRet)
								End If
							End If
						End If
					End If
			End Select
			Return "OK"
		Catch ex As Exception
			msuStatistics.SaveFailCount += 1
			strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
			Return strRet
		End Try
	End Function

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
					Dim oPigKeyValue As New PigKeyValue(KeyName)
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

	Private mintDefaultSaveType As PigKeyValue.enmSaveType = PigKeyValue.enmSaveType.SaveSpace
	Public Property DefaultSaveType As PigKeyValue.enmSaveType
		Get
			Return mintDefaultSaveType
		End Get
		Friend Set(value As PigKeyValue.enmSaveType)
			mintDefaultSaveType = value
		End Set
	End Property

	Private mintDefaultTextType As PigText.enmTextType = PigText.enmTextType.UTF8
	Public Property DefaultTextType As PigText.enmTextType
		Get
			Return mintDefaultTextType
		End Get
		Friend Set(value As PigText.enmTextType)
			mintDefaultTextType = value
		End Set
	End Property

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
				.SetValue(SuSMHead.SaveValueMD5)
				If .LastErr <> "" Then
					strStepName &= ".SaveValueMD5"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.TextType)
				If .LastErr <> "" Then
					strStepName &= ".TextType"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.SaveType)
				If .LastErr <> "" Then
					strStepName &= ".SaveType"
					Throw New Exception(.LastErr)
				End If
				.SetValue(SuSMHead.SaveValueLen)
				If .LastErr <> "" Then
					strStepName &= ".SaveValueLen"
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
			If OutPigKeyValue Is Nothing Then
				strStepName = "New PigKeyValue"
				OutPigKeyValue = New PigKeyValue(KeyName)
				If OutPigKeyValue.LastErr <> "" Then
					strStepName &= "(" & KeyName & ")"
					Throw New Exception(OutPigKeyValue.LastErr)
				End If
			End If
			Dim suSMHead As StruSMHead
			ReDim suSMHead.SaveValueMD5(0)
			strStepName = "mGetStruSMHead"
			strRet = Me.mGetStruSMHead(suSMHead, OutPigKeyValue.fSMNameHead, Me.CacheWorkDir)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(strRet)
			End If
			Dim abSMBody As Byte()
			ReDim abSMBody(0)
			strStepName = "mGetBytesFileBody"
			strRet = Me.mGetBytesFileBody(abSMBody, suSMHead, OutPigKeyValue.fSMNameBody, Me.CacheWorkDir)
			If strRet = "OK" Then
				If abSMBody.Length <> suSMHead.SaveValueLen Then
					strRet = "The imported data length does not match." & "(" & abSMBody.Length & "," & suSMHead.SaveValueLen & ")"
				Else
					strStepName &= ",mChkSaveBodyBytes"
					strRet = Me.mChkSaveBodyBytes(abSMBody, suSMHead)
					'Dim oPigBytes As New PigBytes(abSMBody)
					'If Me.mIsBytesMatch(oPigBytes.PigMD5Bytes, suSMHead.ValueMD5) = False Then
					'	strRet = "The imported data does not match PigMD5"
					'End If
					'oPigBytes = Nothing
				End If
			End If
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(strRet)
			End If
			strStepName = "InitBytesBySave"
			strRet = OutPigKeyValue.InitBytesBySave(suSMHead, abSMBody)
			If strRet <> "OK" Then
				strStepName &= "(" & KeyName & ")"
				Throw New Exception(strRet)
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
			Dim strDelFile As String = Me.CacheWorkDir & Me.OsPathSep & SrcItem.fSMNameBody
			If IO.File.Exists(strDelFile) = True Then
				strStepName = "Delete" & strDelFile
				IO.File.Delete(strDelFile)
			End If
			strDelFile = Me.CacheWorkDir & Me.OsPathSep & SrcItem.fSMNameHead
			If IO.File.Exists(strDelFile) = True Then
				strStepName = "Delete" & strDelFile
				IO.File.Delete(strDelFile)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("mRemoveFile", strStepName, ex)
		End Try
	End Function

	Private Function mGetFileLen(FilePath As String, ByRef FileLen As Long) As String
		Dim strStepName As String = ""
		Try
			strStepName = "New FileInfo"
			Dim oFileInfo As New FileInfo(FilePath)
			FileLen = oFileInfo.Length
			oFileInfo = Nothing
			Return "OK"
		Catch ex As Exception
			strStepName &= "(" & FilePath & ")"
			Return Me.GetSubErrInf("mGetFileLen", strStepName, ex)
		End Try
	End Function



End Class
