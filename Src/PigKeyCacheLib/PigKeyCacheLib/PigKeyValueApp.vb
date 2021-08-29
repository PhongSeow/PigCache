'**********************************
'* Name: PigKeyValueApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 豚豚键值应用|Piggy key value application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.19
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
'************************************

Imports PigToolsLiteLib

Public Class PigKeyValueApp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.1.1"

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
		''' <summary>
		''' It is suitable for multi server, multi process and multi-threaded programs, and has the highest requirements for the availability of cached content, but the writing performance is poor, but the advantage is that it can share the database with the application to reduce the point of failure.
		''' </summary>
		ToDB = 40
		''' <summary>
		''' It is suitable for multi server, multi process and multi thread programs. The read and write performance is very good, but redis needs to be installed, which needs to increase the cost of managing the high availability of redis.
		''' </summary>
		ToRedis = 50
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
		Dim SaveToShareMemCount As Long
		Dim SaveToFileCount As Long
		Dim SaveToDBCount As Long
		Dim SaveToRedisCount As Long
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

	Public Sub New()
		MyBase.New(CLS_VERSION)
		mNew()
	End Sub

	Private Sub mNew(Optional ShareMemRoot As String = "", Optional CacheLevel As enmCacheLevel = enmCacheLevel.ToShareMem, Optional ForceRefCacheTime As Integer = 60, Optional CacheWorkDir As String = "")
		Try
			If ShareMemRoot = "" Then ShareMemRoot = Me.AppTitle
			Dim oPigMD5 As New PigMD5(ShareMemRoot, PigMD5.enmTextType.UTF8)
			Me.ShareMemRoot = oPigMD5.PigMD5()
			Select Case CacheLevel
				Case enmCacheLevel.ToList, enmCacheLevel.ToShareMem
					Me.CacheLevel = CacheLevel
				Case enmCacheLevel.ToFile
					Me.CacheLevel = CacheLevel
					Me.CacheWorkDir = CacheWorkDir
				Case Else
					Throw New Exception("Currently unsupported cachelevel")
			End Select
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

	Public Sub New(ShareMemRoot As String, CacheLevel As enmCacheLevel)
		MyBase.New(CLS_VERSION)
		Me.mNew(ShareMemRoot, CacheLevel)
	End Sub

	Public Sub New(ShareMemRoot As String, CacheWorkDir As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(ShareMemRoot, enmCacheLevel.ToFile,, CacheWorkDir)
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
				strRet = Me.mGetBytesSMBody(abSMBody, SuSMHead, pkvNew.SMNameBody, Me.CacheWorkDir)
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

	Private Function mGetPigKeyValueByShareMem(KeyName As String, ByRef OutPigKeyValue As PigKeyValue) As String
		Const SUB_NAME As String = "mGetPigKeyValueByShareMem"
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

	Private Function mIsForceRefCache() As Boolean
		Try
			If Math.Abs(DateDiff(DateInterval.Second, Me.mLastRefCacheTime, Now)) > Me.ForceRefCacheTime Then
				Return True
			Else
				Return False
			End If
		Catch ex As Exception
			Me.SetSubErrInf("mIsForceRefCache", ex)
			Return False
		End Try
	End Function

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
			If Me.PigKeyValues.IsItemExists(strKeyName) = True Then
				strStepName = "mRemovePigKeyValueFromList"
				strRet = Me.mRemovePigKeyValueFromList(strKeyName)
				If strRet <> "OK" Then
					Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
					If Me.PigKeyValues.IsItemExists(strKeyName) Then
						strStepName &= "(" & strKeyName & ")"
						Throw New Exception("Cannot remove exists item")
					End If
				End If
			End If
			strStepName = "PigKeyValues.Add"
			Me.PigKeyValues.Add(NewItem)
			If Me.PigKeyValues.LastErr <> "" Then
				strStepName &= "(" & strKeyName & ")"
				Throw New Exception(strKeyName)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("mAddPigKeyValueToList", ex)
		End Try
	End Function

	Public Function GetPigKeyValue(KeyName As String) As PigKeyValue
		Const SUB_NAME As String = "GetKeyValue"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			msuStatistics.GetCount += 1
			strStepName = "GetByList"
			GetPigKeyValue = Me.PigKeyValues.Item(KeyName)
			If Me.PigKeyValues.LastErr <> "" Then Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					If Not GetPigKeyValue Is Nothing Then
						msuStatistics.CacheByListCount += 1
						msuStatistics.CacheCount += 1
					End If
				Case enmCacheLevel.ToShareMem
					If GetPigKeyValue Is Nothing Or Me.mIsForceRefCache = True Then
						strStepName = "mGetPigKeyValueByShareMem"
						Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
						If Me.LastErr <> "" Then Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
						If GetPigKeyValue Is Nothing Then
							strStepName = "mRemovePigKeyValueFromList"
							strRet = Me.mRemovePigKeyValueFromList(KeyName)
							If strRet <> "OK" Then Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
						Else
							strStepName = "mAddPigKeyValueToList"
							strRet = Me.mAddPigKeyValueToList(GetPigKeyValue)
							If strRet <> "OK" Then Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
							msuStatistics.CacheByShareMemCount += 1
							msuStatistics.CacheCount += 1
						End If
					Else
						msuStatistics.CacheByListCount += 1
						msuStatistics.CacheCount += 1
					End If
				Case Else
					strStepName = KeyName
					Throw New Exception("Unsupported CacheLevel")
			End Select

			'If Not GetPigKeyValue Is Nothing Then
			'	If Me.mIsForceRefCache = True Then
			'		Select Case Me.CacheLevel
			'			Case enmCacheLevel.ToList
			'				strStepName = "mGetPigKeyValueByShareMem_2"
			'				Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
			'				If Me.LastErr <> "" Then Me.PrintDebugLog(SUB_NAME, strStepName, Me.LastErr)
			'				If GetPigKeyValue Is Nothing Then
			'					strStepName = "mRemovePigKeyValueFromList"
			'					strRet = Me.mRemovePigKeyValueFromList(KeyName)
			'					If strRet <> "OK" Then Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
			'				End If
			'		End Select
			'		Me.mLastRefCacheTime = Now
			'	End If
			'End If
			'If Not GetPigKeyValue Is Nothing Then
			'	msuStatistics.CacheCount += 1
			'	Select Case Me.CacheLevel
			'		Case enmCacheLevel.ToList
			'			msuStatistics.CacheByListCount += 1
			'		Case enmCacheLevel.ToShareMem
			'			msuStatistics.CacheByShareMemCount += 1
			'	End Select
			'End If
			'strStepName = "GetItem by list"
			'pkvNew = Me.PigKeyValues.Item(KeyName)
			'If pkvNew Is Nothing Or Me.mIsForceRefCache = True Then
			'	strStepName = "GetByCache"
			'	Dim strRet As String
			'	Select Case Me.CacheLevel
			'		Case enmCacheLevel.ToList
			'			Return Nothing
			'		Case enmCacheLevel.ToShareMem
			'			GetPigKeyValue = Me.mGetPigKeyValueByShareMem(KeyName, GetPigKeyValue)
			'		Case enmCacheLevel.ToShareMem, enmCacheLevel.ToFile, enmCacheLevel.ToDB, enmCacheLevel.ToRedis
			'			If Not GetPigKeyValue Is Nothing Then
			'				If Me.IsPigKeyValueExists(KeyName) = True Then
			'					strStepName = "RemovePigKeyValue"
			'					strRet = Me.RemovePigKeyValue(KeyName)
			'					If strRet <> "" Then
			'						strStepName &= "(" & KeyName & ")"
			'						Throw New Exception(strRet)
			'					End If
			'				End If
			'			End If
			'			Select Case Me.CacheLevel
			'				Case enmCacheLevel.ToShareMem
			'					strStepName = "New PigKeyValue(ToShareMem)"
			'					Dim pkvNew As New PigKeyValue(KeyName, Now.AddMinutes(1), "")
			'					pkvNew.Parent = Me
			'					Dim suSMHead As StruSMHead
			'					ReDim suSMHead.ValueMD5(0)
			'					strStepName = "mGetStruSMHead"
			'					strRet = Me.mGetStruSMHead(suSMHead, pkvNew.SMNameHead)
			'					If strRet <> "OK" Then
			'						strStepName &= strStepName & "(" & KeyName & "." & pkvNew.SMNameHead & ")"
			'						bolIsNotLog = True
			'						Throw New Exception(strRet)
			'					End If
			'					Dim abBody As Byte()
			'					ReDim abBody(0)
			'					strStepName = "mGetBytesSMBody"
			'					strRet = Me.mGetBytesSMBody(abBody, suSMHead, pkvNew.SMNameBody)
			'					If strRet <> "OK" Then
			'						strStepName &= strStepName & "(" & KeyName & "." & pkvNew.SMNameBody & ")"
			'						Throw New Exception(strRet)
			'					End If
			'					If Me.PigKeyValues.IsItemExists(KeyName) = False Then

			'					End If
			'					pkvNew = Nothing
			'					pkvNew = New PigKeyValue(KeyName, suSMHead.ExpTime, abBody, suSMHead.ValueType, suSMHead.ValueMD5)
			'					strStepName = "Add(pkvNew)"
			'					Me.PigKeyValues.Add(pkvNew)
			'					If Me.PigKeyValues.LastErr <> "" Then
			'						strStepName &= "(" & pkvNew.KeyName & ")"
			'						Throw New Exception(Me.PigKeyValues.LastErr)
			'					End If
			'					msuStatistics.CacheCount += 1
			'					msuStatistics.CacheByShareMemCount += 1
			'					GetPigKeyValue = pkvNew
			'				Case Else
			'					strStepName = ""
			'					Throw New Exception("Currently unsupported cachelevel")
			'			End Select
			'		Case Else
			'			strStepName = ""
			'			Throw New Exception("Currently unsupported cachelevel")
			'	End Select
			'Me.mLastRefCacheTime = Now
			'Else
			'	msuStatistics.CacheCount += 1
			'	msuStatistics.CacheByListCount += 1
			'End If
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

	Private Function mGetBytesSMBody(ByRef BodyBytes As Byte(), SuSMHead As StruSMHead, SMNameBody As String, CacheWorkDir As String) As String
		Const SUB_NAME As String = "mGetBytesSMBody"
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

	Public Sub SavePigKeyValue(NewItem As PigKeyValue, Optional IsOverwrite As Boolean = True)
		Const SUB_NAME As String = "SavePigKeyValue"
		Dim strStepName As String = ""
		Dim strRet As String
		Try
			Dim strKeyName As String = NewItem.KeyName
			msuStatistics.SaveCount += 1
			strStepName = "IsPigKeyValueExists"
			If Me.IsPigKeyValueExists(strKeyName) = True Then
				If IsOverwrite = True Then
					strStepName = "RemovePigKeyValue"
					strRet = Me.RemovePigKeyValue(strKeyName)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Throw New Exception(strRet)
					End If
				Else
					strStepName &= "(" & strKeyName & ")"
					Throw New Exception("PigKeyValue Exists")
				End If
			End If
			If NewItem.Parent Is Nothing Then NewItem.Parent = Me
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					strStepName = "mAddPigKeyValueToList"
					strRet = Me.mAddPigKeyValueToList(NewItem)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Throw New Exception(strRet)
					End If
				Case enmCacheLevel.ToShareMem
					strStepName = "mSavePigKeyValueToShareMem"
					strRet = Me.mSavePigKeyValueToShareMem(NewItem)
					If strRet <> "OK" Then
						strStepName &= "(" & strKeyName & ")"
						Throw New Exception(strRet)
					End If
					msuStatistics.SaveToShareMemCount += 1
			End Select


			'Select Case Me.CacheLevel
			'	Case enmCacheLevel.ToList
			'	Case enmCacheLevel.ToShareMem, enmCacheLevel.ToFile, enmCacheLevel.ToDB, enmCacheLevel.ToRedis
			'		Select Case Me.CacheLevel
			'			Case enmCacheLevel.ToShareMem
			'				If NewItem.Parent Is Nothing Then NewItem.Parent = Me
			'				strStepName = "mSavePigKeyValueToShareMem"
			'				Me.mSavePigKeyValueToShareMem(NewItem)
			'				If Me.LastErr <> "" Then
			'					strStepName &= "(" & NewItem.KeyName & ")"
			'					Throw New Exception(Me.LastErr)
			'				End If
			'				msuStatistics.SaveToShareMemCount += 1
			'			Case Else
			'				strStepName = ""
			'				Throw New Exception("Currently unsupported cachelevel")
			'		End Select
			'	Case Else
			'		strStepName = ""
			'		Throw New Exception("Currently unsupported cachelevel")
			'End Select
			'strStepName = "List.Add(NewItem)"
			'Me.PigKeyValues.Add(NewItem)
			'If Me.PigKeyValues.LastErr <> "" Then
			'	strStepName &= "(" & NewItem.KeyName & ")"
			'	Throw New Exception(Me.PigKeyValues.LastErr)
			'End If
			Me.ClearErr()
		Catch ex As Exception
			msuStatistics.SaveFailCount += 1
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			Me.PrintDebugLog(SUB_NAME, "Catch Exception", Me.LastErr)
		End Try
	End Sub

	Public Function RemovePigKeyValue(KeyName As String) As String
		Const SUB_NAME As String = "RemovePigKeyValue"
		Dim strStepName As String = ""
		Dim strRet As String = ""
		Try
			strStepName = "IsPigKeyValueExists"
			Dim bolIsToList As Boolean = False, bolIsToShareMem As Boolean = False
			Select Case Me.CacheLevel
				Case enmCacheLevel.ToList
					bolIsToList = True
				Case enmCacheLevel.ToShareMem
					bolIsToList = True
					bolIsToShareMem = True
			End Select
			Dim strErr As String = ""
			If bolIsToShareMem = True Then
				strStepName = "mClearShareMem"
				strRet = Me.mClearShareMem(KeyName)
				If strRet <> "OK" Then strErr &= strStepName & ":" & strRet
			End If
			If bolIsToList = True Then
				strStepName = "mRemovePigKeyValueFromList"
				strRet = Me.mRemovePigKeyValueFromList(KeyName)
				If strRet <> "OK" Then strErr &= strStepName & ":" & strRet
			End If
			If strErr <> "" Then
				strStepName &= "(" & KeyName & ")" : Throw New Exception(strErr)
			End If
			Return "OK"
		Catch ex As Exception
			strRet = Me.GetSubErrInf(SUB_NAME, strStepName, ex)
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
				Dim strRet As String
				For i = 1 To intItems
					strStepName = "RemovePigKeyValue"
					strRet = Me.RemovePigKeyValue(astrKeyName(i))
					If strRet <> "OK" Then
						strStepName &= "(" & astrKeyName(i) & ")"
						Me.PrintDebugLog(SUB_NAME, strStepName, strRet)
					End If
					'Me.PigKeyValues.Remove(astrKeyName(i))
					'If Me.PigKeyValues.LastErr <> "" Then
					'	strStepName = "Remove " & astrKeyName(i)
					'	Me.PrintDebugLog(SUB_NAME, strStepName, Me.PigKeyValues.LastErr)
					'End If
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
				oPigXml.AddEle("CacheCount", .CacheCount)
				oPigXml.AddEle("SaveCount", .SaveCount)
				oPigXml.AddEle("SaveFailCount", .SaveFailCount)
				oPigXml.AddEle("CacheByListCount", .CacheByListCount)
				Select Case Me.CacheLevel
					Case enmCacheLevel.ToShareMem
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
						oPigXml.AddEle("SaveToShareMemCount", .SaveToShareMemCount)
					Case enmCacheLevel.ToFile
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
						oPigXml.AddEle("SaveToShareMemCount", .SaveToShareMemCount)
						oPigXml.AddEle("CacheByFileCount", .CacheByFileCount)
						oPigXml.AddEle("SaveToFileCountSaveToFileCount", .SaveToFileCount)
					Case enmCacheLevel.ToDB
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
						oPigXml.AddEle("CacheByDBCount", .CacheByDBCount)
					Case enmCacheLevel.ToRedis
						oPigXml.AddEle("CacheByShareMemCount", .CacheByShareMemCount)
						oPigXml.AddEle("CacheByRedisCount", .CacheByRedisCount)
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
			Return menmCacheLevel
		End Get
		Friend Set(value As Integer)
			mintForceRefCacheTime = value
		End Set
	End Property

	Private Property mLastRefCacheTime As DateTime = DateTime.MinValue

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

End Class
