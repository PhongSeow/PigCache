'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 31/8/2021
'* 1.1	21/9/2021 
'************************************
Imports PigKeyCacheLib
Public Class PigKeyValue
    Inherits PigKeyCacheLib.PigKeyValue
    'Private Const CLS_VERSION As String = "1.1.1"
    'Private moPigBaseMini As New PigBaseMini(CLS_VERSION)

    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As String, Optional MatchValueMD5 As String = "")
        MyBase.New(KeyName, ExpTime, KeyValue, MatchValueMD5)
    End Sub

    Public Sub New(KeyName As String, ExpTime As DateTime, KeyValue As Byte(), ValueType As enmValueType, Optional MatchValueMD5Bytes As Byte() = Nothing)
        MyBase.New(KeyName, ExpTime, KeyValue, ValueType, MatchValueMD5Bytes)
    End Sub

End Class
