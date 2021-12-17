'**********************************
'* Name: PigKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 键值项
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 2.0
'* Create Time: 31/8/2021
'* 1.1	21/9/2021 
'* 1.2	3/12/2021 Add more new
'* 1.3	6/12/2021 Add GetSaveData,Check,InitBytesBySave
'* 2.0	15/12/2021 Modify New
'************************************
Imports PigKeyCacheLib
Imports PigToolsLiteLib
Public Class PigKeyValue
    Inherits PigKeyCacheLib.PigKeyValue

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String)
        MyBase.New(KeyName, ExpTime, KeyValue)
    End Sub

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String, TextType As PigText.enmTextType)
        MyBase.New(KeyName, ExpTime, KeyValue, TextType)
    End Sub

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue As String, TextType As PigText.enmTextType, SaveType As EnmSaveType)
        MyBase.New(KeyName, ExpTime, KeyValue, TextType, SaveType)
    End Sub

    Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte)
        MyBase.New(KeyName, ExpTime, KeyValue)
    End Sub
    Public Sub New(KeyName As String, ExpTime As Date, KeyValue() As Byte, SaveType As EnmSaveType)
        MyBase.New(KeyName, ExpTime, KeyValue, SaveType)
    End Sub


End Class
