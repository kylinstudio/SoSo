Dim myV
Const HKEY_CURRENT_USER = &H80000001
myV=Array("11.0","12.0","13.0","14.0")
If MsgBox ("卸载SoSo会自动关闭Excel，是否进行卸载",vbYesNo,"SoSo提示您")= vbyes Then
	Call Checkit
	Set objWMI = GetObject("winmgmts:\\.\root\default:StdRegProv")
	'--------------------------------------------------
	Set regex1 = CreateObject("VBSCRIPT.REGEXP")    'RegEx为建立正则表达式
	With regEx1
		.Global = True    '设置全局可用
		.Pattern = "OPEN\d?|SoSo"
		.IgnoreCase=False
	End With
	'--------------------------------------------------
	For j=0 To UBound(myV)
		myV1 = "Software\Microsoft\Office\"& myv(j) & "\Excel\Options"
		objWMI.EnumValues HKEY_CURRENT_USER, myV1, vValue, vType
		If IsArray(vValue) = True Then
			ir = UBound(vValue)
			For i = 0 To ir
				If regEx1.Test(vValue(i))=True Then
					objWMI.GetStringValue HKEY_CURRENT_USER, myV1, vValue(i), Vname
					If regEx1.Test(Vname)=True Then objWMI.DeleteValue HKEY_CURRENT_USER, myV1, vValue(i)
				End If
			Next 
		End If
	Next 
	MsgBox "SoSo卸载完成,重新打开Excel即可",48,"SoSo提示您"
End If

'--------------------------------------------------
Sub Checkit()
	Dim arr,i,ir
	arr=Array("Excel.exe")
	strComputer="." 
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
	For i= 0 To UBound(arr)
		Set colProcessList=objWMIService.ExecQuery ("select * from Win32_Process where Name='"& arr(i)&  "'") 
		For Each objProcess In colProcessList 
			objProcess.Terminate()
		Next 
	Next
End Sub