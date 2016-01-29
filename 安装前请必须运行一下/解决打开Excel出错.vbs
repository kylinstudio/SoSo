
On Error Resume Next 
Dim myFile,myDllname,SySPath
myDllname="MSCOMCTL.OCX"
'--------------------------------------------------'系统路径
Set MyFile = CreateObject("Scripting.FileSystemObject")
SySPath = MyFile.GetSpecialFolder(1)&"\"
'--------------------------------------------------’当前文件路径
sPath=MyFile.GetFile(WScript.ScriptFullName).ParentFolder.Path&"\"
If MsgBox("安装前需要关闭【Excel】【Word】和【月光迷你钟】,是否现在关闭",vbYesNo,"【SoSo提示您】")=vbyes Then
	Call Checkit
	'--------------------------------------------------'注册dll
	Set objfile = MyFile.GetFile(SySPath & myDllname)
	objfile.Name = "MSCOMCTLs.OCX"'重命名
	'--------------------------------------------------
	MyFile.CopyFile sPath & myDllname,SySPath
	Set objShell=CreateObject("WScript.Shell")
	sFile=SySPath & myDllname
	objShell.Run "regsvr32 " & Chr(34) & sFile & Chr(34)
Else
	Call WsClose
End If

Sub Checkit()
	Dim arr,i,ir
	arr=Array("Excel.exe","Word.exe","月光迷你钟.exe")
	strComputer="." 
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
	For i= 0 To UBound(arr)
		Set colProcessList=objWMIService.ExecQuery ("select * from Win32_Process where Name='"& arr(i)&  "'") 
		For Each objProcess In colProcessList 
			objProcess.Terminate()
		Next 
	Next
End Sub
Sub WsClose()
	WScript.Quit
End Sub