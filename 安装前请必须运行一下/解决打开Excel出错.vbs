
On Error Resume Next 
Dim myFile,myDllname,SySPath
myDllname="MSCOMCTL.OCX"
'--------------------------------------------------'ϵͳ·��
Set MyFile = CreateObject("Scripting.FileSystemObject")
SySPath = MyFile.GetSpecialFolder(1)&"\"
'--------------------------------------------------����ǰ�ļ�·��
sPath=MyFile.GetFile(WScript.ScriptFullName).ParentFolder.Path&"\"
If MsgBox("��װǰ��Ҫ�رա�Excel����Word���͡��¹������ӡ�,�Ƿ����ڹر�",vbYesNo,"��SoSo��ʾ����")=vbyes Then
	Call Checkit
	'--------------------------------------------------'ע��dll
	Set objfile = MyFile.GetFile(SySPath & myDllname)
	objfile.Name = "MSCOMCTLs.OCX"'������
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
	arr=Array("Excel.exe","Word.exe","�¹�������.exe")
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