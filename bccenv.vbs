' BCC�|�[�^�u���ݒ� Ver.2
Option Explicit
Dim objWshShell,objFS,objEnv,objStream
Dim strDrv,strBccPath
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' �J�����g�h���C�u�̃h���C�u���^�[���擾����
strDrv = objFS.GetDriveName(WScript.ScriptFullName)
' borland\bcc55 �t�H���_�����݂��邩���`�F�b�N
strBccPath = strDrv & "\borland\bcc55"
If objFS.FolderExists(strBccPath)=0 Then
    WScript.Echo strBccPath & "��������܂���̂ŏI�����܂�"
    WScript.Quit 0
End If
' ���ϐ�PATH�̐擪��bin,borland\bcc55\Bin��ǉ�
Set objEnv = objWshShell.Environment("Process")
objEnv.Item("PATH") = strDrv & "\bin;" & strBccPath & "\Bin;" & objEnv.Item("PATH")
' ���ϐ�INCLUDE��LIB��ݒ�
objEnv.Item("INCLUDE") = strBccPath & "\Include"
objEnv.Item("LIB") = strBccPath & "\Lib"
' �R�}���h�v�����v�g�E�B���h�E���J��
objWshShell.Run "%COMSPEC%"