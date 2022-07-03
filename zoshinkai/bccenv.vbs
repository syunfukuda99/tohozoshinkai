' BCCポータブル設定 Ver.2
Option Explicit
Dim objWshShell,objFS,objEnv,objStream
Dim strDrv,strBccPath
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' カレントドライブのドライブレターを取得する
strDrv = objFS.GetDriveName(WScript.ScriptFullName)
' borland\bcc55 フォルダが存在するかをチェック
strBccPath = strDrv & "\borland\bcc55"
If objFS.FolderExists(strBccPath)=0 Then
    WScript.Echo strBccPath & "が見つかりませんので終了します"
    WScript.Quit 0
End If
' 環境変数PATHの先頭にbin,borland\bcc55\Binを追加
Set objEnv = objWshShell.Environment("Process")
objEnv.Item("PATH") = strDrv & "\bin;" & strBccPath & "\Bin;" & objEnv.Item("PATH")
' 環境変数INCLUDEとLIBを設定
objEnv.Item("INCLUDE") = strBccPath & "\Include"
objEnv.Item("LIB") = strBccPath & "\Lib"
' コマンドプロンプトウィンドウを開く
objWshShell.Run "%COMSPEC%"