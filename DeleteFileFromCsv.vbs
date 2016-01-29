' DeleteFileFromCsv.vbs

'初期設定
On Error Resume Next

'ファイル、その他
Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim Target
Dim Filepath

'Csv設定
Dim Csv
Set Csv = Fso.OpenTextFile(".\OldFileList.csv",1)

Csv.ReadLine

Do Until Csv.AtEndOfStream
    Target=Csv.ReadLine
    Filepath=Split(Target,",")
    Fso.DeleteFile Filepath(0),True
    'WScript.Echo "Delete " & Filepath(0)
Loop

Csv.Close

WScript.Echo "Delete End"

Set Fso = Nothing
Set Target = Nothing
Set Filepath = Nothing
Set Csv = Nothing
