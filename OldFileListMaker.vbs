' OldFileListMaker.vbs

'初期宣言
On Error Resume Next
Dim Checkdate
Checkdate = DateSerial(2010, 1, 1)
'フォルダー、ファイル、そのステータス
Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")
Dim Folder
Set Folder = Fso.GetFolder(".")
Dim Subfolder
Dim File
Dim Fname
Dim Cdate
Dim Mdate
Dim Adate
'CSV
Dim Csv
Set Csv = Fso.CreateTextFile(".\OldFileList.csv",True)
Csv.WriteLine("File Name, Created Date, Modified Date, Accessed Date")

'処理開始
Call CheckFolder(Folder)

Sub CheckFolder (Folder)
    'ファイル探索
    For Each File In Folder.Files
       Fname = File.Path
       Cdate = File.DateCreated
       Mdate = File.DateLastModified
       Adate = File.DateLastAccessed
       if Checkdate > Mdate Then
           Csv.WriteLine(Fname & "," & Cdate & "," & Mdate & "," & Adate)
       End if
    Next

    'サブフォルダー探索
    For Each Subfolder In Folder.SubFolders
        Fname = "<DIR> " & SubFolder.Name
        Cdate = Subfolder.DateCreated
        Mdate = Subfolder.DateLastModified
        Adate = Subfolder.DateLastAccessed
        'Csv.WriteLine(Fname & "," & Cdate & "," & Mdate & "," & Adate)

        '再帰的処理
        Call CheckFolder (Subfolder)
    Next

End Sub

WScript.Echo "Check End"

'開放宣言
Set Fso = Nothing
Set Folder = Nothing
Set Subfolder = Nothing
Set File = Nothing
Set Csv = Nothing
Set Cdate = Nothing
Set Mdate = Nothing
Set Adate = Nothing
Set Fname = Nothing
