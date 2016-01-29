# OldFileDeleteVBS

## ファイル

- OldFileListMaker.vbs
- DeleteFileFromCsv.vbs
- README.md

以上3つ。

- OldFileListMaker.vbs
  `OldFileListMaker.vbs`の置いてあるディレクトリ及び、そのサブディレクトリを再帰的に探索し、  
  古いファイル(デフォルトの設定では2010/01/01よりも古いファイル)をCSVファイルに書き出す。  
  実行が終わると「Check End」とメッセージウィンドウが表示される。

  これで作成されるCSVファイルは`OldFileListMaker.vbs`と同じディレクトリに OldFileList.csv という名前で置かれる。  
  ここに記述してあるファイルは`DeleteFileFromCsv.vbs`を起動すると消えるので、絶対にこのファイルに対して書き込みは行わないこと。  
  削除するときは行ごと削除すること。
　
- DeleteFileFromCsv.vbs
  上記OldFileListMaker.vbsから作成されたOldFileList.csvを読み取り、そこに記述してあるファイルを削除する。  
  消してほしくないファイルはCSVから行ごと削除しておくこと。  
  また、Delete Endと表示されるまでOldFileList.csvには触らないこと。

- README.md
  今読んでいるこれのことです。

## 手順

1. `OldFileListMaker.vbs`と`DeleteFileFromCsv.vbs`の設置  
2. `OldFileListMaker.vbs`の実行  
3. 出来上がった`OldFileList.csv`のチェック  
4. `DeleteFileFromCsv.vbs`の実行  
5. `OldFileListMaker.vbs`の実行  
6. 出来上がった`OldFileList.csv`のチェック。何もなければおっけー  
