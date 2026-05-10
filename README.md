# DExcelAssist v119 Installer

DExcelAssistのインストーラ版です。

## 主な変更

- アップデート確認は、GitHub main上の `DExcelAssistInstaller.zip` をダウンロードしてインストーラを起動する方式に変更しました。
- `DExcelAssistAppEvents` はインストール時にクラスモジュールとして生成する方式に変更し、VBEでヘッダーが通常モジュールに混入する問題を回避します。
- 他のExcelアドインは削除しません。

## インストール

1. Excelをすべて閉じる
2. `DExcelAssist.bat` を実行
3. `1: インストール/修復` を実行

## GitHub mainブランチへ置くもの

`5: リリース用ファイル作成` で作成される `_release/main_branch_upload` の中身だけを main ブランチ直下に配置してください。

- VERSION.txt
- version.json
- DExcelAssistInstaller.zip
- README.md


## リリース用ファイル作成

DExcelAssist.bat を実行し、`5: リリース用ファイル作成` を選択すると、GitHub main ブランチ直下に配置するファイル一式が `_release/main_branch_upload` に作成されます。

コマンドで作成する場合は以下も利用できます。

```bat
DExcelAssist.bat /release
```

main ブランチには `_release/main_branch_upload` の中身だけを配置してください。
