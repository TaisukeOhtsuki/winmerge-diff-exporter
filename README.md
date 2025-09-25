# WinMerge Diff to Excel Exporter

フォルダ間の差分を**WinMerge**で比較し、結果を見やすい**Excelファイル**として出力するGUIアプリケーションです。

<img width="1002" height="522" alt="image" src="https://github.com/user-attachments/assets/396c3c7a-3868-4199-898b-61229ab489f9" />

## 機能

- **フォルダ間の差分比較**: 2つのフォルダを比較し、追加・変更・削除されたファイルを検出
- **Excelファイル出力**: 比較結果をExcel形式で保存
- **詳細な差分表示**: ファイル内容の行レベルでの差分を表示
- **ドラッグ&ドロップ対応**: フォルダをGUIに直接ドラッグして選択可能
- **プログレスバー**: 処理進行状況をリアルタイムで表示

## 必要な環境

- **Windows 10/11** (WinMergeが必要)
- **Python 3.8以上**
- **WinMerge** (以下のパスにインストールされている必要があります)
  ```
  C:\Program Files\WinMerge\WinMergeU.exe
  ```

## インストール方法

### 1. リポジトリのクローン
```bash
git clone https://github.com/your-username/winmerge-diff-exporter.git
cd winmerge-diff-exporter
```

### 2. 仮想環境の作成（推奨）
```bash
python -m venv venv
venv\Scripts\activate
```

### 3. 依存関係のインストール
```bash
pip install -r requirements.txt
```

### 4. WinMergeのインストール
WinMergeが未インストールの場合は、[公式サイト](https://winmerge.org/)からダウンロードしてインストールしてください。

## 使用方法

### 1. アプリケーションの起動
```bash
python main.py
```

### 2. フォルダの選択
- **Base Folder**: 比較元のフォルダを選択
- **Comparison Target Folder**: 比較先のフォルダを選択  
- **Output File**: 出力するExcelファイル名を指定

### 3. 実行
「Run (Compare and Export to Excel)」ボタンをクリックして比較を開始します。

### 4. 結果の確認
指定したExcelファイルが生成され、以下のシートが含まれます：
- **compare**: 差分の詳細表示
- **Summary**: ファイル一覧と変更状況
- **個別ファイルシート**: 各ファイルの詳細な差分

## プロジェクト構成

```
winmerge-diff-exporter/
├── main.py                    # メインエントリーポイント
├── gui.py                     # GUI実装
├── winmergexlsx.py           # WinMerge連携とExcel変換
├── diffdetailsheetcreater.py # 差分詳細シート作成
├── common.py                  # 共通関数とユーティリティ
├── requirements.txt           # Python依存関係
├── qt.conf                   # Qt設定ファイル
├── README.md                 # このファイル
└── LICENSE                   # MITライセンス
```

## 技術仕様

- **GUI フレームワーク**: PyQt6
- **Excel操作**: openpyxl + pywin32 (COM経由)
- **差分比較**: WinMerge (外部プロセス)
- **マルチスレッド**: QThread使用でUIブロックを回避


## ライセンス

MIT License - 詳細は[LICENSE](LICENSE)ファイルを参照してください。

## 貢献

バグ報告や機能提案は[Issues](https://github.com/your-username/winmerge-diff-exporter/issues)でお願いします。

## 謝辞

このソフトウェアは[winmerge_xlsx](https://github.com/y-tetsu/winmerge_xlsx.git)のコードを含んでいます。
