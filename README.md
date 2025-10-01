# WinMerge Diff to Excel Exporter

フォルダ間の差分を**WinMerge**で比較し、結果を見やすい**Excelファイル**として出力するGUIアプリケーションです。

<img width="1002" height="522" alt="image" src="https:/## 技術仕様

### アーキテクチャ
- **GUIフレームワーク**: PyQt6 6.9.1
- **Excel操作**: openpyxl 3.1.5 (純粋Python、COM不要)
- **HTML解析**: BeautifulSoup4 4.12.3 + lxml 5.3.0
- **差分比較**: WinMerge (外部プロセス、HTML出力形式)
- **マルチスレッド**: QThread使用でUIブロック回避
- **エラーハンドリング**: 3段階のファイル保存戦略

### 主な設計パターン
- **MVC分離**: UI、ビジネスロジック、データ処理を分離
- **シグナル/スロット**: Qt非同期通信パターン
- **ワーカースレッド**: 長時間処理をバックグラウンドで実行
- **リトライメカニズム**: ファイルロック時の自動再試行

### ファイル名抽出アルゴリズム
WinMergeのHTML出力ファイル名（`folder1_folder2_..._filename`形式）から、確実にファイル名を抽出：
- フォルダキーワード検出（`modules`, `ctrl`, `tool`, `gui`, `etc` など）
- 拡張子から逆順スキャン
- フォルダキーワードの次からファイル名として認識
- 例: `sc2stb_ctrl_v4_LINK_CTRLA_SP3_2.l` → `CTRLA_SP3_2.l`

### 差分検出ロジック
- WinMergeの差分色 `FFEFCB05` (黄色) を検出
- コンテキスト行数: 前後4行
- 隣接する差分ブロックを自動マージ
- フォルダ行はリンク有無で識別して除外r-attachments/assets/396c3c7a-3868-4199-898b-61229ab489f9" />

## 目次

- [主な機能](#主な機能)
- [システム要件](#システム要件)
- [インストール方法](#インストール方法)
- [使用方法](#使用方法)
- [プロジェクト構成](#プロジェクト構成)
- [技術仕様](#技術仕様)
- [設定のカスタマイズ](#設定のカスタマイズ)
- [トラブルシューティング](#トラブルシューティング)
- [開発](#開発)
- [変更履歴](#変更履歴)
- [ライセンス](#ライセンス)

## 主な機能

- **フォルダ間の差分比較**: 2つのフォルダを比較し、追加・変更・削除されたファイルを自動検出
- **Excel形式で出力**: 比較結果を見やすいExcelファイルとして保存
- **詳細な差分表示**: ファイル内容の行レベルでの差分を色分けして表示
- **ドラッグ&ドロップ対応**: フォルダをGUIに直接ドラッグして簡単選択
- **プログレスバー**: 処理進行状況をリアルタイムで表示
- **複数シート生成**: 
  - **compareシート**: 差分ブロックのみを抽出した詳細表示（ファイル名ラベル付き）
  - **個別ファイルシート**: 各ファイルの完全な差分（ファイル名がシート名）
  - **Summaryシート**: ファイル一覧と変更状況の概要（フォルダ行を自動除外）

## システム要件

- **Python 3.8以上** (Python 3.13.3で動作確認済み)
- **WinMerge** (デフォルトパス: `C:\Program Files\WinMerge\WinMergeU.exe`)

### ? バージョン 2.0 の主な改善点

**Microsoft Excel のインストールは不要になりました！**

- ? **Excelなしで動作**: Excelがインストールされていなくても完全に動作
- ? **ファイルロック対策**: Excelでファイルを開いていても実行可能
  - 3段階の保存戦略（直接保存→一時ファイル経由→タイムスタンプ付き）
- ? **高速で安定**: COM依存を排除し、純粋なPythonライブラリで処理
- ? **改善された体裁**: 
  - 縦罫線のみでスッキリした見た目（横罫線なし）
  - 行番号列の「.」を自動削除
  - 差分がある行のみ背景色を表示
- ? **正確なファイル名表示**:
  - シート名がファイル名に（パス形式ではなく）
  - compareシートでも正しいファイル名を表示
  - 長いパス名でも確実にファイル名を抽出
- ? **適切な差分検出**:
  - WinMergeの黄色差分行を正確に検出
  - Summaryシートからフォルダ行を自動除外

## インストール方法

### 1. リポジトリのクローン
```bash
git clone https://github.com/TaisukeOhtsuki/winmerge-diff-exporter.git
cd winmerge-diff-exporter
```

### 2. 仮想環境の作成（推奨）
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/macOS
source venv/bin/activate
```

### 3. 依存関係のインストール
```bash
pip install -r requirements.txt
```

**必要なライブラリ:**
- PyQt6 6.9.1 - GUI フレームワーク
- openpyxl 3.1.5 - Excel ファイル操作
- beautifulsoup4 4.12.3 - HTML パース
- lxml 5.3.0 - XML/HTML パーサー

### 4. WinMergeのインストール
WinMergeが未インストールの場合は、[公式サイト](https://winmerge.org/)からダウンロードしてインストールしてください。

**デフォルトインストールパス:**
```
C:\Program Files\WinMerge\WinMergeU.exe
```

カスタムパスの場合は `src/core/config.py` で設定を変更できます。

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
- **compareシート**: 差分ブロックのみを抽出（コンテキスト4行付き、ファイル名ラベル表示）
- **Summaryシート**: ファイル一覧と変更状況の概要（フォルダ行は除外）
- **個別ファイルシート**: 各ファイルの完全な差分（シート名はファイル名、行番号・色付き）

### 5. ファイルロック時の動作
出力ファイルが開かれている場合、以下の戦略で保存を試みます：
1. **直接保存** (5回リトライ、段階的待機)
2. **一時ファイル経由** (アトミックな名前変更)
3. **タイムスタンプ付き** (例: `output_20250930_143052.xlsx`)

## プロジェクト構成

```
winmerge-diff-exporter/
├── main.py                 # アプリケーションエントリーポイント
├── requirements.txt        # Python依存パッケージ
├── qt.conf                # Qt設定
├── LICENSE                # MITライセンス
├── README.md              # このファイル
│
├── src/                   # ソースコード
│   ├── core/              # コアビジネスロジック
│   │   ├── common.py      # ロガーとタイマー
│   │   ├── config.py      # 設定管理
│   │   ├── exceptions.py  # カスタム例外
│   │   ├── utils.py       # ファイル操作とExcel整形
│   │   ├── winmergexlsx.py           # WinMerge統合
│   │   └── diffdetailsheetcreater.py # 差分詳細シート作成
│   │
│   ├── converters/        # ファイル変換
│   │   └── html_to_excel.py  # HTML→Excel変換（COM不要）
│   │
│   └── ui/                # ユーザーインターフェース
│       └── gui.py         # PyQt6 GUI
│
├── docs/                  # ドキュメント
│   ├── PROJECT_STRUCTURE.md       # プロジェクト構造詳細
│   ├── EXCEL_COM_REMOVAL.md       # Excel COM削除の技術詳細
│   ├── FILE_LOCK_COMPLETE_FIX.md  # ファイルロック対策
│   ├── REFACTORING_SUMMARY.md     # リファクタリング履歴
│   └── RELEASE_NOTES_v2.0.md      # リリースノート
│
├── tests/                 # ユニットテスト
├── output/                # 出力ファイル
└── venv/                  # Python仮想環境
```

詳細は [`docs/PROJECT_STRUCTURE.md`](docs/PROJECT_STRUCTURE.md) を参照してください。

## 技術仕様

### アーキテクチャ
- **GUIフレームワーク**: PyQt6 6.9.1
- **Excel操作**: openpyxl 3.1.5 (純粋Python、COM不要)
- **HTML解析**: BeautifulSoup4 4.12.3 + lxml 5.3.0
- **差分比較**: WinMerge (外部プロセス、HTML出力形式)
- **マルチスレッド**: QThread使用でUIブロック回避
- **エラーハンドリング**: 3段階のファイル保存戦略

### 主な設計パターン
- **MVC分離**: UI、ビジネスロジック、データ処理を分離
- **シグナル/スロット**: Qt非同期通信パターン
- **ワーカースレッド**: 長時間処理をバックグラウンドで実行
- **リトライメカニズム**: ファイルロック時の自動再試行


## 設定のカスタマイズ

`src/core/config.py` で以下の設定を変更できます：

- **WinMergeパス**: カスタムインストール先を指定
  ```python
  executable_path: str = r'C:\Program Files\WinMerge\WinMergeU.exe'
  ```
- **差分色**: WinMerge差分行の色コード
  ```python
  yellow_color: str = 'FFEFCB05'  # WinMerge diff color (yellow)
  ```
- **コンテキスト行数**: 差分前後の表示行数
  ```python
  context_lines: int = 4
  ```
- **列幅**: Excel列の幅設定（`diff_formats` 辞書）
- **フォルダキーワード**: ファイル名抽出時に認識するフォルダ名（`_extract_filename_from_stem` メソッド）

## トラブルシューティング

### WinMergeが見つからない
```
WinMergeNotFoundError: WinMerge not found at ...
```
→ `src/core/config.py` の `winmerge_path` を正しいパスに設定してください。

### ファイルが保存できない
出力ファイルが開かれている場合、タイムスタンプ付きファイルが作成されます。
例: `output_20250930_143052.xlsx`

### DPIスケーリング問題
高DPI環境でGUIが正しく表示されない場合、`main.py` の DPI設定が自動調整します。

## 開発

### 新機能の追加
1. コアロジック → `src/core/`
2. UI コンポーネント → `src/ui/`
3. ファイル変換 → `src/converters/`
4. テスト → `tests/`
5. ドキュメント → `docs/`

### コードスタイル
- 相対インポート: パッケージ内 (`.module`)
- 絶対インポート: パッケージ外 (`src.package.module`)
- UTF-8エンコーディング
- 英語コメント推奨

### テスト実行
```bash
# 構文チェック
python -m py_compile main.py

# アプリケーション実行
python main.py
```

## 変更履歴

### Version 2.1 (2025-10-01)
- ? シート名をファイル名に変更（パス形式から改善）
- ? compareシートのファイル名表示を修正
- ? 正確なファイル名抽出アルゴリズム実装
- ? 差分色の修正（灰色→黄色 `FFEFCB05`）
- ? Summaryシートからフォルダ行を除外
- ? 実際の保存ファイル名をGUIに表示
- ? デバッグログの追加

### Version 2.0 (2025-09)
- ? Excel COM依存を完全削除
- ? ファイルロック対策の実装
- ? プロジェクト構造の再編成
- ? 体裁の改善（縦罫線のみ、行番号の「.」削除）
- ? DPIスケーリング対応
- ? アプリケーションフリーズ問題の修正

詳細は [`docs/RELEASE_NOTES_v2.0.md`](docs/RELEASE_NOTES_v2.0.md) を参照してください。

## ライセンス

MIT License - 詳細は[LICENSE](LICENSE)ファイルを参照してください。

## 貢献

バグ報告や機能提案は[Issues](https://github.com/TaisukeOhtsuki/winmerge-diff-exporter/issues)でお願いします。

プルリクエストも歓迎します！

## 謝辞

このソフトウェアは[winmerge_xlsx](https://github.com/y-tetsu/winmerge_xlsx.git)のコードをベースに開発されました。

## 作者

**TaisukeOhtsuki** - [GitHub](https://github.com/TaisukeOhtsuki)

---

? このプロジェクトが役に立ったら、スターをお願いします！
