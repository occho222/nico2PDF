# Nico2PDF - リリースノート

## バージョン 1.4.0 (2024年12月17日)

### ? 新機能
- **クリップボード機能の改善**
  - ファイル名のみをシンプルな改行区切り形式でクリップボードにコピーする機能を追加
  - 従来の詳細情報付きコピー機能に加えて、シンプルなファイル名リストの取得が可能
  - 外部アプリケーションとの連携が容易になりました

### ?? 改善点
- **クリップボード処理の最適化**
  - ユーザーの使用頻度が高いシンプルなファイル名コピー機能を優先実装
  - 選択されたファイルのみ、または全ファイルの選択オプションを維持
  - ユーザーエクスペリエンスの向上

### ?? バグ修正
- クリップボード機能のデバッグ完了
- 各種UI表示の安定性向上

### ?? 技術的な変更
- `FileManagementService.CopyFileNamesToClipboard` メソッドを新規追加
- `MainWindow.BtnCopyToClipboard_Click` メソッドをシンプルな形式に変更
- バージョン情報を1.4.0に更新

---

## バージョン 1.3.0 (2024年12月14日)

### ? 新機能
- **ファイル更新時の自動選択機能**: ファイル更新ボタンを押した際に、PDF化されていないファイルに自動的にチェックが入るようになりました
  - 新規追加されたファイルで未変換のものは自動選択
  - 既存ファイルでPDFが存在しなくなった場合も自動選択
  - PDFファイル自体は対象外（変更なし）

### ?? 改善点
- ファイル管理効率の向上：PDF変換が必要なファイルを見つけやすくなりました
- ユーザビリティの向上：手動でチェックを入れる手間を軽減

### ?? バグ修正
- なし

### ?? 技術的詳細
- `FileManagementService.UpdateFiles`メソッドにPDF未変換ファイルの自動選択ロジックを追加
- PDFファイルの存在確認を強化し、適切な選択状態を維持

---

## バージョン 1.2.0 (以前のリリース)

### 主要機能
- Office文書（Word、Excel、PowerPoint）のPDF変換
- 複数PDFファイルの結合機能
- プロジェクト管理機能
- ファイル名変更機能（単体・一括）
- サブフォルダ対応
- ドラッグ&ドロップ対応

### サポートファイル形式
- Microsoft Word: .doc, .docx
- Microsoft Excel: .xls, .xlsx, .xlsm
- Microsoft PowerPoint: .ppt, .pptx
- PDF: .pdf

### システム要件
- .NET 6.0 以上
- Windows OS
- Microsoft Office（Word、Excel、PowerPoint）

### 依存関係
- Microsoft.Office.Interop.Word
- Microsoft.Office.Interop.Excel
- Microsoft.Office.Interop.PowerPoint
- iTextSharp
- BouncyCastle.Cryptography

---

## インストール・使用方法

1. Microsoft Office がインストールされていることを確認
2. .NET 6.0 ランタイムがインストールされていることを確認
3. Nico2PDF.exe を実行
4. プロジェクトを作成または選択
5. 対象フォルダを指定
6. ファイルを読み込み
7. PDF変換・結合を実行
8. ファイル一覧をクリップボードにコピー（新機能！）

詳細な使用方法については README.md をご参照ください。

---

## 動作環境
- Windows 10/11
- .NET 6.0
- Microsoft Office（Word、Excel、PowerPoint）

## 今後の予定
- UI/UXのさらなる改善
- 変換オプションの追加
- パフォーマンスの最適化