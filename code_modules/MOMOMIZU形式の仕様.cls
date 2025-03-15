VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MOMOMIZU形式の仕様"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'## *MOMOMIZU Containerフォーマット仕様（Version 1）*
'
'*このフォーマットは、バイナリデータをbase64形式でエンコードし、任意のメタデータを付与したテキスト形式のファイルとして表現するものです。完全なテキスト形式とし、人間が容易に読める必要はありませんが、構造化されていて効率的にパース可能である必要があります。*
'
'### ファイル構造
'
'MomoMizu Containerファイルは以下のテキストベースの構造で構成されます｡
'
'#### ヘッダー部（キー・バリュー形式）
'
'ヘッダー部は各行1項目のキー・バリュー形式で構成されます。各項目は改行で区切られます。
'
'- Format: 固定文字列 "MMMZ"
'- Version: フォーマットのバージョン番号（例: 1）
'- Filename: 元のファイル名（UTF-8）
'- MimeType: [任意] MIMEタイプ（省略可）
'- CreatedAt: [任意] 作成日時 ISO8601形式（省略可）
'- Author: [任意] 作成者（省略可）
'- OriginalSize: [任意] オリジナルのファイルサイズ（バイト単位、省略可）
'
'例:
'
'```
'Format: MMMZ
'Version: 1
'Filename: cat.png
'MimeType: image/png
'CreatedAt: 2025-03-05T12:00:00Z
'OriginalSize: 204800
'```
'
'ヘッダー部の終了は空行で示され､その後にデータ部が続きます｡
'
'#### データ部（Base64エンコードされたバイナリデータ）
'
'データ部はヘッダー部の後、開始・終了タグで明示的に囲まれます。
'
'```
'---BEGIN DATA---
'iVBORw0KGgoAAAANSUhEUgAA...
'---END DATA---
'```
'
'#### フッター部（任意、MD5チェックサムなど）
'
'データの整合性確認のためのMD5チェックサムを含める場合､データ部の直後にフッターとして追加します｡
'
'```
'checksum -md5: a1b2c3d4e5f67890123456789abcdef0
'```
'
'---
'
'## 各処理の仕様
'
'### エンコーダー仕様
'
'- バイナリファイルを読み込み、ヘッダー部を作成（任意のメタデータも付与可能）。
'- ヘッダー終了後に空行を挿入し、その後にbase64エンコードしたデータをデータ部として記述。
'- データのMD5チェックサムをフッター部として追加可能。
'- 保存時のファイル名は `元のファイル名.元の拡張子.txt` の形式とする。
'
'### デコーダー仕様
'
'- MOMOMIZU Container形式かどうかをヘッダー部（Formatフィールド）で判定。
'- MOMOMIZU形式の場合、ヘッダー部をパースし、メタデータ（Filename、MimeType等）を取得。
'- データ部をbase64デコードしてバイナリデータを復元。
'- フッター部にチェックサムが存在すればMD5で整合性確認を行う。
'- ファイル保存時には取得したFilenameをデフォルト名として提示する。
'- MOMOMIZU形式でない一般的なbase64データの場合は、デコード後にマジックナンバーを元に自動で適切な拡張子を判定し、ファイル名を生成。
'
'### パーサー仕様
'
'- MOMOMIZU Containerファイルを読み込み、ヘッダー部を解析して構造化データとして提供。
'- データ部のbase64文字列を抽出。
'- フッター部（チェックサム等）を解析し提供。
'
'---
'
'## 特殊ケース処理
'
'- MIMEタイプが存在しない場合でも処理可能とする。
'- ファイル名命名規則に沿った一般的なbase64テキストファイルを、命名規則から自動判定してデコード可能とする。
'- Version番号を用いて将来的なフォーマットの拡張に柔軟に対応可能とする。
'
'---
'
'## エラー処理・仕様違反時の挙動
'
'- MOMOMIZU Containerとしてパースできない場合、一般的なbase64データとして扱い、デコード処理を試みる。
'- チェックサム不一致の場合はユーザーに警告を表示し、処理を継続するか選択可能とする。
'
'---
'



' --- sample ---
' =================================
'Format: MMMZ
'Version: 1
'Filename: example.png
'CreatedAt: 2025-03-05T12:00:00Z
'
'---BEGIN DATA---
'/9j/4AAQSkZJRgABAgAAZABkAAD/7AARRHVja3kAAQAEAAAAKAAA/+.....
'---END DATA---
'
'checksum -md5: a1b2c3d4e5f67890123456789abcdef0
' =================================
