
## 1. Overview
This tool is an Excel VBA-based utility designed to enable the transfer of files required for legitimate business purposes even in environments with strict file transfer restrictions. It works by converting the information and contents of an original file (e.g., an Excel file with macros) into a predefined text format, which can then be decoded by the recipient to recreate the original file.

## 2. Background and Purpose
In many organizations, the connection of USB drives and direct file transfers are tightly controlled, with detailed logs maintained. Even downloads via web browsers may undergo sanitization (filtering).  
To address these challenges, this tool was developed with the following ideas in mind:
- **Converting Files to Text Format**  
  The tool converts the original file’s name, file size, version information, and content (encoded in Base64) into a unique text format.
- **Text-Based Transfer Method**  
  If filtering occurs during browser downloads, it is also possible to share the text file via services like Google Drive by simply passing the URL to the recipient. The recipient can then open the text file in a browser, copy its entire content to the clipboard, paste it into a plain text editor (e.g., Notepad), and finally load it into the tool to restore the original file.

## 3. Features
- **Custom Text Format**  
  The tool consolidates the basic metadata of the original file (including file name, size, version information—see Developer Information for details) and the encoded content data into a single text file.
- **Flexible Transfer Options**  
  If downloads via a browser are filtered, you can share the text file via Google Drive (or a similar service), enabling transfer via browser-based copy and paste.
- **Excel VBA Implementation**  
  The tool is implemented using Excel VBA and is composed of multiple `.bas` (standard modules) and `.cls` (class modules) files that run within an Excel VBA environment.

## 4. Usage Guidelines and Disclaimer
- **For Legitimate Business Use Only**  
  This tool is intended solely for legitimate business purposes. Avoid any improper usage or applications that violate your organization's security policies. It is recommended that you obtain the appropriate approvals within your organization’s security policy and internal control framework.
- **Regarding Safety**  
  The tool only converts and restores the file content into text format; it does not guarantee the security of the original file. Please consult with your organization’s security personnel before use.
- **Disclaimer**  
  The developer assumes no responsibility for any issues arising from the use of this tool. Users are responsible for ensuring that they comply with organizational regulations and have implemented adequate risk management measures.

## 5. Installation Instructions
1. Launch Excel and open the VBA editor.
2. Import the provided `.bas` and `.cls` files into your VBA project.
3. Adjust the project settings and Excel’s security levels as needed.

## 6. Usage Instructions
### 6.1. Converting a File
1. Select the file you wish to transfer and run the conversion process.
2. The tool will generate text data that includes the file name, size, version, and file content (encoded in Base64).

For ease of use, assign the `DecodeMomomizuContainerFile` and `EncodeBinaryFileToMomomizuContainer` procedures to buttons or similar controls.

---

## 1. 概要
本ツールは、厳しいファイル転送制限が敷かれている環境でも、正当な業務目的で必要なファイルのやりとりを可能にするために作成されたExcel VBAベースのユーティリティです。元のファイル（例：マクロ付きExcelファイル）の情報と内容を、あらかじめ定めたテキスト形式に変換し、相手側でそのテキストを復号することで、元のファイルを再現できる仕組みです。

## 2. 背景・目的
多くの組織では、USBメモリの接続や直接のファイル転送が厳重に管理され、ログが詳細に記録されています。また、ブラウザ経由のダウンロードであっても、無害化処理（フィルタリング）がかかる場合があります。  
そこで、以下の発想でツールを開発しました：
- **ファイルをテキストファイルに変換する**  
  ファイル名、ファイルサイズ、バージョン情報に加え、元のファイル内容をエンコード（Base64形式）した独自フォーマットでテキスト化します。
- **テキスト形式ならではの転送手段**  
  もしブラウザ経由でダウンロードする際にフィルタリングがかかってしまう場合、Google Driveなどでテキストファイルを共有し、相手にはそのURLを渡すことも可能です。受け取った側は、ブラウザでそのテキストファイルを開き、全内容をコピーしてクリップボードへ取り込み、メモ帳等にペースト後、ツールで読み込むことでファイルを復元できます。

## 3. ツールの特徴
- **カスタムテキストフォーマット**  
  元のファイルの基本情報（ファイル名、サイズ、バージョンなどを含む。詳細はdevelopers info参照）と、内容をエンコードしたデータを一つのテキストファイルにまとめます。
- **柔軟な転送手段**  
  ブラウザでのダウンロードがフィルタリングされる場合は、Google Driveなどで共有し、ブラウザ経由のコピー＆ペーストによる転送も想定しています。
- **Excel VBA実装**  
  ツールは、複数の`.bas`（標準モジュール）と`.cls`（クラスモジュール）により構成されており、Excel VBA環境で動作します。

## 4. 利用上の注意と免責事項
- **正当な業務利用に限定**  
  本ツールは、業務上の正当な目的での利用を前提としています。不正な使用や、組織のセキュリティポリシーに反する利用は避けてください。組織のセキュリティポリシーおよび内部統制の枠組み内で、関係部門の承認を得た上で行うことが望まれます。
- **安全性について**  
  ツールはファイル内容をテキストに変換・復元するのみで、元のファイルの安全性自体を保証するものではありません。利用前に、所属組織のセキュリティ担当者と十分に確認してください。
- **免責事項**  
  本ツールの利用により発生したいかなる問題についても、開発者は一切の責任を負いません。利用者自身が、組織の規定や管理体制に則って十分なリスク管理を行うことが前提です。

## 5. インストール方法
1. Excelを起動し、VBAエディタを開きます。
2. 本ツールに含まれる .bas および .cls ファイルをプロジェクトにインポートしてください。
3. 必要に応じて、プロジェクトの設定やExcelのセキュリティレベルを調整してください。

## 6. 使用方法
### 6.1. ファイルの変換
1. 転送したいファイルを選択して、変換処理を実行します。  
2. ツールは、ファイル名、サイズ、バージョン、及びファイル内容（Base64エンコード済み）などの情報を含むテキストデータを生成します。

`DecodeMomomizuContainerFile`と`EncodeBinaryFileToMomomizuContainer`をボタンなどに割り当てて呼ぶと楽です


### 6.2. 転送方法について
- **直接ダウンロードの場合**  
  生成されたテキストデータをファイルとして保存し、通常のダウンロード手段で相手側へ渡します。
- **Google Drive経由の場合**  
  テキストファイルがダウンロード時にフィルタリングされる場合は、Google Drive等でファイルを共有し、URLのみを相手に伝えます。  
  受け取った側は、ブラウザでそのテキストファイルを開き、全内容をコピーしてクリップボードへ取り込み、メモ帳等にペースト&保存後、ツールのデコード機能で読み込みます。

### 6.3. ファイルの復元
1. 転送されたテキストデータをツールの復元機能で読み込みます。  
2. ツールがテキストデータから各情報を解析し、元のファイルを再現します。

## 7. 注意事項・既知の問題
- テキストファイルの作成やコピー時に、不要な改行や空白が混入する可能性があります。必要に応じて、テキストの整形や検証を行ってください。
- 組織のセキュリティ設定により、テキストファイル自体がフィルタリング対象となったり、一部の動作が制限される場合があります。利用前に、実際の環境で十分なテストを実施してください。


## 8. ライセンス
MIT

---

## Developer Information

