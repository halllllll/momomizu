Attribute VB_Name = "HandleBase64"
Option Explicit


Function Base64Decode(ByVal base64String As String) As Byte()
    Dim xmlDoc As Object, xmlNode As Object
    On Error GoTo ErrHandler
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = base64String
    Base64Decode = xmlNode.nodeTypedValue
    Exit Function
ErrHandler:
    MsgBox "デコード中にエラーが発生しました。文字列の形式をご確認ください。", vbCritical
    Base64Decode = VBA.Array() ' 空配列
End Function

Function Base64Encode(ByRef byteData() As Byte) As String
    Dim xmlDoc As Object, xmlNode As Object
    On Error GoTo ErrHandler
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = byteData
    Base64Encode = xmlNode.Text
    Exit Function
ErrHandler:
    MsgBox "エンコード中にエラーが発生しました。", vbCritical
    Base64Encode = ""
End Function


' --------- デバッグ用 -----------

' Base64が書かれたテキストファイルを読み込み、デコード後にバイナリファイルとして保存する処理
Sub DecodeBase64FromFile()
    Dim base64String As String
    Dim decodedBytes() As Byte
    Dim filePath As Variant
    Dim fileNum As Integer
    Dim fileContent As String
    Dim savePath As Variant

    ' ユーザーへの案内
    MsgBox "【デコード処理】" & vbCrLf & _
           "Base64が書かれたテキストファイルを選択します。", vbInformation

    ' テキストファイル選択
    filePath = Application.GetOpenFilename(FileFilter:="Text Files (*.txt), *.txt", _
                                             Title:="Base64が書かれたテキストファイルを選択してください")
    If filePath = False Then
        MsgBox "ファイルの選択がキャンセルされました。", vbInformation
        Exit Sub
    End If

    ' テキストファイルの内容を読み込む
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input(LOF(fileNum), #fileNum)
    Close #fileNum

    ' 余計な空白や改行を除去
    base64String = Trim(fileContent)

    If Len(base64String) = 0 Then
        MsgBox "テキストファイルの内容が空です。", vbExclamation
        Exit Sub
    End If

    ' Base64文字列のデコード
    On Error GoTo DecodeError
    decodedBytes = Base64Decode(base64String)

    If Not IsArray(decodedBytes) Or UBound(decodedBytes) < 0 Then
        MsgBox "デコードされたデータが空です。Base64文字列の内容をご確認ください。", vbCritical
        Exit Sub
    End If

    ' 保存先の指定
    savePath = Application.GetSaveAsFilename(InitialFileName:="output.bin", _
                                               FileFilter:="All Files (*.*),*.*", _
                                               Title:="デコード後のファイルの保存先とファイル名を指定してください")
    If savePath = False Then
        MsgBox "ファイルの保存がキャンセルされました。", vbInformation
        Exit Sub
    End If

    ' Printによる上書き防止

    savePath = GetUniqueFileName(CStr(savePath))

    ' バイナリファイルとして保存
    fileNum = FreeFile
    On Error GoTo FileError
    Open savePath For Binary Access Write As #fileNum
    Put #fileNum, , decodedBytes
    Close #fileNum

    MsgBox "デコード完了！" & vbCrLf & "ファイルの保存が完了しました。", vbInformation
    Exit Sub

DecodeError:
    MsgBox "Base64文字列のデコードに失敗しました。形式が正しいかご確認ください。", vbCritical
    Exit Sub

FileError:
    MsgBox "ファイルの保存中にエラーが発生しました。保存先のパスや権限を確認してください。", vbCritical
    If fileNum <> 0 Then Close #fileNum
End Sub

' バイナリファイルを読み込み、Base64形式のテキストファイルとして保存する処理
Sub EncodeFileToBase64()
    Dim filePath As Variant
    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim encodedString As String
    Dim savePath As Variant
    Dim fileSize As Long

    ' ユーザーへの案内
    MsgBox "【エンコード処理】" & vbCrLf & _
           "バイナリファイルを選択し、Base64形式のテキストファイルとして保存します。", vbInformation

    ' バイナリファイル選択
    filePath = Application.GetOpenFilename(FileFilter:="All Files (*.*),*.*", _
                                             Title:="エンコードするバイナリファイルを選択してください")
    If filePath = False Then
        MsgBox "ファイルの選択がキャンセルされました。", vbInformation
        Exit Sub
    End If

    ' バイナリファイルの読み込み
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    If fileSize = 0 Then
        MsgBox "選択されたファイルが空です。", vbExclamation
        Close #fileNum
        Exit Sub
    End If
    ReDim fileData(0 To fileSize - 1) As Byte
    Get #fileNum, , fileData
    Close #fileNum

    ' Base64形式にエンコード
    On Error GoTo EncodeError
    encodedString = Base64Encode(fileData)

    If Len(encodedString) = 0 Then
        MsgBox "エンコードされたデータが空です。", vbCritical
        Exit Sub
    End If

    ' 保存先の指定
    savePath = Application.GetSaveAsFilename(InitialFileName:="encoded.txt", _
                                               FileFilter:="Text Files (*.txt), *.txt", _
                                               Title:="エンコード後のBase64文字列の保存先とファイル名を指定してください")
    If savePath = False Then
        MsgBox "ファイルの保存がキャンセルされました。", vbInformation
        Exit Sub
    End If

    ' Printによる上書き防止
    ' savePath = GetUniqueFileName(savePath)
    savePath = GetUniqueFileName(CStr(savePath))

    ' Base64文字列をテキストファイルとして保存
    fileNum = FreeFile
    On Error GoTo FileSaveError
    Open savePath For Output As #fileNum
    Print #fileNum, encodedString
    Close #fileNum

    MsgBox "エンコード完了！" & vbCrLf & "Base64形式のテキストファイルの保存が完了しました。", vbInformation
    Exit Sub

EncodeError:
    MsgBox "ファイルのエンコードに失敗しました。", vbCritical
    Exit Sub

FileSaveError:
    MsgBox "エンコード結果の保存中にエラーが発生しました。保存先のパスや権限を確認してください。", vbCritical
    If fileNum <> 0 Then Close #fileNum
End Sub

Private Function GetUniqueFileName(savePath As String) As String
    Dim fileDir As String, baseName As String, ext As String
    Dim pos As Long, dotPos As Long, counter As Integer
    Dim newPath As String

    ' ディレクトリ部分の取得
    pos = InStrRev(savePath, "\")
    fileDir = Left(savePath, pos)

    ' 拡張子と基本ファイル名の抽出
    dotPos = InStrRev(savePath, ".")
    If dotPos > pos Then
        baseName = Mid(savePath, pos + 1, dotPos - pos - 1)
        ext = Mid(savePath, dotPos)
    Else
        baseName = Mid(savePath, pos + 1)
        ext = ""
    End If

    newPath = savePath
    counter = 1
    ' 同名ファイルが存在する限り、連番付きの名前に変更する
    Do While Dir(newPath) <> ""
        newPath = fileDir & baseName & " (" & counter & ")" & ext
        counter = counter + 1
    Loop

    GetUniqueFileName = newPath
End Function


