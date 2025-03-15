Attribute VB_Name = "MomoMizu"
Option Explicit

' バイナリファイルをMOMOMIZU Container形式でエンコードして保存
Public Sub EncodeBinaryFileToMomomizuContainer()
    MsgBox "【エンコード処理】" & vbCrLf & _
           "バイナリファイルを選択し、MOMOMIZU形式のテキストファイルとして保存します。", vbInformation
    
    Dim filePath As Variant, savePath As Variant
    Dim fileNum As Integer, fileSize As Long
    Dim fileData() As Byte, encodedData As String

    filePath = Application.GetOpenFilename(FileFilter:="All Files (*.*),*.*", _
                Title:="エンコードするバイナリファイルを選択してください")

    If filePath = False Then
        MsgBox "ファイルの選択がキャンセルされました。処理を終了します", vbInformation
        Exit Sub
    End If

    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    ReDim fileData(0 To fileSize - 1)
    Get #fileNum, , fileData
    Close #fileNum

    Dim base64String As String
    base64String = Base64Encode(fileData)

    Dim checksum As String
    ' checksum = MD5HashFromBytes(fileData)
    checksum = ""

    Dim containerContent As String
    containerContent = "Format: MMMZ" & vbCrLf & _
                       "Version: 1" & vbCrLf & _
                       "Filename: " & Dir(filePath) & vbCrLf & _
                       "CreatedAt: " & Format(Now, "yyyy-mm-dd\THH:nn:ss\Z") & vbCrLf & _
                       "OriginalSize: " & UBound(fileData) + 1 & vbCrLf & vbCrLf & _
                       "---BEGIN DATA---" & vbCrLf & base64String & vbCrLf & "---END DATA---" & vbCrLf & _
                       "Checksum-MD5: " & checksum

    savePath = Application.GetSaveAsFilename(InitialFileName:=Dir(filePath) & ".txt", _
                FileFilter:="Text Files (*.txt), *.txt", _
                Title:="MOMOMIZU形式として保存")
    
    If savePath = False Then
        MsgBox "ファイルの保存がキャンセルされました。", vbInformation
        Exit Sub
    End If

    ' Printによる上書き防止
    savePath = GetUniqueFileName(CStr(savePath))


    fileNum = FreeFile
    Open savePath For Output As #fileNum
    
    containerContent = Replace(containerContent, vbCr, "")
    containerContent = Replace(containerContent, vbLf, vbCrLf)

    Print #fileNum, containerContent

    Close #fileNum

    MsgBox "ファイルのエンコードと保存が完了しました！" & vbCrLf & vbCrLf & "保存したファイルはMOMOMIZで復元できます", vbInformation
End Sub

' ファイル拡張子判定
Private Function DetectFileExtension(ByRef fileData() As Byte) As String
    Dim magic As String
    ' 先頭32バイト分だけみる
    magic = Left(BytesToHex(fileData), 32)

    Select Case True
        ' PNG: 89 50 4E 47 0D 0A 1A 0A
    Case Left(magic, 16) Like "89504E470D0A1A0A*"
        DetectFileExtension = ".png"
        ' JPEG: FF D8 FF
    Case Left(magic, 6) Like "FFD8FF*"
        DetectFileExtension = ".jpg"
        ' GIF: 47 49 46 38
    Case Left(magic, 8) Like "47494638*"
        DetectFileExtension = ".gif"
        ' PDF: 25 50 44 46
    Case Left(magic, 8) Like "25504446*"
        DetectFileExtension = ".pdf"
        ' ZIP: 50 4B 03 04
    Case Left(magic, 8) Like "504B0304*"
        ' DetectFileExtension = ".zip"
            Dim userInputZIP As String
            userInputZIP = InputBox("このファイルはZIP形式の可能性があります。" & vbCrLf & _
                                  "正しい拡張子を入力してください（zip, docx, pptx, xlsx, マクロ付きのdocm, pptm, xlsm など）" & vbCrLf & "* 入力した拡張子で保存します", "拡張子入力")
            userInputZIP = CleanExtension(userInputZIP)
            DetectFileExtension = userInputZIP
        ' BMP: 42 4D
    Case Left(magic, 4) Like "424D*"
        DetectFileExtension = ".bmp"
        ' TIFF: LE (49 49 2A 00) or BE (4D 4D 00 2A)
    Case Left(magic, 8) Like "49492A00*" Or Left(magic, 8) Like "4D4D002A*"
        DetectFileExtension = ".tiff"
        ' DOC or XLS or PPT (旧Office文書): D0 CF 11 E0 A1 B1 1A E1
    Case Left(magic, 16) Like "D0CF11E0A1B11AE1*"
                    Dim userInputODE As String
            userInputODE = InputBox("このファイルは旧Office文書の可能性があります。" & vbCrLf & _
                                 "正しい拡張子を入力してください（doc, ppt, xls）。" & vbCrLf & "* 入力した拡張子で保存します", "拡張子入力")
            userInputODE = CleanExtension(userInputODE)
            Select Case LCase(Trim(userInputODE))
                Case ".doc", ".ppt", ".xls"
                    DetectFileExtension = userInputODE
                Case Else
                    DetectFileExtension = ".bin"
            End Select
        ' RAR: 52 61 72 21 1A 07 00
    Case Left(magic, 14) Like "526172211A0700*"
        DetectFileExtension = ".rar"
        ' 7z: 37 7A BC AF 27 1C
    Case Left(magic, 12) Like "377ABCAF271C*"
        DetectFileExtension = ".7z"
    Case Left(magic, 8) Like "52494646*"
        If Mid(magic, 9, 8) Like "57415645*" Then
            DetectFileExtension = ".wav"
        ElseIf Mid(magic, 9, 8) Like "41564920*" Then
            DetectFileExtension = ".avi"
        Else
            DetectFileExtension = ".bin"
        End If

        ' MP4/MOVの判定（ftyp）
    Case InStr(magic, "66747970") > 0
        If InStr(magic, "7174") > 0 Then
            DetectFileExtension = ".mov"
        Else
            DetectFileExtension = ".mp4"
        End If

        ' MP3判定
    Case Left(magic, 6) Like "494433*"
        DetectFileExtension = ".mp3"
    Case Left(magic, 4) Like "FFFB*"
        DetectFileExtension = ".mp3"
    Case Else
        DetectFileExtension = ".bin"
    End Select
End Function

Private Function CleanExtension(ByVal extInput As String) As String
    Dim ext As String
    ext = Trim(extInput)
    ext = Replace(ext, ".", "")
    If ext = "" Then
        CleanExtension = ".bin"
    Else
        CleanExtension = "." & LCase(ext)
    End If
End Function



Private Function BytesToHex(ByRef bytes() As Byte) As String
    Dim i As Long, result As String
    For i = LBound(bytes) To Application.Min(UBound(bytes), 7)
        result = result & Right("0" & Hex(bytes(i)), 2)
    Next i
    BytesToHex = result
End Function



Public Sub DecodeMomomizuContainerFile()
    MsgBox "【デコード処理】" & vbCrLf & _
           "MOMOMIZUでエクスポートしたテキストファイルもしくはBase64形式のファイルを選択します。", vbInformation

    Dim filePath As Variant
    filePath = Application.GetOpenFilename(FileFilter:="Text Files (*.txt), *.txt", _
                                           Title:="MOMOMIZU ContainerまたはBase64ファイルを選択してください")
    If filePath = False Then
        MsgBox "ファイルの選択がキャンセルされました。", vbInformation
        Exit Sub
    End If

    Dim fileNum As Integer
    fileNum = FreeFile
    
    Dim line As String, fileContent As String
    fileContent = ""

    Open filePath For Input As #fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        fileContent = fileContent & line & vbCrLf
    Loop
    Close #fileNum
    Dim container As MOMOMIZUContainer
    
    
    container = ParseMomomizuContainer(fileContent)
    Dim savePath As Variant ' GetSaveAsFilenameはpathを返すがキャンセルの場合はFalseを返すらしい...
    Dim decodedBytes() As Byte
    If container.IsValid Then
        decodedBytes = Base64Decode(container.base64Data)
        If container.ChecksumMD5 <> "" Then
            Dim calculatedChecksum As String
            ' calculatedChecksum = MD5HashFromBytes(decodedBytes)
            calculatedChecksum = ""
            If container.ChecksumMD5 <> calculatedChecksum Then
                If MsgBox("チェックサムが一致しません。処理を続けますか？", vbExclamation + vbYesNo) = vbNo Then Exit Sub
            End If
        End If
        filePath = container.Filename
        savePath = Application.GetSaveAsFilename(InitialFileName:=filePath, _
                                                 FileFilter:="(*" & Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1) & "), *" & Right(filePath, Len(filePath) - InStrRev(filePath, ".")), _
                                                 Title:="デコードしたファイルを保存")
    Else
        decodedBytes = Base64Decode(Trim(fileContent))
        Dim ext As String: ext = DetectFileExtension(decodedBytes)
        filePath = "output" & ext
        
        savePath = Application.GetSaveAsFilename(InitialFileName:=filePath, _
                                                 FileFilter:=" ,*" & ext, _
                                                 Title:="デコードしたファイルを保存")
    End If
    
    If savePath = False Then
        MsgBox "ファイルの保存がキャンセルされました。", vbInformation
        Exit Sub
    End If
        
    ' Printによる上書き防止
    savePath = GetUniqueFileName(CStr(savePath))
    
    fileNum = FreeFile
    Open savePath For Binary Access Write As #fileNum
    Put #fileNum, , decodedBytes
    Close #fileNum

    MsgBox "デコード完了！" & vbCrLf & "ファイルを保存しました。", vbInformation
End Sub


Public Function GetUniqueFileName(ByVal savePath As String) As String
    Dim fileDir As String, baseName As String, ext As String
    Dim pos As Long, dotPos As Long, counter As Integer
    Dim newPath As String

    ' directory path
    pos = InStrRev(savePath, "\")
    fileDir = Left(savePath, pos)
    
    ' extension and basename
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
    ' sequencialed incremental name
    Do While Dir(newPath) <> ""
        newPath = fileDir & baseName & " (" & counter & ")" & ext
        counter = counter + 1
    Loop

    GetUniqueFileName = newPath
End Function



