Attribute VB_Name = "MomomizuParser"
Option Explicit

' MOMOMIZU Container用の構造体
Type MOMOMIZUContainer
    IsValid As Boolean
    Format As String
    Version As Integer
    Filename As String
    MimeType As String
    CreatedAt As String
    Author As String
    OriginalSize As Long
    base64Data As String
    ChecksumMD5 As String
End Type


Public Function ParseMomomizuContainer(ByVal content As String) As MOMOMIZUContainer
    Dim lines() As String
    Dim i As Long, posDataStart As Long, posDataEnd As Long
    Dim container As MOMOMIZUContainer
    Dim key As String, value As String
    container.IsValid = False
    ' 改行コードをwindows向けに統一しておく
    content = Replace(content, vbCr, "")
    content = Replace(content, vbLf, vbCrLf)
    
    lines = Split(content, vbCrLf)

    ' ヘッダー解析
    For i = 0 To UBound(lines)
        If lines(i) = "" Then
            Exit For                             ' 空行ならヘッダー終了
        ElseIf InStr(lines(i), ":") > 0 Then
            key = Trim(Left(lines(i), InStr(lines(i), ":") - 1))
            value = Trim(Mid(lines(i), InStr(lines(i), ":") + 1))

            Select Case key
            Case "Format": container.Format = value
            Case "Version": container.Version = CInt(value)
            Case "Filename": container.Filename = value
            Case "MimeType": container.MimeType = value
            Case "CreatedAt": container.CreatedAt = value
            Case "Author": container.Author = value
            Case "OriginalSize": container.OriginalSize = CLng(value)
            End Select
        End If
    Next i

    ' データ部の開始位置を探す
    For i = i + 1 To UBound(lines)
        If lines(i) = "---BEGIN DATA---" Then
            posDataStart = i + 1
        ElseIf lines(i) = "---END DATA---" Then
            posDataEnd = i - 1
            Exit For
        End If
    Next i

    If posDataStart = 0 Or posDataEnd = 0 Then Exit Function

    ' データ部分
    Dim base64Data As String
    base64Data = ""
    For i = posDataStart To posDataEnd
        base64Data = base64Data & Trim(lines(i))
    Next i

    ' データ部分 sanitize
    base64Data = Replace(base64Data, vbCr, "")
    base64Data = Replace(base64Data, vbLf, "")
    base64Data = Replace(base64Data, " ", "")

    container.base64Data = base64Data
    ' analyze footer (md5 checksum)
    For i = posDataEnd + 2 To UBound(lines)
        If InStr(lines(i), "Checksum-MD5:") > 0 Then
            container.ChecksumMD5 = Trim(Mid(lines(i), Len("Checksum-MD5:") + 1))
            Exit For
        End If
    Next i
    
    ' 最終的な妥当性チェック
    If container.Format = "MMMZ" And container.Filename <> "" And container.base64Data <> "" Then
        container.IsValid = True
    End If

    ParseMomomizuContainer = container
End Function


Public Sub TestParseMomomizu()
    Dim testContent As String
    testContent = "Format: MMMZ" & vbCrLf & _
                  "Version: 1" & vbCrLf & _
                  "Filename: test.png" & vbCrLf & _
                  "MimeType: image/png" & vbCrLf & _
                  "CreatedAt: 2025-03-07T12:00:00Z" & vbCrLf & _
                  "OriginalSize: 12345" & vbCrLf & vbCrLf & _
                  "---BEGIN DATA---" & vbCrLf & _
                  "dGVzdCBkYXRhIGhlcmU=" & vbCrLf & _
                  "---END DATA---" & vbCrLf & _
                  "Checksum-MD5: abcdef1234567890abcdef1234567890"

    Dim result As MOMOMIZUContainer
    result = ParseMomomizuContainer(testContent)

    If result.IsValid Then
        Debug.Print "Parse successful!"
        Debug.Print "Filename: " & result.Filename
        Debug.Print "MimeType: " & result.MimeType
        Debug.Print "CreatedAt: " & result.CreatedAt
        Debug.Print "OriginalSize: " & result.OriginalSize
        Debug.Print "Base64Data: " & result.base64Data
        Debug.Print "ChecksumMD5: " & result.ChecksumMD5
    Else
        Debug.Print "Parse failed."
    End If
End Sub


