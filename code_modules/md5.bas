Attribute VB_Name = "md5"

Option Explicit

' MD5ハッシュ生成関数（COM参照不要）
'Public Function MD5HashFromBytes(byteData() As Byte) As String
'    Dim enc As Object
'    Dim hashBytes() As Byte
'    Dim i As Long
'
'    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
'    hashBytes = enc.ComputeHash_2(byteData)
'
'    Dim hashHex As String
'    hashHex = ""
'
'    For i = LBound(hashBytes) To UBound(hashBytes)
'        hashHex = hashHex & LCase(Right("0" & Hex(hashBytes(i)), 2))
'    Next i
'
'    MD5HashFromBytes = hashHex
'End Function
'
'Public Sub TestMD5Hash()
'    Dim testData() As Byte
'    testData = StrConv("test data here", vbFromUnicode)
'
'    Dim hashResult As String
'    hashResult = MD5HashFromBytes(testData)
'
'    Debug.Print "MD5 Hash: " & hashResult
'End Sub
