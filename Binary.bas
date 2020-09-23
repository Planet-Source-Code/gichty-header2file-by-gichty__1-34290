Attribute VB_Name = "Binary"
' Header 2 File by GiChTy
' www.blueadeptz.org
' contact me: GiChTy@blueadeptz.org
' contact us: crew@blueadeptz.org



Const HEADER_LEN = 1024 ' default bytes to seek back from fileend, here: 1 KB


'Functions
Function ReadStringFromFile(File As String, Key As String)
Dim temp As String, i As Long, filenum As Integer, filelenght As Long, strkey As String, strvalue As String, keylen As Long, valuelen As Long, mark1 As Long, mark2 As Long, mark3 As Long
i = 0
temp = 0
mark1 = 0
mark2 = 0
mark3 = 0
filenum = FreeFile
Open File For Binary Access Read As #filenum
Filelength = LOF(filenum)
If Filelength >= HEADER_LEN Then Seek #filenum, Filelength - HEADER_LEN
i = Loc(filenum)
Do Until EOF(filenum)
    temp = Space$(1)
    Get filenum, i, temp
    If temp = "<" Then mark1 = Loc(filenum)
    If temp = "=" Then mark2 = Loc(filenum)
    If temp = ">" Then mark3 = Loc(filenum)
    If mark1 = 0 Or mark2 = 0 Or mark3 = 0 Then GoTo ready
    If mark1 > mark2 Then mark1 = 0
    If mark2 > mark3 Then mark1 = 0: mark2 = 0: mark3 = 0: GoTo ready
    keylen = (mark2 - mark1) - 1
    strkey = ReadString(File, mark1 + 1, keylen)
        If strkey = Key Then
            valuelen = (mark3 - mark2) - 1
            strvalue = ReadString(File, mark2 + 1, valuelen)
            ReadStringFromFile = strvalue
            Close #filenum
            Exit Function
        Else
            mark1 = 0
            mark2 = 0
            mark3 = 0
        End If
ready:
i = i + 1
Loop
Close #filenum
End Function

Sub WriteString2Exe(File As String, Key As String, Value As String)
Dim Filelength As Long
Filelength = FileLen(File) + 1
WriteString File, "<" & Key & "=" & Value & ">", Filelength
End Sub



'Raw-Modes
Function ReadString(File As String, Position As Long, Lenght As Long)
On Error Resume Next
Dim temp As String
temp = Space$(Lenght)
filenum = FreeFile
Open File For Binary Access Read As #filenum
Get filenum, Position, temp
ReadString = temp
Close #filenum
End Function
Sub WriteString(File As String, str As String, Position As Long)
filenum = FreeFile
Open File For Binary Access Write As #filenum
Put #filenum, Position, str
Close #filenum
End Sub
