VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private InitTrue As Boolean
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Private AddTbl(255, 255) As Byte
Private XTbl(255, 255) As Byte
Private LsTbl(255, 255) As Byte
Private RsTbl(255, 255) As Byte

Private Sub InitTbl()
If InitTrue = True Then Exit Sub
Dim i As Integer
Dim j As Integer
Dim k As Integer
For i = 0 To 255
    For j = 0 To 255
        XTbl(i, j) = CByte(i Xor j)
        AddTbl(i, j) = CByte((i + j) Mod 255)
    Next j
Next i
InitTrue = True
End Sub
Private Sub Append(ByRef StringData As String, Optional Length As Long)
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Private Function DeHex(Data As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next
    DeHex = GData
    Reset
End Function
Public Function EnHex(Data As String) As String
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    EnHex = GData
    Reset
End Function
Private Function FileExist(FileName As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    FileExist = False
    Exit Function

begin:
    Call FileLen(FileName)
    FileExist = True
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Function EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    EncryptFile = False
    Exit Function
    
begin:
    If FileExist(InFile) = False Then
        EncryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        EncryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte, bKey() As Byte, bOut() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO))
        Buffer(LOF(1)) = 32
        Get #FileO, , Buffer()
    Close #FileO
    
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = EncryptByte(Buffer(), bKey())
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , bOut()
    Close #FileO
    EncryptFile = True
End Function
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputInHex As Boolean) As String
    Dim byteArray() As Byte, bKey() As Byte, bOut() As Byte
    Text = Text & " "
    byteArray() = StrConv(Text, vbFromUnicode)
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = EncryptByte(byteArray(), bKey())
    EncryptString = StrConv(bOut(), vbUnicode)
    If OutputInHex = True Then EncryptString = EnHex(EncryptString)
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextInHex As Boolean) As String
    Dim byteArray() As Byte, bKey() As Byte, bOut() As Byte
    If IsTextInHex = True Then Text = DeHex(Text)
    byteArray() = StrConv(Text, vbFromUnicode)
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = DecryptByte(byteArray(), bKey())
    DecryptString = StrConv(bOut(), vbUnicode)
End Function
Public Function DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    DecryptFile = False
    Exit Function
    
begin:
    If FileExist(InFile) = False Then
        DecryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        DecryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte, bKey() As Byte, bOut() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = DecryptByte(Buffer(), bKey())
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , bOut()
    Close #FileO
    DecryptFile = True
End Function
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Public Function EncryptByte(ds() As Byte, pass() As Byte)
Call InitTbl
Dim tmp2() As Byte
Dim p As Integer
Dim i As Long
Dim Bound As Integer
ReDim tmp2((UBound(ds)) + 4)
Randomize
tmp2(0) = Int((Rnd * 254) + 1)
tmp2(1) = Int((Rnd * 254) + 1)
tmp2(2) = Int((Rnd * 254) + 1)
tmp2(3) = Int((Rnd * 254) + 1)
tmp2(4) = Int((Rnd * 254) + 1)

Call CopyMem(tmp2(5), ds(0), UBound(ds))
ReDim ds(UBound(tmp2)) As Byte
ds() = tmp2()
ReDim tmp2(0)
Bound = (UBound(pass) - 1)
p = 0

For i = 0 To UBound(ds) - 1
    If p = Bound Then p = 0
    ds(i) = XTbl(ds(i), AddTbl(ds(i + 1), pass(p)))
    ds(i + 1) = XTbl(ds(i), ds(i + 1))
    ds(i) = XTbl(ds(i), AddTbl(ds(i + 1), pass(p + 1)))
    p = p + 1
Next i

EncryptByte = ds()
End Function
Public Function DecryptByte(ds() As Byte, pass() As Byte)
Call InitTbl
Dim tmp2() As Byte
Dim p As Long
Dim i As Long
Dim Bound As Integer
Bound = (UBound(pass) - 1)
p = (UBound(ds)) Mod (UBound(pass) - 1)
For i = (UBound(ds)) To 1 Step -1
    If p = 0 Then p = Bound
    ds(i - 1) = XTbl(ds(i - 1), AddTbl(ds(i), pass(p)))
    ds(i) = XTbl(ds(i - 1), ds(i))
    ds(i - 1) = XTbl(ds(i - 1), AddTbl(ds(i), pass(p - 1)))
    p = p - 1
Next i
tmp2() = ds()
ReDim ds((UBound(tmp2)) - 4) As Byte
Call CopyMem(ds(0), tmp2(5), UBound(ds))
ReDim Preserve ds(UBound(ds) - 1) As Byte
DecryptByte = ds()
End Function
Private Function LShift(ByVal ds As Byte, ByVal n As Byte)
    Dim Lsbyte As Byte
    Dim i As Byte
    n = n Mod 8
    For i = 0 To n
        Lsbyte = 128 * (ds And 1)
        Lsbyte = Lsbyte + ((ds And 254) / 2)
        LShift = Lsbyte
    Next i
End Function
Private Function RShift(ByVal ds As Byte, ByVal n As Byte)
    Dim Rsbyte As Byte
    Dim i As Byte
    n = n Mod 8
    For i = 0 To n
        Rsbyte = ((ds And 128) / 128)
        Rsbyte = Rsbyte + ((ds And 127) * 2)
        RShift = Rsbyte
    Next i
End Function



