Attribute VB_Name = "MD8DFS"
Option Explicit
Option Base 1

Public ArrayTeks()
Public cariKolom As Integer, cariBaris As Integer
Public namaFileGambar As String, namaFileData As String, namaFileTeks As String
Public namaFile3D As String

Public VolAir As Single, DayaSerap As Single
Public matNiDEM() As Single, statusKunjungan() As Byte, matJKSel() As Integer


Public koordX As Long, koordY As Long, DataElevasi() As Single
Public kXAwal As Long, kYAwal As Long, kXAkhir As Long, kYAkhir As Long

Dim r As Integer, c As Integer, i As Integer, j As Integer

Public panggil As Integer

Public Function OpenTextFile()
    Dim nFileNum As Integer, i As Integer
    Dim sNextLine As String, lLineCount As Long, a As Integer
    
    nFileNum = FreeFile
    
    ReDim Preserve ArrayTeks(cariBaris)
    
    Open namaFileData For Input As nFileNum
        lLineCount = 0: a = 1
        Do While Not EOF(nFileNum)
            Line Input #nFileNum, sNextLine
            
            lLineCount = lLineCount + 1
            If lLineCount > 6 Then
                ArrayTeks(a) = sNextLine
                a = a + 1
            End If
        Loop
    Close nFileNum
    
    Open namaFileTeks For Output As nFileNum
        For i = 1 To UBound(ArrayTeks)
            Print #nFileNum, ArrayTeks(i)
        Next i
    Close nFileNum
End Function

Public Sub LoadDataElevasi()
    ReDim matNiDEM(cariKolom, cariBaris) As Single
    
    On Error GoTo ERR
    
    Open namaFileTeks For Input As #3

    For r = 1 To cariKolom
        For c = 1 To cariBaris
            Input #3, matNiDEM(r, c)
        Next c
    Next r
    Close #3
    
    Exit Sub
    
ERR:
    MsgBox "File tidak ditemukan", vbOKOnly, "--- Perhatian"
    Exit Sub
End Sub

Public Sub coba()
    ReDim Preserve matJKSel(cariBaris, cariKolom) As Integer
    
    For i = 1 To cariBaris
        For j = 1 To cariKolom
            matJKSel(i, j) = 0
        Next j
    Next i
End Sub

'Public Sub statusSel(ByVal Status As Byte, ByVal r1, ByVal c1)
    'matSSEl(r1, c1) = Status
'End Sub

'Public Sub JKSel(ByVal r2, ByVal c2)
    'If matSSEl(r2, c2) = 1 Then
    '    matJKSel(r2, c2) = matJKSel(r2, c2) + 1
    'End If
'End Sub
