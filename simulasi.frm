VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fSimulasi 
   BackColor       =   &H00808080&
   Caption         =   "Simulasi Arah Aliran Air menggunakan Algortima MD8 dan DFS"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13740
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   916
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   1429
      ButtonWidth     =   2328
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buka Gambar"
            ImageIndex      =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Set Parameter"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Lihat Data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Proses Simulasi"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Reset Parameter"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Zoom In"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Zoom Out"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   13740
      TabIndex        =   10
      Top             =   7080
      Width           =   13740
      Begin VB.TextBox txtRute 
         Appearance      =   0  'Flat
         Height          =   570
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   420
         Width           =   15000
      End
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   14775
         TabIndex        =   22
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   6270
      Left            =   0
      ScaleHeight     =   414
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   9
      Top             =   810
      Width           =   3555
      Begin VB.CommandButton cmdHis 
         Caption         =   "Lihat History Lintasan"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   6120
         Width           =   3050
      End
      Begin VB.CommandButton cmdTitik 
         Caption         =   "....."
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   5640
         Width           =   3050
      End
      Begin VB.Frame frmSet 
         Caption         =   "Tentukan Area Hujan (Koordinat X, Y) :"
         Height          =   5400
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3300
         Begin VB.ListBox lstData 
            Appearance      =   0  'Flat
            Height          =   3735
            Left            =   105
            TabIndex        =   6
            Top             =   1560
            Width           =   3090
         End
         Begin VB.CommandButton cmdProses 
            Caption         =   "Set Area Hujan"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   3050
         End
         Begin VB.TextBox txtKoorAkhirY 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2160
            TabIndex        =   3
            Text            =   "88"
            Top             =   675
            Width           =   700
         End
         Begin VB.TextBox txtKoorAkhirX 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   720
            TabIndex        =   4
            Text            =   "90"
            Top             =   675
            Width           =   700
         End
         Begin VB.TextBox txtKoorAwalY 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2160
            TabIndex        =   1
            Text            =   "88"
            Top             =   360
            Width           =   700
         End
         Begin VB.TextBox txtKoorAwalX 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   720
            TabIndex        =   2
            Text            =   "90"
            Top             =   360
            Width           =   700
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Y1"
            Height          =   195
            Left            =   1800
            TabIndex        =   20
            Top             =   720
            Width           =   195
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Y0"
            Height          =   195
            Left            =   1800
            TabIndex        =   19
            Top             =   420
            Width           =   195
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "X1"
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   720
            Width           =   195
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "X0"
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   420
            Width           =   195
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   45
         End
      End
      Begin VB.TextBox txtDayaSerap 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2520
         TabIndex        =   15
         Text            =   "0,15"
         Top             =   600
         Width           =   700
      End
      Begin VB.TextBox txtVolAir 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   16
         Text            =   "5"
         Top             =   960
         Width           =   700
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volume Air"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1020
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daya Serap"
         Height          =   195
         Left            =   2160
         TabIndex        =   17
         Top             =   1020
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   24
         X2              =   216
         Y1              =   96
         Y2              =   96
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pGambar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   3645
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   667
      TabIndex        =   8
      Top             =   885
      Width           =   10005
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "simulasi.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "simulasi.frx":02DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "simulasi.frx":39E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "simulasi.frx":730F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "simulasi.frx":ABEE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pTemp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   3645
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   667
      TabIndex        =   24
      Top             =   885
      Width           =   10005
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnBuka 
         Caption         =   "Buka Gambar"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnProses 
      Caption         =   "&Proses"
      Begin VB.Menu mnAtur 
         Caption         =   "Set Parameter"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnLihat 
         Caption         =   "Lihat Data"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnJK 
         Caption         =   "Lihat Jumlah Kunjungan Sel"
         Enabled         =   0   'False
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSimulasi 
         Caption         =   "Proses &Simulasi"
         Enabled         =   0   'False
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnReset 
         Caption         =   "&Reset Parameter"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "fSimulasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ia As Integer, m As Integer, n As Integer, ncc, P, stp As Integer
Dim iii As Integer, ketemu As Boolean, strXY, warnaAliran
Dim Titik, nextPoint, jJalur, scbg, teksInfo, ZInOut

Dim xyPath As Variant, jalur(8, 10000) As Variant, data(), xyPath1, xyPath2
Dim step As Integer, state As Integer, statusStep() As Integer, alurAir(1000, 1000)
Dim XX(1000), YY(1000), MJalur(10000, 1000), pStack(1000, 8) As Variant, xyPanjang

Private Sub aktif()
    frmSet.Enabled = True
    txtKoorAwalX.Enabled = True
    txtKoorAwalY.Enabled = True
    txtKoorAkhirX.Enabled = True
    txtKoorAkhirY.Enabled = True
    cmdProses.Enabled = True
    lstData.Enabled = True

    txtKoorAwalX.BackColor = &H80000005
    txtKoorAwalY.BackColor = &H80000005
    txtKoorAkhirX.BackColor = &H80000005
    txtKoorAkhirY.BackColor = &H80000005
    lstData.BackColor = &H80000005
End Sub

Private Sub nonAktif()
    frmSet.Enabled = False
    txtKoorAwalX.Enabled = False
    txtKoorAwalY.Enabled = False
    txtKoorAkhirX.Enabled = False
    txtKoorAkhirY.Enabled = False
    cmdProses.Enabled = False
    lstData.Enabled = False
    
    txtKoorAwalX.BackColor = &H8000000F
    txtKoorAwalY.BackColor = &H8000000F
    txtKoorAkhirX.BackColor = &H8000000F
    txtKoorAkhirY.BackColor = &H8000000F
    lstData.BackColor = &H8000000F
End Sub

Private Sub OpenFile()
    On Error GoTo ErrCut

    '------- To open a file ---------------------
    cd1.CancelError = True
    cd1.Filter = "JPEG|*.jpg"
    cd1.ShowOpen
    
    namaFileGambar = cd1.FileName
    If FileLen(namaFileGambar) > 650000 Then
        MsgBox "This file is too large to open."
        Exit Sub
    End If
    
    pGambar.Picture = LoadPicture(namaFileGambar)
    cariBaris = pGambar.Width: cariKolom = pGambar.Height
        
    Dim i As Integer, P As Byte, Q As Byte, CARI As String
    P = Len(namaFileGambar): Q = P - 4
    namaFileData = Left(namaFileGambar, Q)
    P = Len(namaFileData): Q = P - 3
    
    For i = 1 To Len(namaFileData)
        CARI = Mid(namaFileData, i, 1)
        If CARI = "." Then
            namaFileTeks = Left(namaFileData, Q) & "txt"
        End If
    Next i
    
    Call OpenTextFile
    Call LoadDataElevasi
    Exit Sub

ErrCut:
    Close #1
    Exit Sub
End Sub

Private Sub BGambar()
    Call OpenFile
    
    On Error Resume Next
    
    mnAtur.Enabled = True
    mnLihat.Enabled = True
    mnReset.Enabled = True
    
    tb.Buttons(2).Enabled = True
    tb.Buttons(3).Enabled = True
    
    Dim i As Integer, j As Integer
    'Status Step
    ReDim Preserve statusStep(cariBaris, cariKolom) As Integer
    'Status Kunjungan Sel
    ReDim Preserve statusKunjungan(cariBaris, cariKolom) As Byte
    'Status Jumlah Kunjungan Sel
    ReDim Preserve matJKSel(cariBaris, cariKolom) As Integer
    
    For i = 1 To cariBaris
        For j = 1 To cariKolom
            statusStep(i, j) = 0
            statusKunjungan(i, j) = 0
            matJKSel(i, j) = 0
        Next j
    Next i
End Sub

Private Sub cmdHis_Click()
    fHis.Show
End Sub

Private Sub PSimulasi()
    Dim i, j As Integer, a As Integer, XYMax, jlr, clr
    Dim oo, pp As Integer, ST As Long, cwarna As Integer, ci As Integer
    Dim nStr, kolom, baris, Jml(10000), CARI, jmaks, brsKe, ktm
    
    cmdTitik.Enabled = True
    fHis.txtHis.Text = ""
    lblInfo.Caption = "Proses Simulasi dimulai......"
    lblInfo.ForeColor = vbBlack
    
    ReDim matJKSel(cariBaris, cariKolom) As Integer 'Matriks Akumlasi Kunjungan Sel
    P = 1: cwarna = 1: ci = 0: jmaks = -99
    
    For a = 1 To UBound(data)
    
        pBar.Value = a
        
        clr = aturWarna(ci, cwarna)
        If cwarna Mod 8 = 0 Then
            ci = ci + 1
        End If

        XYMax = Split(data(a), ",")
        strXY = Fix(matNiDEM(XYMax(1), XYMax(2))) & "," & XYMax(1) & "," & XYMax(2) & "," & 0
        jalur(1, P) = strXY
        
        MJalur(1, P) = jalur(1, P)
        
        teksInfo = "Proses : " & a & ", dari titik koordinat : " & XYMax(1) & "," & XYMax(2) & _
              ", dengan nilai ketinggian (h) = " & Fix(matNiDEM(XYMax(1), XYMax(2))) & vbCrLf
        
        lblInfo.ForeColor = clr
        lblInfo.Caption = "Proses ke : " & a & ", dari koordinat : " & XYMax(1) & "," & XYMax(2) '& " berwarna " & warnaAliran
         
        Call M8DFS(1, XYMax(1), XYMax(2), 1, 1)
        'matJKSel(XYMax(1), XYMax(2)) = matJKSel(XYMax(1), XYMax(2)) + 1
        
        'Mengosongkan Matriks status kunjungan untuk proses X dan Y berikutnya
        ReDim statusKunjungan(cariBaris, cariKolom) As Byte
        
        txtRute.Text = ""
        
        baris = 1: kolom = 1
        For i = 1 To 8
            Do While jalur(i, kolom) <> Empty
                xyPath1 = Split(jalur(i, kolom), ",")
                
                jalur(i, kolom) = Empty
                If i > 1 Then jalur(i, kolom - 1) = Empty
                 
                If kolom > jmaks Then jmaks = kolom
                txtRute.Text = txtRute.Text & "-->" & xyPath1(1) & ", " & xyPath1(2)
                Call visualizePath(xyPath1(2), xyPath1(1), clr)
                Call Picture3.Refresh
                
                kolom = kolom + 1
            Loop
            kolom = 2
            If i > 2 Then
                jalur(i, 1) = Empty
            End If
        Next i
        
        pBar.Max = UBound(data)
        P = 1
        cwarna = cwarna + 1
        
        For i = 0 To 3
            'xyPath(i) = Empty: xyPath1(i) = Empty ': xyPath2(i) = Empty
        Next i
        
        baris = 1: kolom = 1
        Do While pStack(baris, kolom) <> Empty
            For i = 1 To 8
                pStack(baris, i) = Empty
            Next i
            baris = baris + 1
        Loop
        
        baris = 1: kolom = 1: jmaks = -99
        Do While MJalur(baris, kolom) <> Empty
            Jml(baris) = Split(MJalur(baris, kolom), ",")
            
            kolom = kolom + 1
            
            If kolom > jmaks Then
                jmaks = kolom - 1: brsKe = baris
            End If
            
            If MJalur(baris, kolom) = Empty Then
                baris = baris + 1
                kolom = 1
            End If
        Loop
        
        For oo = 1 To jmaks
            jJalur = a
            alurAir(a, oo) = MJalur(brsKe, oo)
            nStr = Len(alurAir(a, oo))
            For j = 1 To nStr
                CARI = Mid(alurAir(a, oo), j, 1)
                If CARI = "," Then
                    ktm = Left(alurAir(a, oo), j - 1)
                    Exit For
                End If
            Next j
            teksInfo = teksInfo & "-->" & ktm
        Next oo
        
        teksInfo = teksInfo & " = " & jmaks & vbCrLf
        teksInfo = teksInfo & "dan berakhir pada koordinat : " & Jml(brsKe)(1) & "," & Jml(brsKe)(2) & _
                   ", dengan nilai ketinggian (h) = " & Jml(brsKe)(0) & vbCrLf & vbCrLf
        
        baris = 1: kolom = 1
        Do While MJalur(baris, kolom) <> Empty
            Do While MJalur(baris, kolom) <> Empty
                MJalur(baris, kolom) = Empty
                kolom = kolom + 1
            Loop
            baris = baris + 1
            kolom = 1
        Loop
        
        brsKe = Empty
        'fHis.txtHis.Text = fHis.txtHis.Text & teksInfo
    'Next a
    
    txtRute.Text = ""
    
    Dim Lebar, h(), LuasA, KelP, JariR, vn, V, Q(), Re, SAliran, SA, STuj
    Dim jh, jd, rh(), jQ, rQ()
    
    Lebar = 0.003: vn = 0
    kolom = 1
    'For j = 1 To jJalur
        teksInfo = teksInfo & "Kecepatan aliran air dari satu sel ke sel berikutnya adalah : " & vbCrLf
        
        ReDim Preserve rh(a)
        ReDim Preserve rQ(a)
        Do While alurAir(a, kolom) <> Empty
            ReDim Preserve h(kolom)
            ReDim Preserve Q(kolom)
            SA = Split(alurAir(a, kolom), ",")
            STuj = Split(alurAir(a, kolom + 1), ",")
            
            h(kolom) = (SA(0) - STuj(0)) / (Sqr(((SA(0) - STuj(0)) ^ 2) * ((SA(0) - STuj(0)) ^ 2)))
            jh = (jh + h(kolom)): jd = kolom
            LuasA = Lebar * h(kolom)
            vn = Sqr(((2 * 10 * (SA(0) - STuj(0))) + vn) - 20)
            
            teksInfo = teksInfo & " --> " & Round(vn, 2) & " m/s"
            
            Q(kolom) = vn * LuasA
            jQ = (jQ + Q(kolom))
            
            kolom = kolom + 1
            
            If alurAir(a, kolom + 1) = Empty Then vn = 0: Exit Do
        Loop
        
        teksInfo = teksInfo & vbCrLf & vbCrLf
        rh(a) = Round(jh / jd, 2)
        rQ(a) = Round(jQ / jd, 5)
        jh = 0: jQ = 0
        
        'Lebar = Lebar * kolom
        LuasA = Lebar * rh(a)
        KelP = Lebar + (2 * rh(a))
        JariR = Round(LuasA / KelP, 5)
        V = Round(rQ(a) / LuasA, 3)
        
        Re = Round((4 * V * JariR) / (1.002 * (10 ^ -6)), 0)
        If Re > 4000 Then
            SAliran = "Turbulen"
        ElseIf Re > 2000 Then
            SAliran = "Turbulen dan Laminar"
        Else
            SAliran = "Laminar"
        End If
        
        kolom = 1
        
        'Lebar = 0.003
        
    'Next j
        teksInfo = teksInfo & "Kecepatan rata-rata aliran air adalah : " & V & " m/s" & vbCrLf
        teksInfo = teksInfo & "Angka Reynold yang didapat untuk lintasan ini adalah : " & Re & vbCrLf
        teksInfo = teksInfo & "Dan sifat aliran dari lintasan ini adalah : " & SAliran & vbCrLf & vbCrLf
     fHis.txtHis.Text = fHis.txtHis.Text & teksInfo
    Next a
    
    
    lblInfo.ForeColor = vbBlack
    lblInfo.Caption = "Proses Simulasi selesai......"
    mnJK.Enabled = True
End Sub

Function aturWarna(ByVal dodol, ByVal cw)
    Dim wrn, hitung
    hitung = (cw - (dodol * 8))
        
    Select Case hitung
    Case 1
        wrn = vbRed
        warnaAliran = "Merah"
    Case 2
        wrn = vbYellow
        warnaAliran = "Kuning"
    Case 3
        wrn = RGB(255, 0, 255) 'ungu
        warnaAliran = "Ungu"
    Case 4
        wrn = vbBlue
        warnaAliran = "Biru"
    Case 5
        wrn = &H40&        'Merah Tua
        warnaAliran = "Merah Tua"
    Case 6
        wrn = &H80FF&     'Orange
        warnaAliran = "Orange"
    Case 7
        wrn = &H400000     'Biru Tua
        warnaAliran = "Biru Tua"
    Case 8
        wrn = &HC0C0FF         'merah muda
        warnaAliran = "Merah Muda"
    End Select
    aturWarna = wrn
End Function


Sub visualizePath(ByVal xPath As Double, ByVal yPath As Double, ByVal w)
    'Color the Start node and the chosen path nodes red
    Sleep 10 '1
    pGambar.PSet (xPath * ZInOut, yPath * ZInOut), w 'RGB(255, 0, 0) 'vbRed
End Sub

Sub visualizeCari(ByVal xPathC As Double, ByVal yPathC As Double, ByVal cw)
    'Color the Start node and the chosen path nodes red
    Sleep 100 '0.5
    pGambar.PSet (xPathC * ZInOut, yPathC * ZInOut), cw 'vbYellow
End Sub

Private Sub M8DFS(ByVal level As Long, Optional ByVal xxx, Optional ByVal yyy, Optional ByVal bJ, Optional ByVal cbg)
    Dim selisihTinggi As Double, rJ, rj1, rj2, konter As Integer: konter = 1
    
    On Error GoTo ERR
    
    If level = -1 Then
        Exit Sub
    End If
    
    statusKunjungan(xxx, yyy) = 1
    
    If xxx >= cariBaris Or yyy >= cariKolom Or xxx <= 1 Or yyy <= 1 Then
        Exit Sub
    End If
    
    For m = -1 To 1
        For n = -1 To 1
            selisihTinggi = Fix(matNiDEM(xxx, yyy)) - Fix(matNiDEM(xxx + m, yyy + n))
            If selisihTinggi > 0 Then
                strXY = Fix(matNiDEM(xxx + m, yyy + n)) & "," & xxx + m & "," & yyy + n
                pStack(level, konter) = strXY & "," & konter
                
                konter = konter + 1
                
                'Akumulasi Arah Aliran
                matJKSel(xxx + m, yyy + n) = matJKSel(xxx + m, yyy + n) + 1
            End If
        Next n
    Next m

    
    P = P + 1
    For iii = konter To 1 Step -1
        If pStack(level, iii - 1) <> Empty Then
            xyPath = Split(pStack(level, iii - 1), ",")
            
            If statusKunjungan(xyPath(1), xyPath(2)) <> 1 Then
                level = level + 1
                jalur(cbg, P) = Fix(matNiDEM(xyPath(1), xyPath(2))) & "," & xyPath(1) & "," & xyPath(2) & "," & level
                
                MJalur(bJ, level) = jalur(cbg, P)
            
                Call M8DFS(level, xyPath(1), xyPath(2), bJ, cbg)
            End If
        Else
            For ia = 1 To level
                level = level - 1
                
                If level = 0 Or level = -1 Then ketemu = True: Exit For
                
                If level = 1 Then
                    cbg = cbg + 1: P = 1: scbg = cbg
                    jalur(cbg, P) = jalur(1, P)
                    P = P + 1
                End If
                
                For rj2 = 1 To 8
                    If pStack(level, rj2) <> Empty Then
                        ncc = ncc + 1
                    End If
                Next rj2
                
                For rJ = ncc To 1 Step -1
                    xyPath2 = Split(pStack(level, rJ), ",")
                    
                    If statusKunjungan(xyPath2(1), xyPath2(2)) <> 1 Then
                        level = level + 1
                        jalur(cbg, P) = Fix(matNiDEM(xyPath2(1), xyPath2(2))) & "," & xyPath2(1) & "," & xyPath2(2) & "," & level
                       
                        ncc = 0
                        bJ = bJ + 1
                        
                        If bJ = 11 Then
                            If level = 36 Then
                            Me.Caption = "t"
                            End If
                        End If
                        
                        For rj1 = 1 To level - 1
                            MJalur(bJ, rj1) = MJalur(bJ - 1, rj1) 'jalur(cbg, rj1)
                        Next rj1
                        MJalur(bJ, level) = jalur(cbg, P)
                        
                        Call M8DFS(level, xyPath2(1), xyPath2(2), bJ, cbg)
                    End If
                Next rJ
                
                ncc = 0
            Next ia
            
            If ketemu = True Then Call M8DFS(-1)
        End If
    Next iii
    
    Exit Sub
    
ERR:

    Dim baris, kolom, i, jmaks, clr
    
    clr = vbRed
    baris = 1: kolom = 1
    For i = 1 To 8
        Do While jalur(i, kolom) <> Empty
            xyPath1 = Split(jalur(i, kolom), ",")
            
            jalur(i, kolom) = Empty
            If i > 1 Then jalur(i, kolom - 1) = Empty
                 
            If kolom > jmaks Then jmaks = kolom
            txtRute.Text = txtRute.Text & "-->" & xyPath1(1) & ", " & xyPath1(2)
            If i > 1 Then clr = vbYellow
            Call visualizePath(xyPath1(2), xyPath1(1), clr)
            Call Picture3.Refresh
                
            kolom = kolom + 1
        Loop
        kolom = 2
        If i > 2 Then
            jalur(i, 1) = Empty
        End If
    Next i

    baris = 1: kolom = 1
    Do While pStack(baris, kolom) <> Empty
        For i = 1 To 8
            pStack(baris, i) = Empty
        Next i
        baris = baris + 1
    Loop
    
    If ERR.Number = 28 Then Call M8DFS(-1)
    Exit Sub

End Sub

Private Sub cmdTitik_Click()
    'nextPoint = nextPoint + 1
    cmdHis.Enabled = True
    
    Dim baris, kolom, B, clr
    baris = 1: kolom = 1
    
    For B = 1 To jJalur
    
        If B Mod 2 = 0 Then
            clr = vbBlue
        Else
            clr = vbCyan
        End If
        
        Do While alurAir(B, kolom) <> Empty
            xyPanjang = Split(alurAir(B, kolom), ",")
            Call visualizeCari(xyPanjang(2), xyPanjang(1), clr)
            kolom = kolom + 1
        Loop
        kolom = 1
        
    Next B
    
    For B = 0 To 3
        xyPanjang(B) = Empty
    Next B
    
    kolom = 1
    For B = 0 To jJalur
        Do While alurAir(B, kolom) <> Empty
            alurAir(B, kolom) = Empty
            kolom = kolom + 1
        Loop
        kolom = 1
    Next B
    
    pGambar.AutoRedraw = True
End Sub

Private Sub mnJK_Click()
    panggil = 2
    fShowData.Show
End Sub

Private Sub mnSimulasi_Click()
    Call PSimulasi
End Sub

Private Sub LData()
    panggil = 1
    fShowData.Show
End Sub

Private Sub cmdProses_Click()
    mnSimulasi.Enabled = True
    tb.Buttons(4).Enabled = True
    tb.Buttons(5).Enabled = True
    tb.Buttons(6).Enabled = True
    tb.Buttons(7).Enabled = True
    lstData.Clear
    
    kXAwal = txtKoorAwalX.Text
    kYAwal = txtKoorAwalY.Text
    kXAkhir = txtKoorAkhirX.Text
    kYAkhir = txtKoorAkhirY.Text
    
    If kXAwal = 0 Or kYAwal = 0 Or kXAkhir = 0 Or kYAkhir = 0 Then
        MsgBox "Silahkan tentukan koordinat area hujan", vbOKOnly, "Perhatian"
        Call ResetParameter
        Exit Sub
    End If
    
    LoadDataElevasi
        
    lstData.AddItem "Koordinat Area Hujan (X, Y) :"
    lstData.AddItem "------------------------------------------------------"
    lstData.AddItem "X0 : " & kXAwal & " --- Y0 : " & kYAwal & ", h(X0,Y0) = " & Fix(matNiDEM(kYAwal, kXAwal))
    lstData.AddItem "X1 : " & kXAkhir & " --- Y1 : " & kYAkhir & ", h(X1,Y1) = " & Fix(matNiDEM(kYAkhir, kXAkhir))
    lstData.AddItem "------------------------------------------------------"
    lstData.AddItem ""
    
    
    Dim r As Integer, c As Integer
    For r = kXAwal To kXAkhir
        For c = kYAwal To kYAkhir
            pGambar.PSet (r * ZInOut, c * ZInOut), &HFFFF00 'vbGreen '&H8000&
            statusStep(r, c) = 1  'semua sel yang berwarna Cyan berstatus 1
        Next c
    Next r
    
    Dim ii As Integer, jj As Integer, k As Integer, B As Integer
    k = 1: stp = 1
    For ii = 1 To cariBaris
        For jj = 1 To cariKolom
            '1. Cari semua sel yang berstatus 1, sel yang berisi air yang dapat mengalir
            If statusStep(ii, jj) = stp Then
                XX(jj) = jj: YY(jj) = ii
                
                ReDim Preserve data(k)
                data(k) = Fix(matNiDEM(XX(jj), YY(jj))) & "," & XX(jj) & "," & YY(jj)

                'Mengurutkan data (Besar ke Kecil)
                Call SortDesc(data)
                k = k + 1
            End If
        Next jj
    Next ii
    
    lstData.AddItem "Jumlah koordinat area hujan = " & UBound(data)
    lstData.AddItem ""
    lstData.AddItem "(X,Y) dan (h) diurutkan :"
    lstData.AddItem "------------------------------------------------------"
    Dim xyu(100000)
    For B = 1 To UBound(data)
        xyu(B) = Split(data(B), ",")
        lstData.AddItem B & ". " & xyu(B)(2) & "," & xyu(B)(1) & " = " & xyu(B)(0)
    Next B
    lstData.AddItem "------------------------------------------------------"
    
    lstData.SetFocus
    lstData.ListIndex = 10
    
    'nextPoint = 1
End Sub

Function SortDesc(UData())
    Dim tmp, i, j, n As Integer
    n = UBound(UData)
    For i = 1 To n + 1
        For j = (i + 1) To n
            If Val(UData(j)) >= Val(UData(i)) Then
                tmp = UData(i)
                UData(i) = UData(j)
                UData(j) = tmp
            End If
        Next j
    Next i
End Function

Private Sub RsParameter()
    Call ResetParameter
    
    lblInfo.Caption = "Info"
    txtRute.Text = ""
    
    Dim i
    ReDim statusStep(cariBaris, cariKolom) As Integer
    
    For i = 1 To UBound(data)
        data(i) = Empty
    Next
End Sub

Private Sub SParameter()
    ReDim statusStep(cariBaris, cariKolom) As Integer
    
    Call aktif
End Sub

Private Sub Form_Load()
    Call nonAktif
End Sub

Private Sub mnAtur_Click()
    Call aktif
End Sub

Private Sub mnBuka_Click()
    Call OpenFile
    
    mnAtur.Enabled = True
    mnLihat.Enabled = True
    mnReset.Enabled = True
    
    tb.Buttons(2).Enabled = True
    tb.Buttons(3).Enabled = True
End Sub

Private Sub mnExit_Click()
    End
End Sub

Private Sub mnLihat_Click()
    panggil = 1
    fShowData.Show
End Sub

Private Sub ResetParameter()
    nonAktif
    
    mnAtur.Enabled = False
    mnLihat.Enabled = False
    mnReset.Enabled = False
    mnSimulasi.Enabled = False
    mnJK.Enabled = False
    
    tb.Buttons(2).Enabled = False
    tb.Buttons(3).Enabled = False
    tb.Buttons(4).Enabled = False
    tb.Buttons(5).Enabled = False
    

    txtKoorAwalX.Text = 0
    txtKoorAwalY.Text = 0
    
    txtKoorAkhirX.Text = 0
    txtKoorAkhirY.Text = 0
    
    namaFileGambar = ""
    namaFileData = ""
    namaFileTeks = ""
    pGambar.Picture = LoadPicture("")
    pGambar.Width = 627: pGambar.Height = 400: pGambar.Top = 59
    
    cariKolom = 0: cariBaris = 0
    
    lstData.Clear
    pBar.Value = 0
    
End Sub

Private Sub mnReset_Click()
    Call ResetParameter
End Sub

Private Sub pGambar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtKoorAwalX.Text = koordX
    txtKoorAwalY.Text = koordY
End Sub

Private Sub pGambar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    koordX = MouseX(pGambar.hWnd)
    koordY = MouseY(pGambar.hWnd)
    
    If namaFileGambar = "" Then
        Me.Caption = "X = 0 dan Y = 0"
        Me.Caption = "X = 0 dan Y = 0"
    Else
        Me.Caption = "X = " & koordX & ", Y = " & koordY
        Me.Caption = "X = " & koordX & ", Y = " & koordY
        pGambar.ToolTipText = "Area Hujan"
    End If
End Sub

Private Sub pGambar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If koordX < 0 Then
        txtKoorAkhirX.Text = 0
        If koordY < 0 Then
            txtKoorAkhirY.Text = 0
            Exit Sub
        End If
        If koordY > pGambar.Height Then
            txtKoorAkhirY.Text = pGambar.Height
            Exit Sub
        End If
        txtKoorAkhirY.Text = koordY
        Exit Sub
    ElseIf koordY < 0 Then
        txtKoorAkhirX.Text = koordX
        If koordX > pGambar.Width Then
            txtKoorAkhirX.Text = pGambar.Width
            txtKoorAkhirY.Text = 0
        End If
        Exit Sub
    ElseIf koordX > pGambar.Width Then
        txtKoorAkhirX.Text = pGambar.Width
        If koordY > pGambar.Height Then
            txtKoorAkhirY.Text = pGambar.Height
            Exit Sub
        End If
        txtKoorAkhirY.Text = koordY
        Exit Sub
    ElseIf koordY > pGambar.Height Then
        If koordX > pGambar.Width Then
            txtKoorAkhirX.Text = pGambar.Width
            txtKoorAkhirY.Text = pGambar.Height
            Exit Sub
        End If
        txtKoorAkhirX.Text = koordX
        txtKoorAkhirY.Text = pGambar.Height
        Exit Sub
    End If
    
    txtKoorAkhirX.Text = koordX
    txtKoorAkhirY.Text = koordY
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
    pGambar.AutoRedraw = False
    
    Select Case Button
    Case "Buka Gambar"
        Call BGambar
        
        With pTemp 'source
            .AutoRedraw = True
            .ScaleMode = vbPixels
            .Visible = False
            .AutoSize = True
            .Picture = LoadPicture(namaFileGambar)
        End With
    
        With pGambar 'dest
            'I've heard this improves quality
            SetStretchBltMode .hdc, HALFTONE
            .ScaleMode = vbPixels
            .Move 243, 59, pTemp.Width, pTemp.Height
            .Picture = pTemp.Picture
        End With
        ZInOut = 1
    
    Case "Set Parameter"
        Call SParameter
        
    Case "Lihat Data"
        Call LData
    
    Case "Proses Simulasi"
        Call PSimulasi
    
    Case "Reset Parameter"
        Call RsParameter
        
        cmdTitik.Enabled = False
        cmdHis.Enabled = False
        tb.Buttons(6).Enabled = False
        tb.Buttons(7).Enabled = False
        
    Case "Zoom In"
        With pGambar
            .Move 243, 59, .Width * 1.2, .Height * 1.2
            .Cls
            StretchBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, pTemp.hdc, 0, 0, pTemp.ScaleWidth, pTemp.ScaleHeight, vbSrcCopy
        End With
        ZInOut = ZInOut * 1.2
        
    Case "Zoom Out"
        With pGambar
            .Move 243, 59, .Width / 1.2, .Height / 1.2
            .Cls
            StretchBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, pTemp.hdc, 0, 0, pTemp.ScaleWidth, pTemp.ScaleHeight, vbSrcCopy
        End With
        ZInOut = ZInOut / 1.2
    End Select
End Sub
