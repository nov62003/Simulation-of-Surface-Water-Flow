VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form fShowData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Ketinggian Pixel dari sebuah Peta"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid msData 
      Height          =   7365
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   9800
      _ExtentX        =   17277
      _ExtentY        =   12991
      _Version        =   393216
      AllowBigSelection=   -1  'True
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "fShowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LDElevasiToGrid()
    Dim i, j As Integer
    
    If panggil = 1 Then
        Me.Caption = "Data Ketinggian Pixel dari sebuah Peta"
        LoadDataElevasi
    End If
    
    msData.Cols = Val(cariBaris) + 1
    msData.Rows = Val(cariKolom) + 1
    msData.TextMatrix(0, 0) = "h"
    msData.ColWidth(0) = 500
    msData.ColAlignment(0) = 3
    
    For i = 1 To msData.Cols - 1
        msData.TextMatrix(0, i) = i
        msData.ColWidth(i) = 700
        msData.ColAlignment(i) = 3
    Next i
    
    For i = 1 To msData.Rows - 1
        msData.TextMatrix(i, 0) = i
    Next i
    
    If panggil = 1 Then
    For i = 1 To cariKolom
        For j = 1 To cariBaris
            msData.TextMatrix(i, j) = matNiDEM(i, j)
        Next j
    Next i
    Else
    Me.Caption = "Jumlah Kunjungan untuk setiap Sel"
    For i = 1 To cariBaris
        For j = 1 To cariKolom
            msData.TextMatrix(i, j) = matJKSel(i, j)
        Next j
    Next i
    End If
End Sub

Private Sub Form_Load()
    Call LDElevasiToGrid
End Sub
