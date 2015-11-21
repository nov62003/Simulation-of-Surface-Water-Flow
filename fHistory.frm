VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form fHis 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histori Lintasan Arah Aliran yang terbentuk"
   ClientHeight    =   8055
   ClientLeft      =   10020
   ClientTop       =   1845
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   5295
   Begin RichTextLib.RichTextBox txtHis 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5300
      _ExtentX        =   9340
      _ExtentY        =   14208
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"fHistory.frx":0000
   End
End
Attribute VB_Name = "fHis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lR As Long
    lR = SetTopMostWindow(Me.hWnd, True)
End Sub
