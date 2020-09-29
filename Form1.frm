VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hilangkan Spasi di Suatu String"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Created by Rizky Khapidsyah
'Source code program dimulai dari sini

Private Sub Command1_Click()
  'Menggunakan fungsi buatan...
  MsgBox HilangkanSpasi("Selamat Belajar")
End Sub

Private Sub Command2_Click()
Dim Hilangkan As String
  'Menggunakan fungsi Replace miliknya VB...
  Hilangkan = Replace("Selamat Belajat !", " ", "")
  MsgBox Hilangkan
End Sub

Public Function HilangkanSpasi(strKalimat As String) _
As String
Dim i As Integer 'Deklarasi untuk counter
Dim Temp As String 'Deklarasi untuk menampung karakter
Dim Huruf As String * 1 'Deklarasi untuk memeriksa
                        'karakter
  Temp$ = "": Huruf = "" 'Inisialisasi awal variabel
  For i% = 1 To Len(strKalimat) 'Proses sebanyak
 'karakter
    'Tampung setiap satu karakter
    Huruf = Chr(Asc(Mid(strKalimat, i%, 1)))
    'Periksa jika terdapat spasi...
    If Len(Trim(Huruf)) < 1 Then
      'Tidak usah ditampung (tidak usah diproses)
    Else 'Jika tidak terdapat spasi...
      'Tampung seperti biasa setiap satu karakter
      Temp$ = Temp$ + Chr(Asc(Mid(strKalimat, i%, 1)))
    End If 'Akhir pemeriksaan spasi
  Next i 'Ke karakter berikutnya
  'Tampung string yang sudah hilang spasi-nya..
  HilangkanSpasi = Temp$
End Function


