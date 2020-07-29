VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Selisih Dua Buah Tanggal (2)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  MsgBox SelisihTanggal(CDate("01/05/1999"), CDate("15/09/2002"))
  'Contoh ini menghasilkan: 3.4 --> artinya: 3 tahun 4
  'bulan.
End Sub

Public Function SelisihTanggal(ByVal TanggalAwal As Date, ByVal TanggalAkhir As Date) As String
'Untuk menghitung selisih tahun dan bulan dari dua buah 'tanggal
Dim Tahun As Integer, Sisa As Integer
Dim SelisihBulan As Integer
On Error GoTo Pesan
  SelisihBulan = DateDiff("m", TanggalAwal, TanggalAkhir)
  Tahun = SelisihBulan \ 12
  Sisa = SelisihBulan Mod 12
  SelisihTanggal = Tahun & " Tahun " & Sisa & " Bulan."
  SelisihTanggal = Tahun & "." & Sisa
  Exit Function
Pesan:
  MsgBox "Tipe tanggal salah!", vbCritical, "Error Tanggal"
End Function


