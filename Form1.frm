VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mempause Program (1)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim PauseTime, Start, Finish, TotalTime
  If (MsgBox("Klik Yes untuk mempause 5 detik", _
             4 + vbQuestion)) = vbYes Then
     PauseTime = 5   'Set durasi pause/interval
     Start = Timer   'Set waktu mulai.
     Do While Timer < Start + PauseTime
        Label1.Caption = Format(Timer, "#,#")
        DoEvents   'Berikan ke proses lainnya.
         'Jika tidak ingin memberikan ke proses lain,
         'tutup statement DoEvents ini.
     Loop
     Finish = Timer   'Set waktu selesai.
     TotalTime = Finish - Start   'Hitung total waktu
                                  'pause.
     'Tampilkan pesan
     MsgBox "Telah dipause selama " & TotalTime & _
            " detik", vbInformation, "Pause"
  Else
     End
  End If
End Sub


