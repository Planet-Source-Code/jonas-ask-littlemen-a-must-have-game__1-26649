VERSION 5.00
Begin VB.Form frmAuthor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmAuthor 
      Caption         =   "Yes, who did? All we know is that it's one of these persons. Do you know?"
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8295
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1875
         Index           =   8
         Left            =   1620
         Picture         =   "frmAuthor.frx":0000
         ScaleHeight     =   123
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   9
         Top             =   240
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   7
         Left            =   1620
         Picture         =   "frmAuthor.frx":9066
         ScaleHeight     =   107
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   8
         Top             =   2160
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1725
         Index           =   6
         Left            =   3180
         Picture         =   "frmAuthor.frx":10E0C
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   7
         Top             =   1920
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   5
         Left            =   6300
         Picture         =   "frmAuthor.frx":192BA
         ScaleHeight     =   99
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   6
         Top             =   2160
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2910
         Index           =   4
         Left            =   4740
         Picture         =   "frmAuthor.frx":20700
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   5
         Top             =   240
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   3
         Left            =   3180
         Picture         =   "frmAuthor.frx":2E842
         ScaleHeight     =   107
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   4
         Top             =   240
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1875
         Index           =   2
         Left            =   6300
         Picture         =   "frmAuthor.frx":365E8
         ScaleHeight     =   123
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   3
         Top             =   240
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   1
         Left            =   60
         Picture         =   "frmAuthor.frx":3F64E
         ScaleHeight     =   119
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   2
         Top             =   1980
         Width           =   1530
      End
      Begin VB.PictureBox PicAut 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1725
         Index           =   0
         Left            =   60
         Picture         =   "frmAuthor.frx":48204
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   1
         Top             =   240
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PicAut_Click(Index As Integer)
    PlaySound "click" 'play a little tune ;)
    Select Case Index
    Case 0: Mybox "Nope, that's not him, that's Thomas.", 2, 0
    Case 1: Mybox "Nope, that is Eirik, not our man this time...", 2, 0
    Case 2: Mybox "Wrong, that there is Chris." & vbNewLine & "He'd wish he made this, though ;)", 2, 0
    Case 3: Mybox "Eh, No, That's Petter... Not him.", 2, 0
    Case 4: Mybox "Yes! That's our man! This is Jonas." & vbNewLine & "He made all this.", 2, 0
    Case 5: Mybox "No... This was not made by Erik", 2, 0
    Case 6: Mybox "Could have been, but it's not Andreas either.", 2, 0
    Case 7: Mybox "Wrong again, Mårten didn't make this game.", 2, 0
    Case 8: Mybox "Sorry, it isn't Knut Petter.", 2, 0
    End Select
    Me.Hide
    Main.SetFocus
End Sub

