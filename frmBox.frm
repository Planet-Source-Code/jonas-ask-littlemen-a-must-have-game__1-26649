VERSION 5.00
Begin VB.Form frmBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2430
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrCountDown 
      Interval        =   1000
      Left            =   840
      Top             =   2040
   End
   Begin VB.PictureBox picImages 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Index           =   4
      Left            =   3540
      Picture         =   "frmBox.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picImages 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Index           =   3
      Left            =   2700
      Picture         =   "frmBox.frx":1DF2
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picImages 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Index           =   2
      Left            =   1860
      Picture         =   "frmBox.frx":3BE4
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picImages 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Index           =   1
      Left            =   1020
      Picture         =   "frmBox.frx":59D6
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picImages 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Index           =   0
      Left            =   180
      Picture         =   "frmBox.frx":77C8
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   180
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   810
      Width           =   810
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   315
      Left            =   3960
      TabIndex        =   0
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      Height          =   1695
      Left            =   1140
      TabIndex        =   2
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TimeLeft As Integer
Private Sub cmdNo_Click()
    myBoxSvar = "no" 'Flag NO
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    myBoxSvar = "ok" 'Flag OK
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    myBoxSvar = "yes" 'Flag YES
    Me.Hide
End Sub

Private Sub Form_Activate()
    If MaxSpeed Then tmrCountDown.Enabled = True 'if we run at max speed, enavle autoclose
    TimeLeft = 12 'Set countdown timer. in Seconds
End Sub

Private Sub tmrCountDown_Timer()
    If TimeLeft > 0 Then 'Is the time up?
        TimeLeft = TimeLeft - 1 'if not, count down one more
    Else
        tmrCountDown.Enabled = False 'Diable counter
        If frmBox.Visible = False Then Exit Sub 'If the user clicked, nothing will happen
        If cmdOk.Visible Then 'What buttons can we chose form?
            cmdOk_Click 'Clear message
        Else
            cmdNo_Click 'Clear message
        End If
    End If
End Sub
