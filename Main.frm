VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LittleMen"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMen 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   3570
      ItemData        =   "Main.frx":030A
      Left            =   10560
      List            =   "Main.frx":030C
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1995
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   10200
      TabIndex        =   23
      Top             =   8520
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   9060
      TabIndex        =   22
      Top             =   8340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdTrap 
      Caption         =   "Place trap"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10980
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame frmBuy 
      Caption         =   "Acme Do-wackies"
      Height          =   2115
      Left            =   10560
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
      Begin VB.Frame Frame1 
         Caption         =   "Univsal Monster Traps"
         Height          =   615
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
         Begin VB.CommandButton cmdBuyTrap 
            Caption         =   "Order"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1080
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Price: 2500"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   300
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "People Kits"
         Height          =   1035
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   1935
         Begin VB.OptionButton chkGender 
            Caption         =   "Man"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton chkGender 
            Caption         =   "Woman"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   480
            Width           =   915
         End
         Begin VB.CommandButton cmdBuyGuy 
            Caption         =   "Order"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1080
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblPrice 
            AutoSize        =   -1  'True
            Caption         =   "Price: 3000"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   810
         End
      End
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave store"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10980
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10500
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build hut"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10980
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Cost: G 1200"
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdFullSpeed 
      Caption         =   "Max Speed"
      Enabled         =   0   'False
      Height          =   315
      Left            =   11580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   10560
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      Height          =   7275
      Left            =   60
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   690
      TabIndex        =   14
      Top             =   540
      Width           =   10410
      Begin VB.TextBox txtConsole 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   315
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   7500
         Width           =   10335
      End
   End
   Begin VB.Label lblMapInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1215
      Left            =   2040
      TabIndex        =   25
      Top             =   7860
      Width           =   6435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   7920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   8640
      TabIndex        =   13
      Top             =   7860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "-------------"
      Height          =   195
      Left            =   1320
      TabIndex        =   17
      Top             =   180
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   3780
      Width           =   1215
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE: these so called "guys" and "men" can be "woman" and "girls". hehe, it's just to simplify the code

Private LastTick As String 'framelimiter

Private Sub cmdBuyGuy_Click()
    If MainPause Then Exit Sub
    If chkGender(0).Value = True Then
        BuyGuy 2 'Buy a Man
    Else
        BuyGuy 1 'Buy a Woman
    End If
End Sub

Private Sub cmdBuyTrap_Click()
    If MainPause Then Exit Sub
    BuyTrap
End Sub

Private Sub cmdFullSpeed_Click()
    If MainPause Then Exit Sub
    MaxSpeed = Not MaxSpeed
End Sub

Private Sub cmdLeave_Click()
    If MainPause Then Exit Sub
    LeaveStore
End Sub

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    cmdBuild.Enabled = True
    cmdPause.Enabled = True
    cmdTrap.Enabled = True
    cmdBuyGuy.Enabled = True
    cmdBuyTrap.Enabled = True
    cmdFullSpeed.Enabled = True
    cmdLeave.Enabled = True
    picMain.Enabled = True
    lstMen.Enabled = True
    MainLoop
End Sub
Private Sub MainLoop()
Dim T(1 To 10) As Long
Dim NowTime As String 'Holds time keeping for the framelimiter
    Do Until EndOfTime 'Hehe ;)
        T(1) = GetTickCount
        DoEvents
        Randomize
        
        Do Until MainPause = False
            DoEvents
        Loop
        
        T(3) = GetTickCount
        NowTime = GetTickCount 'THIS IS A FRAME LIMITER
        Do Until NowTime - LastTick > 80 Or MaxSpeed = True
            DoEvents
            NowTime = GetTickCount
        Loop
        LastTick = NowTime
        
        T(3) = GetTickCount - T(3)
        
        MonsterMove
        
        Pics.PicBuffer(0).Cls
        
        T(5) = GetTickCount
        DoPreGraphics
        T(5) = GetTickCount - T(5)
        
        T(2) = GetTickCount
        For M = 1 To UBound(Men)
            If Men(M).Act Then   'If the guy is in use
                
                If Men(M).Age > 70 Then 'He's in the target group
                    b = Int(10000 - ((Men(M).Age / 10) * 3)) 'The small chance of death
                    If Int(Rnd * 10000) > b Or IIf(Men(M).Age > 101, Rnd > 0.9, F = b) Then
                        PlaySound "die" 'Draw his terminal breath
                        KillGuy M 'Call the Grim Reaper
                        GoTo NextMan
                    End If
                End If
                
                ManMove (M) 'move him
                ManCheckTarget (M) 'see if he has reached his destination
                
                'Paint this gent :)
                If Men(M).Age >= 16 Then 'Paint adult
                    BitBlt Pics.PicBuffer(0).hDC, (Men(M).X - 1) * size, (Men(M).Y - 1) * size, size, size, Pics.PicManM((8 - (Men(M).Gender * 4) + Men(M).Dire - 1)).hDC, 0, 0, SRCAND
                    BitBlt Pics.PicBuffer(0).hDC, (Men(M).X - 1) * size, (Men(M).Y - 1) * size, size, size, Pics.PicMan(8 - (Men(M).Gender * 4) + Men(M).Dire - 1).hDC, 0, 0, SRCPAINT
                Else 'Paint Child
                    BitBlt Pics.PicBuffer(0).hDC, (Men(M).X - 1) * size, (Men(M).Y - 1) * size, size, size, Pics.PicBoyM((8 - (Men(M).Gender * 4) + Men(M).Dire - 1)).hDC, 0, 0, SRCAND
                    BitBlt Pics.PicBuffer(0).hDC, (Men(M).X - 1) * size, (Men(M).Y - 1) * size, size, size, Pics.PicBoy(8 - (Men(M).Gender * 4) + Men(M).Dire - 1).hDC, 0, 0, SRCPAINT
                End If

                If Men(M).Reason = 5 And Men(M).Indoors = True Then 'add blanking
                    BitBlt Pics.PicBuffer(0).hDC, (Men(M).X - 2) * size, (Men(M).Y - 1) * size, size * 3, size, Pics.PicCens.hDC, 0, 0, SRCCOPY
                End If
            End If
NextMan:
        Next M
        T(2) = GetTickCount - T(2)
        
        
        
        For C = 1 To UBound(Caves)
            Caves(C).Food = Caves(C).Food - Caves(C).People
            If Caves(C).Food < 0 Then Caves(C).Food = 0 'We can't eat "luft boller og vente brus" can we? LOL
            
            If Caves(C).Food <= Int((Rnd * 80) + 50) Then 'Got milk?
                For A = 1 To UBound(Caves(C).LiveHere) 'loop trough the poeople living here
                    If Caves(C).LiveHere(A) > 0 Then 'is this a late dude? A stiffy? An ex-parrot? ;)
                        If Men(Caves(C).LiveHere(A)).ThisCave = C And Not Caves(C).FoodOk And TownBell = 0 Then 'see if he's at home, and if some piggy allready went to the marked, and the Townbell is not rung
                            If Men(Caves(C).LiveHere(A)).Age > 7 Or Caves(C).Food = 0 Then 'Old enough to go shoping alone, but if food is 0, go anyway
                                ManGoShoping Caves(C).LiveHere(A) 'Send him SHOPPING!
                                Caves(C).FoodOk = True 'Leave a "gone shopping"-note
                                Exit For
                            End If
                        End If
                    End If
                Next A
            End If
        Next C
        
        'Process Data
        
        If GrowTag = 110 Then
            GrowTag = 0
            MenGrow
        Else
            GrowTag = GrowTag + 1
        End If
                
        
        DoBell
        DoMonster
        CheckforElvis
        MoveTheKing
        DoSigns
        DoKeys
        
        DoInformaition
        
        'Paint the GUI
        T(4) = GetTickCount
        
        
        picMain.Cls
        DoGraphics
        ComposePic
        T(4) = GetTickCount - T(4)
        
        
        lblinfo.Caption = "Gwookees: " & P1.Money & "    " & "Population: " & MenActive
        

        Label4.Caption = TownBell & vbNewLine & M1.Tag & vbNewLine & M1.Tag2 & vbNewLine & MonCountdown
        
        
        Label2.Caption = GetTickCount - T(1) & vbNewLine _
        & "Men: " & T(2) & vbNewLine _
        & "Framelimiter: " & T(3) & vbNewLine _
        & "Grapchis: " & T(4) & vbNewLine _
        & "Pre Grapchis: " & T(5) & vbNewLine
        
        
    Loop
    MsgBox "huh!?! The loop exited! Wow. How did that happen??? Oh well.. :)", vbQuestion, "Something strange"
End Sub

Private Sub cmdPause_Click()
    MainPause = Not MainPause
End Sub

Private Sub cmdBuild_Click()
    If MainPause Then Exit Sub
    If P1.Outside Then
        BuildHut P1.X, P1.Y
    End If
End Sub


Private Sub cmdTrap_Click()
    If MainPause Then Exit Sub
    placeTrap
End Sub

Private Sub Command2_Click()
MonCountdown = 3000
End Sub

Private Sub Command3_Click()
    MakeMonster
End Sub

Private Sub Form_Load()
    Randomize 'Initiate the random generator
    
    LastTick = GetTickCount 'Mark the Framelimiter
    
    Bredde = 46 'Set Witdh
    Hoyde = 32 'Height
    
    ReDim Board(-10 To Bredde + 10, -10 To Hoyde + 10) 'Redim the Building info board the 10's are for safty
    ReDim Board2(-10 To Bredde + 10, -10 To Hoyde + 10) 'Redim the NOMove board
    
    SetUpBuffers 'guess
    SetUpBackground 'Draw the backgound
    
    P1.Money = 500 'Guess this one too
    
    MakeForest
    
    NumCaves = Int((Rnd * 2) + 2) 'Get a random number of starting caves
    ReDim Caves(NumCaves)
    For b = 1 To NumCaves 'make these new caves
        MakeCaves b
    Next b
    
    MonCountdown = Int((Rnd * 1000) + 600) 'Set monster arrive time
    MakeStore 'What could this do?
    PutUpBell 'and this?
    MakePalmTrees 'Call the Polish embassy and declare war. Think so? ;)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Mybox(vbNewLine & "Are you sure you want to quit?", 1, 2) = "yes" Then
        End
    End If
    Cancel = 1
End Sub

Private Sub lstMen_Click()
    ListSel = lstMen.ListIndex
    If ListSel >= 0 Then
        Keeping.Num = ListMen(ListSel + 1)
        Keeping.T = 1
    End If
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ProcessClick X, Y
    Else
        Keeping.Num = 0
        Keeping.T = 0
        Main.lstMen.ListIndex = -1
        ListSel = -1
    End If
End Sub

Private Sub txtConsole_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ConsoleInput txtConsole.Text 'if we hit enter, process the text
End Sub

