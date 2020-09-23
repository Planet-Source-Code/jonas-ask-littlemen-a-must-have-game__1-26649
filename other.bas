Attribute VB_Name = "other"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020

Public Type aCave
 X As Integer
 Y As Integer
 People As Integer
 Food As Integer
 FoodOk As Boolean
 Store As Boolean
 LiveHere(1 To 6) As Byte
End Type

Public Type People 'Human DNA!
 X As Integer
 Y As Integer
 TX As Integer
 Pregnant As Integer
 TY As Integer
 Act As Boolean
 Indoors As Boolean
 Return As Boolean
 MyName As String
 Reason As Integer
 Tag As Integer 'these two tags are use for many things, f.eks. how much food did we buy? How long have we been out walking?
 LeaveTime As Integer
 HomeCave As Integer
 TargetCave As Integer
 ThisCave As Integer
 Gender As Byte
 Age As Byte
 Dire As Byte
End Type

Public Type PalmTree
 X As Integer
 Y As Integer
 Tree As Byte
End Type

Public Palms() As PalmTree

Public Men(1 To 100) As People 'The little guys and gals
Public MenActive As Integer 'How many are alive?

Public Bredde As Integer 'Widht of map
Public Hoyde As Integer 'Height of (Guess what?)

Public Board() As Boolean
Public Board2() As Boolean

Public MainPause As Boolean
Public MaxSpeed As Boolean

Public Caves() As aCave
Public Const Maxcaves As Byte = 20
Public Const HutCost As Integer = -1200

Public Store As aCave
Public StoreOpen As Boolean

Public NumCaves As Integer
Public Const MaxInCave As Byte = 6

Public TownBell As Integer

Public Const WaitTimeMax As Integer = 200
Public Const WaitTimeMin As Integer = 20

Public ListSel As Integer
Public ListMen() As Integer

Public Const size As Integer = 15 'Size of a square in pixels


Public Sub MakeCaves(A)

Makenew:
    Randomize Rnd * 255 'Give the cave a random position
    Caves(A).X = Int((Rnd * Bredde) + 1)
    Caves(A).Y = Int((Rnd * Hoyde) + 1)
    If Caves(A).X = 1 Or Caves(A).X = Bredde Or Caves(A).Y = 1 Or Caves(A).Y = Hoyde Then GoTo Makenew 'Let's stay off the edge shall we...

    'This is to prevent the caves from apearing within a radius of RWO from the other STUFF
    On Error Resume Next
    For Y = -2 To 2
    For X = -2 To 2
        If Board(Caves(A).X + X, Caves(A).Y + Y) = True Then GoTo Makenew
    Next X
    Next Y
    'All is ok, add Cave data
    
    
    Randomize 'Do you eat food you find left over at your new house?
    Caves(A).Food = Int((Rnd * 990) + 10)
    

    ManInitialize A, 0, 0, 0, 2 'A man
    ManInitialize A, 0, 0, 0, 1 'and a woman
    If Int((Rnd * 2) + 1) = 1 Then 'More than two people?
        ManInitialize A, 0, 0, 0, 0 'put 'em in!
    End If
    
    Board(Caves(A).X, Caves(A).Y) = True 'Now add the cave into the Board data
    
End Sub
Public Sub MakeStore()
Makenew:
    'Place the Store
    Store.X = Int((Rnd * Bredde) + 1)
    Store.Y = Int((Rnd * Hoyde) + 1)
    If Store.X = 1 Or Store.X = Bredde Or Store.Y = 1 Or Store.Y = Hoyde Then GoTo Makenew 'Let's stay off the edge shall we...
    For Y = -2 To 2 'Prevent the store form apearing over other things
    For X = -2 To 2
        If Board(Store.X + X, Store.Y + Y) = True Then GoTo Makenew
    Next X
    Next Y
    Store.Store = True
    Board(Store.X, Store.Y) = True
    StoreOpen = True
    
    P1.X = Store.X
    P1.Y = Store.Y
End Sub

Public Sub SetUpBuffers()
    For A = 0 To Pics.PicBuffer.Count - 1 'Makes all the buffer the same size as the main picture
        Pics.PicBuffer(A).Width = Main.picMain.Width
        Pics.PicBuffer(A).Height = Main.picMain.Height
    Next A
End Sub
Public Sub SetUpBackground()
    For Y = 1 To Bredde
        For X = 1 To Bredde
            BitBlt Pics.PicBuffer(1).hDC, (X * size) - size, (Y * size) - size, size, size, Pics.PicGround(Int(Rnd * 5)).hDC, 0, 0, SRCCOPY
        Next X
    Next Y
    
    'Add some stones
    N = Int((Rnd * 5) + 10)
    For A = 1 To N
GetNew:
        X = Int((Rnd * Bredde) + 1)
        Y = Int((Rnd * Hoyde) + 1)
        
        For Y1 = -2 To 2
        For X1 = -2 To 2
            If Board(X + X1, Y + Y1) Then GoTo GetNew
            If Board2(X + X1, Y + Y1) Then GoTo GetNew
        Next X1
        Next Y1
        
        T = Int((Rnd * 3) + 1)
        BitBlt Pics.PicBuffer(1).hDC, (X * size) - size - 7, (Y * size) - size - 7, 30, 30, Pics.PicStoneM(T - 1).hDC, 0, 0, SRCAND
        BitBlt Pics.PicBuffer(1).hDC, (X * size) - size - 7, (Y * size) - size - 7, 30, 30, Pics.PicStone(T - 1).hDC, 0, 0, SRCPAINT
    Next A
    
End Sub

Public Sub DoGraphics()
    'The player
    If P1.Outside Then
        BitBlt Pics.PicBuffer(0).hDC, (P1.X * size) - size, (P1.Y * size) - size, size, size, Pics.PicMeM(P1.Dire - 1).hDC, 0, 0, SRCAND
        BitBlt Pics.PicBuffer(0).hDC, (P1.X * size) - size, (P1.Y * size) - size, size, size, Pics.PicMe(P1.Dire - 1).hDC, 0, 0, SRCPAINT
    End If
    If Elvis.Act Then
        BitBlt Pics.PicBuffer(0).hDC, (Elvis.X * size) - size, (Elvis.Y * size) - size, size, size, Pics.PicKingM(Elvis.Dire - 1).hDC, 0, 0, SRCAND
        BitBlt Pics.PicBuffer(0).hDC, (Elvis.X * size) - size, (Elvis.Y * size) - size, size, size, Pics.PicKing(Elvis.Dire - 1).hDC, 0, 0, SRCPAINT
    End If
    'The Monster
    If M1.Act Then
        If M1.LeaveTime > 0 Then 'paint eating
            BitBlt Pics.PicBuffer(0).hDC, (M1.X * size) - size - 7, (M1.Y * size) - size - 7, 30, 30, Pics.PicMonEatM.hDC, 0, 0, SRCAND
            BitBlt Pics.PicBuffer(0).hDC, (M1.X * size) - size - 7, (M1.Y * size) - size - 7, 30, 30, Pics.PicMonEat.hDC, 0, 0, SRCPAINT
        Else 'paint normal monster
            BitBlt Pics.PicBuffer(0).hDC, (M1.X * size) - size - 7, (M1.Y * size) - size - 7, 30, 30, Pics.PicMonM(M1.Dire - 1).hDC, 0, 0, SRCAND
            BitBlt Pics.PicBuffer(0).hDC, (M1.X * size) - size - 7, (M1.Y * size) - size - 7, 30, 30, Pics.PicMon(M1.Dire - 1).hDC, 0, 0, SRCPAINT
        End If
    End If
    
    'The Bell
    BitBlt Pics.PicBuffer(0).hDC, (BellX - 2) * size, (BellY - 1) * size, 45, 15, Pics.picBellM(BellTag).hDC, 0, 0, SRCAND
    BitBlt Pics.PicBuffer(0).hDC, (BellX - 2) * size, (BellY - 1) * size, 45, 15, Pics.picBell(BellTag).hDC, 0, 0, SRCPAINT
    
    
    For A = 1 To UBound(Caves) 'Make the Caves
        Pics.PicHut(0).Cls
        Pics.PicHut(1).Cls
        Pics.PicHut(0).Line (2, 5)-Step((Caves(A).Food + 1) / 33, 2), RGB(196, 222, 232), BF
        Pics.PicHut(1).Line (2, 5)-Step((Caves(A).Food + 1) / 33, 2), vbBlack, BF
        For b = 1 To Caves(A).People 'Mark each Hut's ocupants
            Pics.PicHut(0).Line ((b * 4) + 1, 0)-Step(1, 3), vbBlue, BF
            Pics.PicHut(1).Line ((b * 4) + 1, 0)-Step(1, 3), vbBlack, BF
        Next b
        BitBlt Pics.PicBuffer(0).hDC, (Caves(A).X * size) - size - 7, (Caves(A).Y * size - 7) - size, size * 2, size * 2, Pics.PicHut(1).hDC, 0, 0, SRCAND
        BitBlt Pics.PicBuffer(0).hDC, (Caves(A).X * size) - size - 7, (Caves(A).Y * size - 7) - size, size * 2, size * 2, Pics.PicHut(0).hDC, 0, 0, SRCPAINT
    Next A

    'Make the Store
    Pics.PicHut(3).Cls
    Pics.PicHut(2).Cls
    If StoreOpen = False Then 'Put in a sign to show thazt the store is closed
        BitBlt Pics.PicHut(3).hDC, Pics.PicHut(3).ScaleWidth - 10, 2, 8, 8, Pics.picXM.hDC, 0, 0, SRCAND
        BitBlt Pics.PicHut(2).hDC, Pics.PicHut(3).ScaleWidth - 10, 2, 8, 8, Pics.picXM.hDC, 0, 0, SRCAND
        BitBlt Pics.PicHut(2).hDC, Pics.PicHut(2).ScaleWidth - 10, 2, 8, 8, Pics.picX.hDC, 0, 0, SRCPAINT
    End If
    BitBlt Pics.PicBuffer(0).hDC, (Store.X * size) - size - 7, (Store.Y * size - 7) - size, size * 2, size * 2, Pics.PicHut(3).hDC, 0, 0, SRCAND
    BitBlt Pics.PicBuffer(0).hDC, (Store.X * size) - size - 7, (Store.Y * size - 7) - size, size * 2, size * 2, Pics.PicHut(2).hDC, 0, 0, SRCPAINT
    
    'The Palmtrees
    For A = 1 To UBound(Palms)
        BitBlt Pics.PicBuffer(0).hDC, (Palms(A).X * size) - size - 7, (Palms(A).Y * size - 7) - size, size * 2, size * 2, Pics.PicPalmM(Palms(A).Tree - 1).hDC, 0, 0, SRCAND
        BitBlt Pics.PicBuffer(0).hDC, (Palms(A).X * size) - size - 7, (Palms(A).Y * size - 7) - size, size * 2, size * 2, Pics.PicPalm(Palms(A).Tree - 1).hDC, 0, 0, SRCPAINT
    Next A
    'The Forest
    If Forest.Act Then
        BitBlt Pics.PicBuffer(0).hDC, (Forest.X * size) - size - 7, (Forest.Y * size - 7) - size, 236, 241, Pics.PicForestM.hDC, 0, 0, SRCAND
        BitBlt Pics.PicBuffer(0).hDC, (Forest.X * size) - size - 7, (Forest.Y * size - 7) - size, 236, 241, Pics.picForest.hDC, 0, 0, SRCPAINT
    End If
    'The Arrow
    If Keeping.T > 0 Then
        Select Case Keeping.T
        Case 1
            BitBlt Pics.PicBuffer(0).hDC, (Men(Keeping.Num).X * size) - size, (Men(Keeping.Num).Y * size) - size - size, size, size, Pics.PicArrowM.hDC, 0, 0, SRCAND
            BitBlt Pics.PicBuffer(0).hDC, (Men(Keeping.Num).X * size) - size, (Men(Keeping.Num).Y * size) - size - size, size, size, Pics.PicArrow.hDC, 0, 0, SRCPAINT
        Case 2
            BitBlt Pics.PicBuffer(0).hDC, (Caves(Keeping.Num).X * size) - size, (Caves(Keeping.Num).Y * size) - size - size - 7, size, size, Pics.PicArrowM.hDC, 0, 0, SRCAND
            BitBlt Pics.PicBuffer(0).hDC, (Caves(Keeping.Num).X * size) - size, (Caves(Keeping.Num).Y * size) - size - size - 7, size, size, Pics.PicArrow.hDC, 0, 0, SRCPAINT
        Case 3
            BitBlt Pics.PicBuffer(0).hDC, (M1.X * size) - size, (M1.Y * size) - size - size - 7, size, size, Pics.PicArrowM.hDC, 0, 0, SRCAND
            BitBlt Pics.PicBuffer(0).hDC, (M1.X * size) - size, (M1.Y * size) - size - size - 7, size, size, Pics.PicArrow.hDC, 0, 0, SRCPAINT
        End Select
    End If
    'The Signs
    For A = 1 To UBound(Signs)
        If Signs(A).Act Then 'is this sign in use?
            Pics.picSign.Cls 'clear it
            Pics.picSign.ForeColor = Signs(A).Color 'set color to red or green
            Pics.picSign.Print "G " & Signs(A).Amount 'print the aount of money
            Pics.picSignM.Cls
            Pics.picSignM.Print "G " & Signs(A).Amount
            BitBlt Pics.PicBuffer(0).hDC, (Signs(A).X * size) - size, (Signs(A).Y * size) - 70 + Signs(A).Tag, Pics.picSign.ScaleWidth, Pics.picSign.ScaleHeight, Pics.picSignM.hDC, 0, 0, SRCAND
            BitBlt Pics.PicBuffer(0).hDC, (Signs(A).X * size) - size, (Signs(A).Y * size) - 70 + Signs(A).Tag, Pics.picSign.ScaleWidth, Pics.picSign.ScaleHeight, Pics.picSign.hDC, 0, 0, SRCPAINT
        End If
    Next A
    
    Main.Pic1.Cls
    For X = 1 To Bredde
    For Y = 1 To Hoyde
        If Board(X, Y) Then Main.Pic1.PSet (X, Y)
        If Board2(X, Y) Then Pics.PicBuffer(0).Line ((X - 1) * size, (Y - 1) * size)-Step(5, 5), , BF
    Next
    Next
End Sub

Public Sub ComposePic()
    BitBlt Main.picMain.hDC, 0, 0, Bredde * size, Hoyde * size, Pics.PicBuffer(0).hDC, 0, 0, SRCCOPY
End Sub

Public Sub ArriveAtCave(C)
    Caves(C).People = Caves(C).People + 1
End Sub

Public Sub LeaveCave(C)
    If C = 0 Then Exit Sub
    Caves(C).People = Caves(C).People - 1
End Sub
Public Function GetNewHomeCave()
Dim Tryed() As Boolean
Dim Tell As Integer
    ReDim Tryed(1 To UBound(Caves))
    Randomize
newA:
    Tell = 0
    For b = 1 To UBound(Tryed)
        If Tryed(b) Then Tell = Tell + 1
    Next b
    If Tell = NumCaves Then Stop 'ID DIDN'T HAVE TO COME TO COME TO THIS, YOU KNOW
    
    A = Int((Rnd * NumCaves) + 1)
    For b = 1 To 6
        If Caves(A).LiveHere(b) = 0 Then
            GetNewHomeCave = A
            Exit Function
        End If
    Next b
    Tryed(A) = True: GoTo newA
    
End Function
Public Sub MoveIntoCave(C, M)
    For b = 1 To 6
        If Caves(C).LiveHere(b) = 0 Then
            Caves(C).LiveHere(b) = M
            Exit Sub
        End If
    Next b
End Sub

Sub GoHome(A)
    Men(A).TargetCave = Men(A).HomeCave
    Men(A).TX = Caves(Men(A).HomeCave).X
    Men(A).TY = Caves(Men(A).HomeCave).Y
End Sub

Public Sub CheckforElvis()
    If Elvis.Tag > 0 Then Exit Sub 'Elvis has left the building, for good
    If Board(BellX - 3, BellY) And Board(BellX, BellY + 3) And Board(BellX + 3, BellY) And Board(BellX, BellY - 3) Then
        Elvis.Act = True
        'GET A RANDOM START POINT
        b = Int((Rnd * 4) + 1)
        Select Case b
        Case 1
            Elvis.X = 0
            Elvis.Y = Int((Rnd * Hoyde) + 1)
        Case 2
            Elvis.X = Int((Rnd * Bredde) + 1)
            Elvis.Y = Hoyde + 1
        Case 3
            Elvis.X = Bredde + 1
            Elvis.Y = Int((Rnd * Hoyde) + 1)
        Case 4
            Elvis.Y = 0
            Elvis.X = Int((Rnd * Bredde) + 1)
        End Select
        
        'GET A RANDOM EXIT POINT
        Select Case b
        Case 1
            Elvis.TX = Bredde + 1
            Elvis.TY = Int((Rnd * Hoyde) + 1)
        Case 2
            Elvis.TX = Int((Rnd * Bredde) + 1)
            Elvis.TY = 0
        Case 3
            Elvis.TX = 0
            Elvis.TY = Int((Rnd * Hoyde) + 1)
        Case 4
            Elvis.TY = Hoyde + 1
            Elvis.TX = Int((Rnd * Bredde) + 1)
        End Select
        PlaySound "elvis2"
        Elvis.Tag = 1
    End If
End Sub
Public Sub MoveTheKing()
    If Elvis.Act Then
        If Elvis.X = Elvis.TX And Elvis.Y = Elvis.TY Then
            Elvis.Act = False 'Deactivate the king
            PlaySound "elvis"
        End If
        If Rnd > 0.94 Then PlaySound "elvis2"
        If Int(Rnd * 2) = 1 Then 'It's random if it'll walk up/down or right/left
            If Elvis.X = Elvis.TX Then GoTo Otherone1
Otherone2:
            If Elvis.X > Elvis.TX Then 'LEFT
                Elvis.X = Elvis.X - 1
                Elvis.Dire = 1
            Else
                Elvis.X = Elvis.X + 1 'RIGHT
                Elvis.Dire = 3
            End If
        Else
            If Elvis.Y = Elvis.TY Then GoTo Otherone2
Otherone1:
            If Elvis.Y > Elvis.TY Then
                Elvis.Y = Elvis.Y - 1 'UP
                Elvis.Dire = 4
            Else
                Elvis.Y = Elvis.Y + 1 'DOWN
                Elvis.Dire = 2
            End If
        End If
    End If
End Sub

Public Sub MakePalmTrees()
    N = Int((Rnd * 8) + 27) 'Get number of palms
    ReDim Palms(1 To N)
    For A = 1 To N
GetOther:
        X = Int((Rnd * Bredde) + 1)
        Y = Int((Rnd * Hoyde) + 1)
        For Y1 = Y - 2 To Y + 2
        For X1 = X - 2 To X + 2
            If Board(X1, Y1) Then GoTo GetOther 'Check if the area is clear
        Next X1
        Next Y1
        If X + 6 > BellX And X - 6 < BellX Then
        If Y + 6 > BellY And Y - 6 < BellY Then
            GoTo GetOther 'too close to the the bell
        End If
        End If
    
        'All is well =)
        Board(X, Y) = True 'Mark the Board
        
        Palms(A).Tree = Int((Rnd * 4) + 1) 'Let me think, what kind of seed?
        Palms(A).X = X
        Palms(A).Y = Y
    Next A
End Sub
Public Sub MakeForest()
GetNew:
    
    If Rnd > 0.6 Then Exit Sub
    
    X = Int((Rnd * (Bredde - 10)) + 1)
    Y = Int((Rnd * (Hoyde - 10)) + 1)
    For Y1 = Y - 2 To Y + 15
    For X1 = X - 2 To X + 15
        If Board(X1, Y1) Then GoTo GetNew 'Check if the area is clear
    Next X1
    Next Y1
    
    Forest.X = X
    Forest.Y = Y
    For A = 1 To 14
    For b = 1 To 14
        Board(X + b, Y + A) = True
    Next b
    Next A
    For A = 0 To 7
        For b = A To 7 - A
            Board(X + A, Y + 7 + b) = False
        Next b
    Next A
    For A = 0 To 2
        For b = 0 To 7
            Board(X + A + 13, Y + b) = False
        Next b
    Next A
    
    Forest.Act = True
End Sub

Public Sub DoPreGraphics()

    'The Backgound
    BitBlt Pics.PicBuffer(0).hDC, 0, 0, Bredde * size, Hoyde * size, Pics.PicBuffer(1).hDC, 0, 0, SRCCOPY
    'The skeletons
    For A = 1 To UBound(Skel)
        If Skel(A).Act Then
            Skel(A).Tag2 = Skel(A).Tag2 - 1
            If Skel(A).Tag2 = 0 Then Skel(A).Act = False
            BitBlt Pics.PicBuffer(0).hDC, (Skel(A).X * size) - size, (Skel(A).Y * size) - size, 30, 30, Pics.PicSkelM.hDC, 0, 0, SRCAND
            BitBlt Pics.PicBuffer(0).hDC, (Skel(A).X * size) - size, (Skel(A).Y * size) - size, 30, 30, Pics.PicSkel.hDC, 0, 0, SRCPAINT
        End If
    Next A
    'The Trap
    If Trap.Act Then
        A = IIf(M1.Trapped <= 0, 0, 1) 'Show trap one or two?
        BitBlt Pics.PicBuffer(0).hDC, (Trap.X * size) - size - 7, (Trap.Y * size) - size - 7, 30, 30, Pics.PicTrapM(A).hDC, 0, 0, SRCAND
        BitBlt Pics.PicBuffer(0).hDC, (Trap.X * size) - size - 7, (Trap.Y * size) - size - 7, 30, 30, Pics.PicTrap(A).hDC, 0, 0, SRCPAINT
    End If

End Sub

Public Sub Dolist()
    Main.lstMen.Clear
    If MenActive = 0 Then Exit Sub
    ReDim ListMen(1 To MenActive)
    For A = 1 To UBound(Men)
        If Men(A).Act Then
            Main.lstMen.AddItem Men(A).MyName & "   " & Men(A).Age
            I = I + 1
            ListMen(I) = A
        End If
    Next A
    ListSel = -1 'We select none
    For A = 1 To UBound(ListMen) 'Then we find the keeping person on the list
        If ListMen(A) = Keeping.Num Then
            ListSel = A - 1
        End If
    Next A
    Main.lstMen.ListIndex = ListSel
End Sub
