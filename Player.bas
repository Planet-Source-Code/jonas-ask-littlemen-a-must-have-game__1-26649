Attribute VB_Name = "Player"
Public Type Player
 Money As Long
 Outside As Boolean
 X As Integer
 Dire As Byte
 Y As Integer
 GOD As Boolean
End Type


Public P1 As Player
Public Elvis As aMonster

Public BellX As Integer
Public BellY As Integer
Public BellTag As Integer
Public myBoxSvar As String

Public Sub LeaveStore()
    If P1.Outside Then Exit Sub 'cant go out out
    P1.Outside = True 'He must have gone out
    P1.Dire = 2 'Face down
    P1.Y = P1.Y + 1 'Take one step
    P1.X = Store.X
    StoreOpen = False 'Out for, well, something
    PlaySound "lock" 'click
End Sub

Public Sub EnterStore()
    P1.Outside = False
    StoreOpen = True
End Sub

Public Sub DoKeys()
    If GetAsyncKeyState(220) <> 0 Then
        ShowConsole
    End If
    If P1.Outside = False Then Exit Sub 'Nothing will happen if we're inside
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        P1.X = P1.X - 1
        P1.Dire = 1
    End If
    If GetAsyncKeyState(vbKeyDown) <> 0 Then
        P1.Y = P1.Y + 1
        P1.Dire = 2
    End If
    If GetAsyncKeyState(vbKeyRight) <> 0 Then
        P1.X = P1.X + 1
        P1.Dire = 3
    End If
    If GetAsyncKeyState(vbKeyUp) <> 0 Then
        P1.Y = P1.Y - 1
        P1.Dire = 4
    End If

    If P1.X <= 1 Then P1.X = 1 'Prevent us from going of the edge of the world
    If P1.Y <= 1 Then P1.Y = 1
    If P1.X >= Bredde Then P1.X = Bredde
    If P1.Y >= Hoyde Then P1.Y = Hoyde
    
    If P1.X = Store.X And P1.Y = Store.Y Then EnterStore 'arrived at the store
    If P1.X = BellX And P1.Y = BellY Then RingTownBell 'arrived at the Bell
    
End Sub

Public Sub RingTownBell()
    If TownBell > 100 Then Exit Sub 'the bell has been rung too recently.
    TownBell = 200
    For A = 1 To UBound(Men)
        If Men(A).Act Then
            With Men(A)
            
            If .ThisCave = .HomeCave Then 'at home
                GoTo ok
            End If
                        
            Select Case .Reason
            Case 1 'On a visit
                If .Indoors = True Then
                    .LeaveTime = 0
                Else
                    .Return = True
                    GoHome A
                End If
            Case 2 'out for a walk
                .Return = True
                GoHome A
            Case 3 'Out shoping
                If .Indoors Then
                    .LeaveTime = 0
                Else
                    .Return = True
                    GoHome A
                End If
            Case 4 'Looking for a date
                .Return = True
                GoHome A
            Case 5 'found partner
                If .Indoors Then
                    .LeaveTime = 0
                Else
                    .Return = True
                    GoHome A
                End If
            End Select
ok:
            End With
        End If
    Next A
End Sub

Public Sub PutUpBell()
Dim X As Integer
Dim Y As Integer
GetNew:
    X = Int((Rnd * Bredde) + 1)
    Y = Int((Rnd * Hoyde) + 1)
    
    If X + 16 > Store.X And X - 16 < Store.X Then
    If Y + 16 > Store.Y And Y - 16 < Store.Y Then
        GoTo GetNew 'too close to the store
    End If
    End If
    
    If X < 3 Or X > Bredde - 3 Then GoTo GetNew
    If Y < 3 Or Y > Hoyde - 3 Then GoTo GetNew 'too close to the edge
    
    For Y2 = -2 To 2 'Prevent the bell form apearing over other things
    For X2 = -2 To 2
        If Board(X + X2, Y + Y2) Then GoTo GetNew
    Next X2
    Next Y2
    
    Board(X, Y) = True 'Mark the map at this point
    
    BellTag = 1
    BellX = X
    BellY = Y
End Sub
Public Sub DoBell()
    'WORK THE BELL---------
    If TownBell > 0 Then
        TownBell = TownBell - 1
        If BellTag = 0 Then 'Cycle the Bell Animation
            BellTag = 3
        Else
            BellTag = BellTag - 1
        End If
    Else
        BellTag = 1 'Bell not ringing, put it in Down position
    End If
End Sub

    
Public Sub EatPlayer()
    PlaySound "mobbaburp"
    Mybox "Oh, no! You were eaten by the Mobbowabba!" & vbNewLine & "Tough luck ;) But thanks for playing the game! ", 2, 4
    End
End Sub

Public Function Mybox(Text, Style, Picture)
    If frmBox.Visible Then Exit Function 'The box is in use
    With frmBox
    .picIcon.PaintPicture .picImages(Picture), 0, 0, 100, 100, 0, 0, 100, 100
    Select Case Style
    Case 1 'YES/NO
        .cmdNo.Visible = True
        .cmdYes.Visible = True
        .cmdOk.Visible = False
    Case 2 'OK only
        .cmdNo.Visible = False
        .cmdYes.Visible = False
        .cmdOk.Visible = True
    End Select
    .lblinfo = Text 'Put in the text
    .Show , Main 'now show it
    myBoxSvar = ""
    
    End With
    Do Until myBoxSvar <> "" 'Loop until the player awnsers the msgbox
        DoEvents
    Loop
    Mybox = myBoxSvar
    myBoxSvar = ""
End Function
