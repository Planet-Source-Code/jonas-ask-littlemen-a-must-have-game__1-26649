Attribute VB_Name = "Monster"
Public Type aMonster
 X As Integer
 Y As Integer
 TX As Integer
 TY As Integer
 Act As Boolean
 Tag As Integer
 Tag2 As Integer
 Hunted As Integer
 LeaveTime As Integer
 Trapped As Integer
 Dire As Byte
End Type

Public Type Misc
 X As Integer
 Y As Integer
 Tag2 As Integer
 Act As Boolean
End Type

Public Trap As Misc
Public Forest As Misc
Public Skel(1 To 10) As Misc

Public M1 As aMonster
Public MonCountdown As Integer
Public MonWarned As Boolean

Public Sub MakeMonster()
    M1.Act = True
    Randomize
    Select Case Int((Rnd * 4) + 1)
    Case 1
        M1.X = 1
        M1.Y = Int((Rnd * Hoyde) + 1)
    Case 2
        M1.X = Int((Rnd * Bredde) + 1)
        M1.Y = Hoyde
    Case 3
        M1.X = Bredde
        M1.Y = Int((Rnd * Hoyde) + 1)
    Case 4
        M1.Y = 1
        M1.X = Int((Rnd * Bredde) + 1)
    End Select
    
    M1.Tag = Int((Rnd * 8) + 14) 'Give a number of things to do before he leaves
    If TownBell = 0 Then M1.Tag = M1.Tag + (MenActive / 9) 'If many people are outside (bell not rung), then add some more thigs to do.
    M1.Hunted = GetAHunted 'get a target
    
    M1.Dire = 1
    
    'Emit a sound
    PlaySound "Mobba" & Int(Rnd * 3) + 1
End Sub
Private Function ChoiseOK(M) As Boolean
    ChoiseOK = True
    If Men(M).Act = False Then ChoiseOK = False 'He's not alive
    If Men(M).Indoors = True Then ChoiseOK = False 'Can't see him, he's inside
    
    'This guy is ok to eat
End Function

Public Sub MonsterMove()
    If M1.Act Then
        If M1.X = Trap.X And M1.Y = Trap.Y And M1.Trapped = 0 And Trap.Act Then
            M1.Trapped = 50 'Capture the monster for 50 ticks
            PlaySound "trap"
            Exit Sub
        End If
        If M1.Trapped = 1 Then 'He's been in the trap for long enough, let him go home
            MonRelease
        End If
        
        If M1.Trapped > 0 Then 'He's in the trap!
            M1.Trapped = M1.Trapped - 1
            M1.Dire = Int((Rnd * 4) + 1)
            If Rnd > 0.9 Then PlaySound "mobbaidle" & Int(Rnd * 4) + 1 'play some sounds
            Exit Sub
        End If
        
        
        If M1.LeaveTime > 1 Then 'Is he still eating?
            M1.LeaveTime = M1.LeaveTime - 1
            Exit Sub
        ElseIf M1.LeaveTime = 1 Then
            M1.LeaveTime = M1.LeaveTime - 1
            If M1.Tag > 0 Then M1.Hunted = GetAHunted
        End If
        
        MonCheckTarget
        
        If M1.Tag > 0 Then
            If M1.Hunted > 0 Then
                If M1.Hunted = 101 Then 'he's hunting us! Good idea: LET'S TELL HIM WHERE WE ARE!
                    M1.TX = P1.X
                    M1.TY = P1.Y
                Else
                    M1.TX = Men(M1.Hunted).X
                    M1.TY = Men(M1.Hunted).Y
                End If
            Else
                M1.Hunted = GetAHunted
            End If
        End If
        

        If M1.Hunted = 101 Then
            If P1.Outside = False Then
                M1.Tag2 = 0
                M1.Hunted = GetAHunted
            End If
        ElseIf M1.Hunted > 0 Then
            If Men(M1.Hunted).Indoors Then
                M1.Tag2 = 0
                M1.Hunted = GetAHunted
            End If
        End If
        
        
        If M1.Tag = 0 Or M1.Tag = -1 Or M1.Tag = -2 Then 'He's done all he can
            M1.Tag = -10
            M1.Hunted = 0
            Randomize 'GET A RANDOM EXIT POINT
            Select Case Int((Rnd * 4) + 1)
            Case 1
                M1.TX = 0
                M1.TY = Int((Rnd * Hoyde) + 1)
            Case 2
                M1.TX = Int((Rnd * Bredde) + 1)
                M1.TY = Hoyde + 1
            Case 3
                M1.TX = Bredde + 1
                M1.TY = Int((Rnd * Hoyde) + 1)
            Case 4
                M1.TY = 0
                M1.TX = Int((Rnd * Bredde) + 1)
            End Select
        End If
        
        If Rnd > 0.98 Then PlaySound "mobbaidle" & Int(Rnd * 4) + 1
        
        Dim TempX As Integer, TempY As Integer
        TempX = M1.X
        TempY = M1.Y
        
        If Int(Rnd * 2) = 1 Then 'It's random if it'll walk up/down or right/left
            If M1.X = M1.TX Then GoTo Otherone1
Otherone2:
            If M1.X > M1.TX Then 'LEFT
                TempX = TempX - 1
                M1.Dire = 1
            Else
                TempX = TempX + 1 'RIGHT
                M1.Dire = 3
            End If
        Else
            If M1.Y = M1.TY Then GoTo Otherone2
Otherone1:
            If M1.Y > M1.TY Then
                TempY = TempY - 1 'UP
                M1.Dire = 4
            Else
                TempY = TempY + 1 'DOWN
                M1.Dire = 2
            End If
        End If

        M1.X = TempX 'it was ok, give coords
        M1.Y = TempY
        
        'Will he jump his target?
        If M1.Hunted > 0 And M1.Hunted < 101 Then 'let's not take the player
            If M1.X + 1 >= Men(M1.Hunted).X And M1.X - 1 <= Men(M1.Hunted).X Then
            If M1.Y + 1 >= Men(M1.Hunted).Y And M1.Y - 1 <= Men(M1.Hunted).Y Then
                M1.X = Men(M1.Hunted).X 'the Jump
                M1.Y = Men(M1.Hunted).Y
            End If
            End If
        End If
    End If

End Sub

Private Function GetAHunted() As Integer
    Dim Tried(1 To 100) As Boolean
    Dim Tell As Byte
    
    If M1.Tag2 > 0 Then 'This is to prevent the monster form looking for guys ALL the time (saves resources)
        M1.Tag2 = M1.Tag2 - 1 'Time till look for a guy is one less
        Exit Function 'Keep target
    Else
        M1.Tag = M1.Tag - 1
        M1.Tag2 = Int((Rnd * 10) + 7) 'Give a random time till next check, 7 to 17 ticks
    End If
    
    If Rnd > 0.8 And P1.Outside Then
        GetAHunted = 101
        Exit Function
    End If
    
    I = Int((Rnd * UBound(Men)) + 1)
    Tried(I) = True
    
    Do Until ChoiseOK(I) Or Tell = 100
        I = Int((Rnd * UBound(Men)) + 1)
        Tried(I) = True
        Tell = 0
        For A = 1 To 100
            If Tried(A) Then Tell = Tell + 1
        Next A
    Loop

    If Tell = 100 Then
        If M1.X = M1.TX And M1.Y = M1.TY Then 'only give new target if we reached the old one
            If P1.Outside Then 'Who's that guy in the funky blue hat?
                GetAHunted = 101 'Hunt us! woooaahh! D=
                PlaySound "mobbahunt" & Int(Rnd * 4) + 1
            Else
                Walkaround
            End If
        Else
            Walkaround
            Exit Function
        End If
    Else 'He found a guy to chase!
        PlaySound "mobbahunt" & Int(Rnd * 4) + 1
        GetAHunted = I
    End If
End Function

Public Sub Walkaround()
    Randomize
    M1.TX = Int((Rnd * Bredde) + 1)
    M1.TY = Int((Rnd * Hoyde) + 1)
End Sub

Public Sub MonCheckTarget()
    If M1.X = M1.TX And M1.Y = M1.TY Then 'If reached target
        If M1.Tag > 0 Then 'If the hunt is not over
            M1.Tag2 = 0
            If M1.Hunted > 0 Then
                If M1.Hunted = 101 Then 'The guy WAS the player:
                    If P1.Outside And P1.GOD = False Then 'woops
                        M1.LeaveTime = 20 'He's eating... US!
                        EatPlayer
                    Else 'The player got inside intime
                        M1.Hunted = GetAHunted
                        M1.Tag = M1.Tag - 2
                        Exit Sub
                    End If
                Else 'The guy was not the player:
                    If Men(M1.Hunted).Indoors Or Men(M1.Hunted).Act = False Then
                        M1.Hunted = GetAHunted
                        M1.Tag = M1.Tag - 2
                        Exit Sub
                    Else
                        Men(M1.Hunted).X = M1.X: Men(M1.Hunted).Y = M1.Y 'adjust die spot
                        KillGuy M1.Hunted '*munch, chomp, crunch*
                        PlaySound "Mobbaeat" 'Make some slurpy noises
                        M1.Tag = M1.Tag - 1 'Eating is why the monster came, so it should count for TWO on the 'to do' note
                        M1.LeaveTime = Int((Rnd * 7) + 15) 'Remember to chew well before you swallow
                    End If
                End If
            End If
            'M1.Hunted = GetAHunted 'He needs a new target
        Else 'The hunt is over
            DeactivateMonster
        End If
    End If
End Sub

Public Sub KillGuy(M)
    If Men(M).Reason = 3 Then 'The man is out shoping
        Caves(Men(M).HomeCave).FoodOk = False 'Flag his homecave that noone is out for food
    End If
    With Caves(Men(M).HomeCave)
        For A = 1 To 6
            If .LiveHere(A) = M Then
                .LiveHere(A) = 0
            End If
        Next A
    End With
    If Men(M).Indoors And Men(M).Reason = 0 Then  'if they are indoors, move the body outside
        Men(M).Y = Men(M).Y + 1
        Caves(Men(M).ThisCave).People = Caves(Men(M).ThisCave).People - 1
    End If
    MakeSkel Men(M).X, Men(M).Y 'Make the skeleton
    Men(M).Act = False
    Men(M).Dire = 0
    Men(M).Age = 0
    Men(M).MyName = ""
    Men(M).Gender = 0
    Men(M).HomeCave = 0
    Men(M).Indoors = False
    Men(M).LeaveTime = 0
    Men(M).Pregnant = 0
    Men(M).Reason = 0
    Men(M).Return = False
    Men(M).Tag = 0
    Men(M).TargetCave = 0
    Men(M).ThisCave = 0
    Men(M).TX = 0
    Men(M).TY = 0
    Men(M).X = 0
    Men(M).Y = 0
    
    MenActive = MenActive - 1 'One less in our proud ranks
    Dolist
End Sub

Public Sub DeactivateMonster()
    M1.Act = False
    M1.Dire = 0
    M1.Hunted = 0
    M1.LeaveTime = 0
    M1.Tag = 0
    M1.Tag2 = 0
    M1.TX = 0
    M1.TY = 0
    M1.X = 0
    M1.Y = 0
    If M1.Trapped = 0 Then 'The monster wasn't trappet, and leard no lesson from it
        MonCountdown = Int((Rnd * 1400) + 600)
    ElseIf M1.Trapped = -1 Then 'We got him this time. It'll be a while till he's back
        MonCountdown = Int((Rnd * 500) + 2500)
    End If
    M1.Trapped = 0
    MonWarned = False
End Sub


Public Sub DoMonster()
    'WORK THE MONSTER COUNTDOWN----------
    If MonCountdown > 0 Then
        MonCountdown = MonCountdown - 1
    ElseIf MonCountdown = 0 Then
        MonCountdown = -1
        MakeMonster
        Mybox "Sound the bell, run for cover! The Mobbowabba is here!", 2, 3
    End If
    
    If MonCountdown < Int((Rnd * 40) + 50) And MonWarned = False Then
        MonWarned = True
        Mybox "Some of the neighboring villages report that they've seen the Mobbowabba heading this way!", 2, 3
    End If
End Sub
Public Sub BuyTrap()
Dim Svar As String
Dim Sex As String
    If Not P1.Money >= 2500 Then Exit Sub
    If Trap.Tag2 = 1 Then Exit Sub
    If Trap.Act = True Then Exit Sub
    If P1.Outside Then Exit Sub
    
    PlaySound "Ring"
    Svar = Mybox("Acme Do-wackies:" & vbNewLine & vbNewLine & "So, you are interested in our excelent 'Universal Monster Trap 4500 C Self Assembler (For all your monster trapping needs) kit'?" & vbNewLine & "If so, the cost is 2500 Gwookees." & vbNewLine & "Shall I place the order, Sir?", 1, 1)
    
    If Not Svar = "yes" Then Exit Sub
    
    'all is OK: Assemble the kit
    Expense -2500, Store.X, Store.Y
    PlaySound "register"
    
    Trap.Tag2 = 1
End Sub

Public Sub placeTrap()
    If Trap.Tag2 = 0 Then Exit Sub
    
    For Y1 = P1.Y - 2 To P1.Y + 2
    For X1 = P1.X - 2 To P1.X + 2
        If Board(X1, Y1) Then Exit Sub 'Check if the area is clear
    Next X1
    Next Y1
    
    'Seems ok, start setting it up
    Trap.Tag2 = 0
    Trap.Act = True
    Trap.X = P1.X
    Trap.Y = P1.Y
End Sub
Public Sub MonRelease()
    PlaySound "mobbamoan"
    
    M1.Tag2 = 0 'He'll go home now
    M1.Tag = 0
    M1.Trapped = -1 'flag him as trapped
    
    'Reset the trap
    Trap.Tag2 = 0
    Trap.Act = False
    Trap.X = 0
    Trap.Y = 0
End Sub
Public Sub MakeSkel(X, Y)
    Dim I As Integer
    I = 1 'find a free skeleton
    Do Until Skel(I).Act = False Or I = UBound(Skel) + 1
        I = I + 1
    Loop
    Skel(I).Act = True
    Skel(I).Tag2 = 100
    Skel(I).X = X
    Skel(I).Y = Y
End Sub
