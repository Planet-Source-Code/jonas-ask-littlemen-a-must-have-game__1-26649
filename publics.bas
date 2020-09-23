Attribute VB_Name = "modMen"
Public GrowTag As Byte 'This keeps time till people gorw one year


Public Sub ManInitialize(C, Mother, X, Y, G)
    If MenActive = UBound(Men) Then Exit Sub 'Simulation is full
    If MenActive >= (UBound(Caves) * MaxInCave) Then  'Simulation is full
        Exit Sub
    End If
    
    MenActive = MenActive + 1
    
    Dim I As Integer
    I = 1 'find a free guy
    Do Until Men(I).Act = False
        I = I + 1
    Loop
    
    Men(I).Act = True
    
    If C > 0 Then 'Differ between given and random start point
        'START IN CAVE
        Randomize
        If G = 0 Then
            Men(I).Gender = Int((Rnd * 2) + 1) 'Get a random sex
        Else
            Men(I).Gender = G 'Set the assigned sex
        End If
        Men(I).X = Caves(C).X
        Men(I).Y = Caves(C).Y
        Men(I).Dire = Int((Rnd * 4) + 1)
        Men(I).Indoors = True
        Men(I).HomeCave = C
        Men(I).ThisCave = C
        Men(I).Age = Int((Rnd * 15) + 10) 'Give a random age
        Men(I).LeaveTime = Int((Rnd * (WaitTimeMax - WaitTimeMin + 1)) + WaitTimeMin)
        ArriveAtCave C
        Caves(C).LiveHere(Caves(C).People) = I
    Else
        'GET BORN IN THE FIELD
        If Mother > 0 Then
            Men(I).X = Men(Mother).X
            Men(I).Y = Men(Mother).Y
            'This guy is BORN, play sound
            Men(I).Age = 0
            PlaySound "baby" & Int((Rnd * 2) + 1)
        Else
            Men(I).X = X 'This guy has a given X Y start point, but no mother...
            Men(I).Y = Y
            Men(I).Age = Int((Rnd * 10) + 16) 'Give a random age
        End If
        If G = 0 Then
            Men(I).Gender = Int((Rnd * 2) + 1) 'Get a random sex
        Else
            Men(I).Gender = G 'Set the assigned sex
        End If
        Men(I).Indoors = False
        Men(I).HomeCave = GetNewHomeCave
        MoveIntoCave Men(I).HomeCave, I
        Men(I).ThisCave = 0
        Men(I).LeaveTime = 0
        Men(I).Dire = Int((Rnd * 4) + 1)
        Men(I).Return = True
        Men(I).Reason = 2
        Men(I).Tag = 0
        Men(I).TargetCave = 0
        ManGiveTarget I
    End If
    
    Men(I).MyName = MenGetAName(Men(I).Gender)
    
    Dolist
End Sub

Public Sub ManMove(A)

    If Men(A).Pregnant >= 1 Then
        Men(A).Pregnant = Men(A).Pregnant - 1
        If Men(A).Pregnant = 0 Then ManInitialize 0, A, 0, 0, 0
    End If
    
    If Men(A).Indoors Then
        Men(A).LeaveTime = Men(A).LeaveTime - 1 'One less tick till he'll leave
        If Men(A).LeaveTime <= 0 Then 'He's leaving
            If Men(A).Reason = 3 Then Income Int(Men(A).Tag / 11), Store.X, Store.Y 'He's done shoping, pay for the groceries
            If ManGiveTarget(A) Then 'if we found him a new target...
                LeaveCave Men(A).ThisCave '... then leave this cave
                Men(A).ThisCave = 0
            Else '... else, wait some more.
                Men(A).LeaveTime = Int((Rnd * (WaitTimeMax - WaitTimeMin + 1)) + WaitTimeMin)
                Exit Sub
            End If
        Else
            Exit Sub 'She's puti'n on makup....
        End If
    End If
    
    If Men(A).Reason = 2 Then 'If we're out for a walk,
        If Rnd > 0.99 And TownBell = 0 Then    'maybe we'll pop in at some friends
            GotoVisit A
        End If
    End If
    
    
    If Men(A).Reason = 4 Then 'We'er looking for a mate here...
        ManMateHunt A
    End If
    
    If Men(A).Reason = 5 And Men(A).Return = False Then  'Are we going to mate, and have they "done it"?
        Men(A).TX = Men(Men(A).Tag).X
        Men(A).TY = Men(Men(A).Tag).Y
    End If
    
    Dim TempX As Integer, TempY As Integer
    TempX = Men(A).X
    TempY = Men(A).Y
    If Int(Rnd * 2) = 1 Then 'It's random if it'll walk up/down or right/left
        If Men(A).X = Men(A).TX Then GoTo Otherone1
Otherone2:
        If Men(A).X > Men(A).TX Then 'For LEFT we set sail!
            TempX = TempX - 1
            Men(A).Dire = 1
        Else
            TempX = TempX + 1 'RIGHT isn't too bad....
            Men(A).Dire = 3
        End If
    Else
        If Men(A).Y = Men(A).TY Then GoTo Otherone2
Otherone1:
        If Men(A).Y > Men(A).TY Then
            TempY = TempY - 1 'ghee, I have to go UP a notch
            Men(A).Dire = 4
        Else
            TempY = TempY + 1 'No, wait, DOWN. My bad
            Men(A).Dire = 2
        End If
    End If
    
    Men(A).X = TempX 'it was ok, give coords
    Men(A).Y = TempY
    
End Sub

Public Sub ManCheckTarget(A)
    If Men(A).Indoors Then Exit Sub 'People at home arn't 'a goi'n anywhere
    If Men(A).X = Men(A).TX And Men(A).Y = Men(A).TY Then
        
        If Men(A).TX = Caves(Men(A).HomeCave).X And Men(A).TY = Caves(Men(A).HomeCave).Y Then  'arrive at homecave
            ArriveAtCave Men(A).HomeCave
            Men(A).ThisCave = Men(A).HomeCave
            Men(A).Indoors = True
            Men(A).Return = False
            Men(A).LeaveTime = Int((Rnd * (WaitTimeMax - WaitTimeMin + 1)) + WaitTimeMin)
            
            'If he arrived from the store, add the food.
            If Men(A).Reason = 3 Then
                Caves(Men(A).TargetCave).Food = Caves(Men(A).TargetCave).Food + Men(A).Tag
                Caves(Men(A).TargetCave).FoodOk = False
            End If
            Men(A).Reason = 0
            Men(A).TargetCave = 0
            Exit Sub
        End If
        
        Select Case Men(A).Reason
        Case 1 'Visit
            arrHut A
        Case 2 'Walk
            arrWalk A
        Case 3 'In the Store
            arrStore A
        Case 4 'At a love-hunt turn point
            arrWalk A 'walk to another coor
        Case 5 'Found Love
            arrPartner A
        End Select
        
    End If
End Sub

Public Function ManGiveTarget(M) As Boolean
Dim Tested As Integer


    
    'This  guy is just about to go home
    If Men(M).Return = True Then
        Men(M).TX = Caves(Men(M).HomeCave).X
        Men(M).TY = Caves(Men(M).HomeCave).Y
        Men(M).Indoors = False
        Men(M).TargetCave = Men(M).HomeCave
        
        ManGiveTarget = True
        Exit Function
    End If
    
    If TownBell > 0 Then 'The town bell is rung, stay at home
        ManGiveTarget = False
        Exit Function
    End If
    
    'Not In Move, Give a Totaly New Target
WontDoIt:
        Randomize
        Select Case Int((Rnd * 10) + 1)
        Case 1 To 3 'Visit
            ManGiveTarget = GotoVisit(M)
        Case 4 To 7 'Go for a stroll
            ManGiveTarget = GotoWalk(M)
        Case 8 To 10  'Goin out looking for loooove
            If Men(M).Age < 16 Or Men(M).Age > 60 Then GoTo WontDoIt  'We do have at least some scruples in this village
            ManGiveTarget = GotoLove(M)
        End Select
        
End Function


Sub ManMateHunt(A)
        For b = 1 To UBound(Men)
            If Men(b).Act Then
                If Men(b).Indoors = False And Men(b).Age > 16 And Men(b).Age < 60 Then
                    If Men(b).Reason = 2 Or Men(b).Reason = 4 Then 'only if he/she is out walking or looking
                        If Men(A).Gender = 1 Then 'A woman
                            If Men(b).Gender = 2 Then
                                If Int((Rnd * 100) + 1) < 1.5 ^ Men(b).Reason Then
                                    Men(A).Reason = 5
                                    Men(A).Tag = b
                                    Men(b).Reason = 5
                                    Men(b).Tag = A
                                End If
                            End If
                        Else 'A man (or a the "else"-guys ;p)
                            If Men(b).Gender = 1 Then
                                If Int((Rnd * 100) + 1) < 1.5 ^ Men(b).Reason Then
                                    Men(A).Reason = 5
                                    Men(A).Tag = b
                                    Men(b).Reason = 5
                                    Men(b).Tag = A
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next b
End Sub

Public Sub BuyGuy(G)
Dim Svar As String
Dim Sex As String
    If Not P1.Money >= 3000 Then Exit Sub
    If HousesAreFull Then Exit Sub
    If P1.Outside Then Exit Sub
    
    PlaySound "Ring"
    Sex = IIf(G = 2, "Male", "Female")
    Svar = Mybox("Acme Do-wackies:" & vbNewLine & vbNewLine & "So, you would like to order one of our 'Ultra-Instant-" & Sex & "++' kits?" & vbNewLine & " If so it'll be 3000 Gwookees." & vbNewLine & "Shall I place the order, Sir?", 1, 1)
    
    If Not Svar = "yes" Then Exit Sub
    
    'all is OK: Assemble the kit
    'Randomize
    X = Int((Rnd * Bredde) + 1)
    Y = Int((Rnd * Hoyde) + 1)
    ManInitialize 0, 0, X, Y, G
        
    Expense -3000, X, Y
    PlaySound "spawn"
End Sub

Function MenGetAName(S) As String
Dim Male As Variant
Dim Female As Variant
    Male = Array("Aaron", "Abe", "Abraham", "Adrian", "Al", "Albert", "Alf", "Allan", "Andreas", "Andrew", "Angus", "Arthur", "Baldwin", "Bart", "Ben", "Bill", "Bob", "Boris", "Brian", "Burt", "Calvin", "Cato", "Charles", "Clive", "Dan", "Dave", "Dick", "Doug", "Eddie", "Elmer", "Eric", "Ewen", "Fabian", "Frank", "Gene", "Glenn", "Greg", "Haakon", "Henry", "Herb", "Hugo", "Ike", "Igor", "Ivan", "Jack", "jeff", "Jim", "Jonas", "Justin", "Kevin", "Kim", "Laban", "Leo", "Luke", "Mac", "Magne", "Manuel", "Mark", "Maxwell", "Morris", "Nathan", "Ned", "Olaf", "Owen", "Pat", "Peary", "Phil", "Ralph", "Raymond", "Richard", "Robert", "Roy", "Sam", "Scott", "Sid", "Steve", "Stan", "Ted", "Thor", "Vic", "Will", "Zacharias")
    Female = Array("Adela", "Agnes", "Alexandra", "Alice", "Amanda", "Ann", "Anna", "Barbara", "Beatrice", "Bessie", "Betty", "Brenda", "Bridget", "Camilla", "Carmen", "Caterina", "Catherine", "Cecile", "Celia", "Clare", "Claudia", "Connie", "Daisy", "Dana", "Debby", "Denise", "Dolly", "Edith", "Edna", "Elaine", "Elinor", "Ellen", "Elsa", "Emily", "Erica", "Eva", "Felicia", "Flora", "Francis", "Gill", "Grace", "Gwen", "Hannah", "Helen", "Hilda", "Ida", "Ira", "Isis", "Ivy", "Jane", "Janice", "Jennifer", "Jenny", "Jessica", "Jill", "Judy", "June", "Karen", "Kay", "Lana", "Lena", "Linda", "Lucia", "Maggie", "Mandy", "Margaret", "Maria", "Melissa", "Myra", "Natalie", "Nina", "Pamela", "Patricia", "Peggy", "Rebecca", "Rita", "Rose", "Sally", "Sandra", "Sonia", "Sylvia", "Tina", "Tracy", "Ulrica", "Veronica", "Victoeia", "Wilma", "Yvonne")
    
    Select Case S
    Case 1
        A = Int(Rnd * UBound(Female))
        MenGetAName = Female(A)
    Case 2
        A = Int(Rnd * UBound(Male))
        MenGetAName = Male(A)
    End Select
End Function
Public Sub MenGrow()
    For A = 1 To UBound(Men)
        If Men(A).Act Then Men(A).Age = Men(A).Age + 1
    Next A
    Dolist
End Sub
