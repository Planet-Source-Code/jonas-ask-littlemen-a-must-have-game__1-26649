Attribute VB_Name = "Arrive"
Public Sub arrHut(A)
    'THIS IS WHERE A HUT IS ENTERED
    If Caves(Men(A).TargetCave).People = 0 Or Caves(Men(A).TargetCave).People >= MaxInCave Then
        'Noone was home, or the hut was full, return home
        If Rnd > 0.9 Then 'a chance we'll go for a walk
            Men(A).Reason = 2
            Men(A).Tag = Int((Rnd * 3) + 1)
            Men(A).TX = Int((Rnd * Bredde) + 1)
            Men(A).TY = Int((Rnd * Hoyde) + 1)
            Men(A).TargetCave = 0
        Else
            Men(A).Return = True
            ManGiveTarget A
        End If
    Else
        'Somebody were home, go inside
        Men(A).ThisCave = Men(A).TargetCave
        Men(A).Indoors = True
        Men(A).Return = True
        Men(A).LeaveTime = Int((Rnd * (WaitTimeMax - WaitTimeMin + 1)) + WaitTimeMin)
        ArriveAtCave Men(A).TargetCave
    End If
End Sub

Public Sub arrWalk(A)
    'THIS IS WHEN A STROLLER DECIDES TO TURN
    Men(A).Tag = Men(A).Tag - 1
    If Men(A).Tag = 0 Then
        Men(A).Return = True
        ManGiveTarget A
    Else
        Randomize
        Men(A).TX = Int((Rnd * Bredde) + 1)
        Men(A).TY = Int((Rnd * Hoyde) + 1)
    End If
End Sub

Public Sub arrStore(A)
    'THIS IS WHEN SOMEONE ARRIVES AT THE STORE
    If StoreOpen Then
        Men(A).Tag = Int((Rnd * 200) + 800) + (Caves(Men(A).HomeCave).People ^ 4) 'The more people, the more food we buy
        Men(A).LeaveTime = Int((Rnd * 50) + 30)
        Men(A).Indoors = True
        Men(A).Return = True
        Men(A).TargetCave = 0
    Else 'The store was closed
        Men(A).LeaveTime = 0
        Men(A).Return = True
        ManGiveTarget A
    End If
End Sub

Public Sub arrPartner(A)
    Men(A).LeaveTime = 20
    Men(A).Indoors = True 'Well, theyre not exactly indoors, but it prevents them from moving
    Men(A).Return = True
    If Men(A).Gender = 1 Then 'if the woman, then make her pregnant
        Men(A).Pregnant = 90
    End If
End Sub
