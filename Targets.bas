Attribute VB_Name = "Gotos"
Public Function GotoWalk(A) As Boolean
    Men(A).Reason = 2
    Men(A).Return = False
    Men(A).Tag = Int((Rnd * 6) + 1)
    Men(A).TX = Int((Rnd * Bredde) + 1)
    Men(A).TY = Int((Rnd * Hoyde) + 1)
    Men(A).Indoors = False
    
    GotoWalk = True
End Function

Public Function GotoVisit(m) As Boolean
GimmeOneMore:
    A = Int((Rnd * NumCaves) + 1) 'get a random cave
    Tested = Tested + 1
    If A = Men(m).HomeCave Then  'Let's not go home
        If Tested = NumCaves + 5 Then 'we've tryed enough times...
            GotoVisit = False 'We haven't got anywhere ti go
            Exit Function
        Else
            GoTo GimmeOneMore
        End If
    End If
    
    Men(m).Reason = 1 ' Visit somebody
    Men(m).TargetCave = A
    Men(m).TX = Caves(A).X
    Men(m).TY = Caves(A).Y
    Men(m).Indoors = False
    Men(m).Return = False
    
    GotoVisit = True
End Function

Public Function GotoLove(m) As Boolean
    If HousesAreFull Then Exit Function
    Men(m).Reason = 4 'Looking for Love
    Men(m).Return = False
    Men(m).Tag = Int((Rnd * 6) + 3)
    Men(m).TX = Int((Rnd * Bredde) + 1)
    Men(m).TY = Int((Rnd * Hoyde) + 1)
    Men(m).Indoors = False
    
    GotoLove = True
End Function

Function HousesAreFull() As Boolean
Dim C As Integer
    HousesAreFull = False
    'count childern
    For A = 1 To UBound(Men)
        If Men(A).Act Then
            If Men(A).Pregnant > 0 Then C = C + 1
        End If
    Next A
    
    If MenActive + C >= UBound(Men) Then HousesAreFull = True
    If MenActive + C >= (NumCaves * MaxInCave) Then HousesAreFull = True
End Function

Public Function ManGoShoping(m)
            Men(m).Reason = 3 'Gone Shoping
            Men(m).Return = False
            Men(m).Tag = 0
            Men(m).TX = Store.X
            Men(m).TY = Store.Y
            Men(m).Indoors = False
            Men(m).LeaveTime = 0
            Men(m).ThisCave = 0
            Men(m).TargetCave = 0
            LeaveCave Men(m).HomeCave
End Function
