Attribute VB_Name = "modInfo"
Public Type Tile
 Num As Integer
 T As Integer
End Type

Public Keeping As Tile
Public IM() As Tile

Public Sub DoInformaition()
    ReDim IM(-10 To Bredde + 10, -10 To Hoyde + 10)
    'The Men
    For A = 1 To UBound(Men)
        If Men(A).Act And Not Men(A).Indoors Then
            IM(Men(A).X, Men(A).Y).Num = A
            IM(Men(A).X, Men(A).Y).T = 1
        End If
    Next A
    'The Caves
    For A = 1 To UBound(Caves)
        IM(Caves(A).X, Caves(A).Y).Num = A
        IM(Caves(A).X, Caves(A).Y).T = 2
    Next A
    'The MobbaWobba
    If M1.Act Then
        IM(M1.X, M1.Y).Num = 1
        IM(M1.X, M1.Y).T = 3
    End If
    
    Main.lblMapInfo.Caption = ""
    Select Case Keeping.T
    Case 1
        PrintInfoDude Keeping.Num
        If Men(Keeping.Num).Act = False Then
            Keeping.Num = 0
            Keeping.T = 0
            ListSel = -1
        End If
    Case 2
        PrintInfoHut Keeping.Num
    Case 3
        PrintInfoMobba
        If M1.Act = False Then
            Keeping.Num = 0
            Keeping.T = 0
        End If
    End Select
End Sub

Public Sub ProcessClick(inX, inY)
Dim X As Integer, Y As Integer
    X = Int(inX / size) + 1
    Y = Int(inY / size) + 1
    
    If X < 1 Or X > Bredde Then Exit Sub
    If Y < 1 Or Y > Hoyde Then Exit Sub
    
    If IM(X, Y).T > 0 Then 'Did we strike anything?
        Keeping.Num = IM(X, Y).Num
        Keeping.T = IM(X, Y).T
    Else 'We missed, nothing here
        For Y1 = Y - 1 To Y + 1 'do a search on the closest array
        For X1 = X - 1 To X + 1
            If IM(X1, Y1).T > 0 Then
                Keeping.T = IM(X1, Y1).T
                Keeping.Num = IM(X1, Y1).Num
                GoTo ItsOk
            End If
        Next X1
        Next Y1
ItsOk:
    End If
    If Keeping.T = 1 Then
        For A = 1 To UBound(ListMen)
            If ListMen(A) = Keeping.Num Then
                ListSel = A - 1
                Main.lstMen.ListIndex = ListSel
            End If
        Next A
    Else
        ListSel = -1
        Main.lstMen.ListIndex = -1
    End If
End Sub

Sub PrintInfoDude(M)
Dim Txt As String
Dim Sex1 As String, Sex2 As String, Sex3 As String
Dim opSex1 As String, opSex2 As String, opSex3 As String

    With Men(M)
    Txt = vbNewLine & "This is " & .MyName & "," & vbNewLine
    Sex1 = IIf(.Gender = 2, "He's", "She's")
    Sex2 = IIf(.Gender = 2, "His", "Here")
    Sex3 = IIf(.Gender = 2, "Boy", "Girl")
    opSex1 = IIf(.Gender = 1, "He's", "She's")
    opSex2 = IIf(.Gender = 1, "His", "Here")
    opSex3 = IIf(.Gender = 1, "Boy", "Girl")
    
    Txt = Txt & Sex1
    Select Case .Reason
    Case 0
        Txt = Txt & " at home."
    Case 1
        If .Return Then
            If .Indoors Then
                Txt = Txt & " inside at some friends."
            Else
                Txt = Txt & " on " & LCase(Sex2) & " way home from some friends."
            End If
        Else
            Txt = Txt & " going to visit some friends in hut " & .TargetCave & "."
        End If
    Case 2
        If .Return Then
            Txt = Txt & " going home after a walk."
        Else
            Txt = Txt & " out for a walk."
        End If
    Case 3
        If .Return Then
            If .Indoors Then
                Txt = Txt & " in the store for food."
            Else
                If .Tag > 0 Then
                    Txt = Txt & " on " & LCase(Sex2) & " way home with groceries."
                Else
                    Txt = Txt & " on " & LCase(Sex2) & " way home. " & Sex1 & " didn't get any groceries."
                End If
            End If
        Else
            Txt = Txt & " on " & LCase(Sex2) & " way to the store."
        End If
    Case 4
        If .Return Then
            Txt = Txt & " going home. " & Sex1 & " didn't find that special " & LCase(opSex3) & " this time."
        Else
            Txt = Txt & " out looking for a date."
        End If
    Case 5
        If .Return Then
            If .Indoors Then
                Txt = Txt & " err, well, go figure."
            Else
                Txt = Txt & " happily returing home."
            End If
        Else
            Txt = Txt & " has found a " & LCase(opSex3) & ". " & opSex2 & " name is " & Men(.Tag).MyName & "."
        End If
    End Select
    Txt = Txt & vbNewLine & Men(Keeping.Num).MyName & " is " & Men(Keeping.Num).Age & IIf(Men(Keeping.Num).Age = 1, " year", " years") & " old."
    If .Gender = 1 And .Pregnant > 0 Then Txt = Txt & vbNewLine & "She's also pregnant!"
    
    Main.lblMapInfo.Caption = Txt
    End With
End Sub

Sub PrintInfoHut(H)
Dim Txt As String
Dim temp As String
Dim I As Integer
Dim RetStr As String
    With Caves(H)
    Txt = vbNewLine & "This is hut " & H & "." & vbNewLine
    Txt = Txt & "People who live here: "
    For A = 1 To UBound(.LiveHere)
        If .LiveHere(A) > 0 Then
            temp = temp & Men(.LiveHere(A)).MyName & ", " 'get all the people that live here
        End If
    Next
    'we've got all the people listed like this "Al, Jill, Bob, "
    'we want it like this "Al, Jill and Bob"
    If Len(temp) > 0 Then temp = Mid(temp, 1, Len(temp) - 2) 'Remove the last two: ", "
    Do Until RetStr = "," Or I = Len(temp) 'Find the lenght of the last name
        RetStr = Mid(temp, Len(temp) - I, 1)
        I = I + 1
    Loop
    If Not I = Len(temp) Then 'If the last name is the same as the lenght it's only one name
        strFront = Mid(temp, 1, Len(temp) - I) 'if not we break it up into before the last ", "
        Strback = Mid(temp, Len(temp) - I + 3, I) ' and after it
        Txt = Txt & strFront & " and " & Strback 'Now we smack it back together with an " and " between
    Else
        Txt = Txt & temp 'only one name, put it on
    End If
    Txt = Txt & vbNewLine
    Txt = Txt & "There " & IIf(.People = 1, "is ", "are ") & LCase(NumberToLetters(.People)) & " in this hut right now."
    
    Main.lblMapInfo.Caption = Txt
    End With
End Sub
Sub PrintInfoMobba()
Dim Txt As String
    Txt = vbNewLine & "AH! This is the MobbaWobba!" & vbNewLine
    If M1.Trapped > 0 Then
        Txt = Txt & "Ha ha! He's in the trap!"
    ElseIf M1.Trapped = -1 Then
        Txt = Txt & "He's returning to the jungle, he has learned a lesson this time."
    Else
        If M1.Hunted > 0 Then
            If M1.Hunted = 101 Then
                If M1.LeaveTime > 0 Then
                    Txt = Txt & "Oh God! He's eating you!!"
                Else
                    Txt = Txt & "Quick! He's comming after you!"
                End If
            Else
                If M1.LeaveTime > 0 Then
                    Txt = Txt & "Oh God! He's eating a villager!!!"
                Else
                    Txt = Txt & "He's running after poor " & Men(M1.Hunted).MyName & "!"
                End If
            End If
        Else
            If M1.Tag > -5 Then
                Txt = Txt & "Hush. He's looking for prey."
            Else
                Txt = Txt & "He's going home! YEAHY!"
            End If
        End If
    End If
    Main.lblMapInfo.Caption = Txt
End Sub

Function NumberToLetters(Num)
Dim temp As Variant
    If Num > 10 Then Num = 10
    temp = Array("Noone", "One person", "Two people", "Three people", "Four people", "Five people", "Six people", "Seven people", "Eight people", "Nine people", "Ten people")
    NumberToLetters = temp(Num)
End Function

