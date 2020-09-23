Attribute VB_Name = "Money"
Public Type Sign
 Amount As Currency
 Color As String
 X As Integer
 Y As Integer
 Act As Boolean
 Tag As Byte
End Type
Public Signs(1 To 20) As Sign

Public Sub Income(A, X, Y)
    P1.Money = P1.Money + A
    AddSign A, X, Y
    PlaySound "click"
End Sub

Public Sub Expense(A, X, Y)
    P1.Money = P1.Money + A
    AddSign A, X, Y
End Sub

Public Sub AddSign(A, X, Y)
    If A = 0 Then Exit Sub 'nothin
    
    i = 1
    Do Until Signs(i).Act = False Or i = 20
        i = i + 1
    Loop
    Signs(i).Act = True
    Signs(i).Amount = A
    Signs(i).Color = IIf(A < 0, vbRed, vbGreen)
    Signs(i).Tag = 70
    Signs(i).X = X
    Signs(i).Y = Y - 1
End Sub

Public Sub DoSigns()
    For A = 1 To UBound(Signs)
        If Signs(A).Act Then
            Signs(A).Tag = Signs(A).Tag - 2 'move it 2 pixels up
            If Signs(A).Tag <= 0 Then 'if the time is up, reset the sign
                Signs(A).Act = False
                Signs(A).Amount = 0
                Signs(A).Color = ""
            End If
        End If
    Next A
End Sub

Public Sub BuildHut(X, Y)
Dim TempC() As aCave
    
    If P1.Money < -HutCost Then Exit Sub 'can we afford this?
    For Y1 = Y - 2 To Y + 2
    For X1 = X - 2 To X + 2
        If Board(X1, Y1) Then Exit Sub 'Check if the area is clear
    Next X1
    Next Y1
    If UBound(Caves) >= Maxcaves Then Exit Sub 'Too Many caves
    If X = 1 Or Y = 1 Then Exit Sub 'too close to the upper or left edge
    
    'Start Building
    Board(X, Y) = True 'Mark this tile as buildt
    
    Expense HutCost, X, Y 'Pay for it
    
    ReDim TempC(1 To UBound(Caves)) 'make a backup of the exciting huts
    For A = 1 To UBound(Caves)
        TempC(A).Food = Caves(A).Food
        TempC(A).FoodOk = Caves(A).FoodOk
        For b = 1 To 6
            TempC(A).LiveHere(b) = Caves(A).LiveHere(b)
        Next b
        TempC(A).People = Caves(A).People
        TempC(A).Store = Caves(A).Store
        TempC(A).X = Caves(A).X
        TempC(A).Y = Caves(A).Y
    Next A
    
    ReDim Caves(1 To UBound(TempC) + 1) 'then redim them one higer
    
    For A = 1 To UBound(TempC) 'Now put stuff back in place
        Caves(A).Food = TempC(A).Food
        Caves(A).FoodOk = TempC(A).FoodOk
        For b = 1 To 6
            Caves(A).LiveHere(b) = TempC(A).LiveHere(b)
        Next b
        Caves(A).People = TempC(A).People
        Caves(A).Store = TempC(A).Store
        Caves(A).X = TempC(A).X
        Caves(A).Y = TempC(A).Y
    Next A
    
    Dim Num As Integer
    Num = UBound(Caves) 'Fill in the last cave
    Caves(Num).Food = 0
    Caves(Num).FoodOk = False
    Caves(Num).People = 0
    Caves(Num).Store = False
    Caves(Num).X = X
    Caves(Num).Y = Y
    
    NumCaves = NumCaves + 1 'Keep count
    
    PlaySound "register"
End Sub


