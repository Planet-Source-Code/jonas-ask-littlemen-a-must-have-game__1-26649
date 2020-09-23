Attribute VB_Name = "Console"
Public ConDone(1) As Boolean

Public Sub ShowConsole()
    MainPause = True 'Pause the game
    Main.cmdPause.Enabled = False 'Diable the unpause
    Main.txtConsole.Visible = True 'Show the console
    Main.txtConsole.Top = 0 'put it at the top
    Main.txtConsole.SetFocus 'Give it focus
    Main.txtConsole.Text = "" 'blanks it
End Sub
Public Sub ConsoleInput(T)
    Select Case LCase(T) 'These are cheatcodes! Stop reading! NOW!
                '^^^^^^ They are converted to Lower Case
    Case "wolla"
        Mybox "Wolla fætter, fed BMW", 2, 0
    Case "the king"
        Mybox "Yes, we all like Elvis, don't we?", 2, 0
    Case "donkey-screwdriver-firebird"
        If ConDone(0) Then GoTo Nope 'Only once
        ConDone(0) = True
        PlaySound "ring"
        Mybox "Hello!" & vbNewLine & "It's me, ur ol' uncle Carlo!" & vbNewLine _
            & "Listen, the cops are after me, take this," & vbNewLine _
            & "and don't tell Aunty! Gotta go now!", 2, 1
        Income 10000, Store.X, Store.Y
    Case "gimme that banana"
        Mybox "No, NO! It's mine, MINE I TELL YA!!", 2, 0
    Case "its a good day to die"
        Mybox "Think so?", 2, 0
    Case "there is no cow level"
        Mybox "No, there isn't.", 2, 0
    Case "god 1"
        Mybox "Woo! God mode's on! Oh oh oh! Goody! You did it! not..", 2, 0
    Case "god"
        PlaySound "ring"
        Mybox "Yes, my child?", 2, 1
    Case "whos your daddy?"
        Mybox "Well, it's not you...", 2, 0
    Case "whos my daddy?"
        Mybox "How the hell should I know? Try asking your mother...", 2, 0
    Case "dnkroz"
        Mybox "Hey! I love Duke Nukem too! =D", 2, 0
    Case "show me the money"
        Mybox "Okay, but show me the stuff first!", 2, 0
    Case "show the stuff"
        Mybox "*sniff sniff* It's fake mon...", 2, 0
    Case "summon the bushman"
        If M1.Act Then GoTo Nope 'He's in there
        MonCountdown = 50
    Case "the simpsons is on"
        RingTownBell
    Case "ja vi elsker"
        If ConDone(1) Then GoTo Nope 'Only once
        ConDone(1) = True
        PlaySound "ring"
        Mybox "Hello? This is the Prime Minsiter of Norway." & vbNewLine _
            & "We're trying to spend all our oil-money by giving them away to the (much) less fortunate." & vbNewLine _
            & "Please accept this friendly donation of 300 000 Gwookees to your village.", 2, 1
        Income 300000, Store.X, Store.Y
    Case "noclip"
        Mybox "Did you honestly expect that to do anything?", 2, 0
    Case "who do you call?"
        Mybox "The Ghostbusters? I'd call them.", 2, 0
    Case "im a nut"
        Mybox "Cool.", 2, 0
    Case "hi"
        Mybox "Yo!", 2, 0
    Case "yo"
        Mybox "Hi!", 2, 0
    Case "whats your name?"
        Mybox "It's Jonas", 2, 0
    Case "credits"
        Mybox "This Game was made by Jonas Ask, " & vbNewLine _
        & "with idea and graphical support from Magne Olsen." & vbNewLine _
        & "We hope you enjoy the game, and found it somewhat humorous :)" & vbNewLine _
        & "- The Crew", 2, 0
    Case "the legend lives on"
        PlaySound "elvis2"
        P1.GOD = True
    Case "live and let die"
        PlaySound "click"
        P1.GOD = False
    Case "debug"
        Main.Label2.Visible = Not Main.Label2.Visible
        Main.Label4.Visible = Not Main.Label4.Visible
    Case "coconut-twisted skid mark"
        Mybox "Ok, can do!", 2, 0
        For A = 1 To UBound(Men)
            If Men(A).Act Then Men(A).Age = Int((Rnd * 5) + 15)
        Next A
        Dolist
    Case "winnie the pooh"
        Mybox "ohooo, didn't know you were into THAT sort of things!", 2, 0
    Case "quit", "exit"
        End
    Case "man, i wish i had some spaghetti"
        Mybox "We'll see what we can do.", 2, 0
        For A = 1 To UBound(Men)
            If Men(A).Act And Men(A).Gender = 2 Then Men(A).MyName = "Jonas"
        Next A
        Dolist
    Case "glittering prices"
        Mybox "What was the heavyest job in the stoneage?" & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "Awnser: yobrepaP", 2, 0
    Case "socrates liked his toast well done"
        Mybox "He did, didn't he?", 2, 0
        For A = 1 To UBound(Men)
            If Men(A).Act Then Men(A).Age = Int((Rnd * 20) + 50)
        Next A
        Dolist
    Case "who made this?"
        frmAuthor.Show , Main
    End Select
    
    'hehe, I enjoyed this =)
    
Nope:
    Main.txtConsole.Text = ""
    Main.txtConsole.Visible = False
    MainPause = False
    Main.cmdPause.Enabled = True
End Sub
