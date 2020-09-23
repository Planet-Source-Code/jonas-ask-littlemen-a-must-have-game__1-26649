Attribute VB_Name = "Sound"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10



Public Function PlaySound(File As String)
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Svar = sndPlaySound(App.Path & "\sound\" & File & ".wav", wFlags%) 'Send the sound to the big world
End Function

