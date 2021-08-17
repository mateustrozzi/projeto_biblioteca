msgbox "Horário atualizado com a internet!" & chr(13) & chr(13) & _
 "Todas vocês são muuuuitooo bocós!",,"Mateus diz:"

'sub tempo()
Dim tempo 
Dim genero
genero = "Boa"


If Time < 0.5 Then
    tempo = " dia!"
    genero = "Bom"
ElseIf Time > 0.5 And Time < 0.75 Then
    tempo = " tarde!"
Else
    tempo = " noite!"
'End If
End If
MsgBox "Olá!" & Chr(13) & genero & tempo, , Time
'end sub