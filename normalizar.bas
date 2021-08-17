Attribute VB_Name = "normalizar"
Sub normalizar()

'''' C�lulas em branco v�o causar inconsist�ncias na extra��o ou inser��o de dados,
''''    ent�o ser� acrescido em c�lulas de linhas preenchidas de modo irregular a palavra NULL
''''    para corrigir essa situa��o at� que o usu�rio as fa�a.

    Dim linha As Integer
    Dim coluna As Integer
    Dim ultima_preenchida As Integer
    
    ultima_preenchida = InputBox("Entre com o n�mero da �ltima c�lula preenchida para normalizar:", "Normalizar")
    
    Worksheets(2).Select
    For coluna = 1 To 9
        For linha = 2 To ultima_preenchida
            If Cells(linha, coluna) = Empty Then
                Cells(linha, coluna) = "NULL"
            End If
        Next linha
    Next coluna
    
End Sub
