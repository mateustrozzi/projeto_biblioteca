Attribute VB_Name = "normalizar"
Sub normalizar()

'''' Células em branco vão causar inconsistências na extração ou inserção de dados,
''''    então será acrescido em células de linhas preenchidas de modo irregular a palavra NULL
''''    para corrigir essa situação até que o usuário as faça.

    Dim linha As Integer
    Dim coluna As Integer
    Dim ultima_preenchida As Integer
    
    ultima_preenchida = InputBox("Entre com o número da última célula preenchida para normalizar:", "Normalizar")
    
    Worksheets(2).Select
    For coluna = 1 To 9
        For linha = 2 To ultima_preenchida
            If Cells(linha, coluna) = Empty Then
                Cells(linha, coluna) = "NULL"
            End If
        Next linha
    Next coluna
    
End Sub
