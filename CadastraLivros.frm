VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastraLivros 
   Caption         =   "Cadastro: Livros"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330
   OleObjectBlob   =   "CadastraLivros.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CadastraLivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_gravar_Click()

Dim msg As String
Dim campo As String
Dim x As Long
Dim livro As String
Dim linha As Integer
Dim qtd As Integer
Dim resp As Integer

'''''''''''''''''''''''''''''''
'    testa campos vazios      '
'''''''''''''''''''''''''''''''

msg = "Campo não pode ficar vazio!"

If txt_titulo.Value = "" Then
    campo = "TÍTULO"
    MsgBox msg, vbCritical, campo
    txt_titulo.SetFocus
    txt_titulo.BackColor = rgbMistyRose
    Exit Sub
End If

If txt_autor.Value = "" Then
    campo = "AUTOR"
    MsgBox msg, vbCritical, campo
    txt_autor.SetFocus
    txt_autor.BackColor = rgbMistyRose
    Exit Sub
End If

If txt_editora.Value = "" Then
    campo = "EDITORA"
    MsgBox msg, vbCritical, campo
    txt_editora.SetFocus
    txt_editora.BackColor = rgbMistyRose
    Exit Sub
End If

If txt_genero.Value = "" Then
    campo = "GÊNERO"
    MsgBox msg, vbCritical, campo
    txt_genero.SetFocus
    txt_genero.BackColor = rgbMistyRose
    Exit Sub
End If

If txt_isbn.Value = "" Then
    campo = "FS/IBN"
    MsgBox msg, vbCritical, campo
    txt_isbn.SetFocus
    txt_isbn.BackColor = rgbMistyRose
    Exit Sub
End If

If txt_paradidatico.Value = "" Then
    campo = "PARADIDÁTICO"
    MsgBox msg, vbCritical, campo
    txt_paradidatico.SetFocus
    txt_paradidatico.BackColor = rgbMistyRose
    Exit Sub
End If

If txt_qtd.Value = "" Then
    campo = "QUANTIDADE"
    MsgBox msg, vbCritical, campo
    txt_qtd.SetFocus
    txt_qtd.BackColor = rgbMistyRose
    Exit Sub
End If

'''''''''''''''''''''''''''''''
'  fim teste campos vazios    '
'''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''
'   evita registro duplicado  '
'''''''''''''''''''''''''''''''
livro = UCase(txt_titulo)
Sheets(2).Select
linha = 2
Do Until Cells(linha, 2) = ""
    If UCase(Cells(linha, 2)) = livro Then
        resp = MsgBox("Livro já cadastrado!" & Chr(13) & Chr(13) & "Gostaria de acrescentar apenas uma quantidade específica?", vbYesNo, "INCONSISTÊNCIA")
        If resp < 7 Then
            qtd = InputBox("Entre com a quantidade para o livro " & livro & ":", "Quantidade")
            Cells(linha, 8).Value = Cells(linha, 8).Value + qtd
            Exit Do
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    linha = linha + 1
Loop

'''''''''''''''''''''''''''''''
' fim teste registro duplicado'
'''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''
' Escreve dados conferidos na base'
'''''''''''''''''''''''''''''''''''


Range("B2").End(xlDown).Select
x = ActiveCell.Row + 1
Cells(x, 2) = txt_titulo
Cells(x, 3) = txt_autor
Cells(x, 4) = txt_editora
Cells(x, 5) = txt_genero
Cells(x, 1) = txt_isbn
Cells(x, 7) = txt_paradidatico
Cells(x, 6) = txt_localizacao
Cells(x, 8).Value = txt_qtd.Value

'''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''
' reseta formulário para nova inserção'
'''''''''''''''''''''''''''''''''''''''

txt_titulo.Text = ""
txt_titulo.SetFocus
txt_autor.Text = ""
txt_editora.Text = ""
txt_genero.Text = ""
txt_isbn.Text = ""
txt_paradidatico.Text = ""
txt_localizacao.Text = ""
txt_qtd.Text = ""
''''''''''''''''''''''''''''''''''''''''


End Sub

Private Sub txt_qtd_Change() ''' não deixa usuario inserir letras em quantidade

Dim texto As Variant
Dim letra As Variant
Dim tamanho As Integer
Dim num As Integer

texto = Me.txt_qtd.Value
tamanho = Len(Me.txt_qtd.Value)
For num = 1 To tamanho
    letra = Mid(texto, num, 1)
    If letra <> "" Then
        If letra < Chr(48) Or letra > Chr(57) Then
            Me.txt_qtd.Value = Replace(texto, letra, "")
        End If
    End If
Next num
letra = 0

End Sub

Private Sub txt_titulo_Exit(ByVal Cancel As MSForms.ReturnBoolean) ''' não deixa usuário cadastrar o mesmo livro mais de uma vez
    Dim livro As String
    Dim linha As Long
    Dim flag As Boolean
    Dim temp As String
    Dim x As Variant
    Dim linha_ As Long
    
    livro = UCase(txt_titulo.Text)
    linha = 2
    flag = False
    
    Worksheets(2).Select
    
    While Cells(linha, 2) <> ""
        temp = Cells(linha, 2).Value
        temp = UCase(temp)
        
        If temp = livro Then
            linha_ = ActiveCell.Row
            x = MsgBox("Livro já cadastrado!" & Chr(13) & "Gostaria de adicionar somente em quantidade?", vbYesNo, "Cadastro de livros")
            
            If x = 6 Then
                x = InputBox("Entre com a quantidade:", livro)
                
                If IsNumeric(x) = False Then
                    txt_titulo.Text = ""
                    Unload Me
                    CadastraLivros.Show
                    
                Else
                    Cells(linha_, 8) = Cells(linha_, 8) + x
                    Unload Me
                    CadastraLivros.Show
                    Exit Sub
                End If
            End If
            If x = 7 Then
                Exit Sub
            End If
        End If
        linha = linha + 1
    Wend
    
End Sub

Private Sub UserForm_Terminate()
Planilha1.Activate
End Sub
