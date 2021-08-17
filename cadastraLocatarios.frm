VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadastraLocatarios 
   Caption         =   "Locatários"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12195
   OleObjectBlob   =   "cadastraLocatarios.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadastraLocatarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCadastrar_Click()
    Dim linha As Long
    Dim id As Variant
    
    Worksheets(3).Select
    linha = Range("A2").End(xlDown).Row
    id = Cells(linha, 1) + 1
    Cells(linha + 1, 1) = id
    
    Cells(linha + 1, 2) = txtLocatario
    
    If optAluno.Value = True Then
        Cells(linha + 1, 3) = ComboBox1.Value & " " & ComboBox2.Value
    Else
        Cells(linha + 1, 3) = "Professor"
    End If
    
    frm1.Visible = False
    txtLocatario.Text = ""
    txtLocatario.SetFocus
    btnRecarregar.Visible = False
    btnLocalizar.Visible = True
    lblMensagem.Visible = True
    
    
End Sub

Private Sub btnLocalizar_Click()

    Dim locatario As String
    Dim linha As Integer
    Dim flag As Boolean
    Dim temp As String
    
    Worksheets(3).Select
    
    ComboBox1.Visible = True
    ComboBox2.Visible = True
    optAluno.Value = True
    flag = False
    
    linha = 2
    locatario = txtLocatario.Text
    locatario = UCase(locatario)
        
    While Cells(linha, 2) <> ""
        temp = Cells(linha, 2).Value
        temp = UCase(temp)
        
        If temp = locatario Then
            MsgBox locatario & " já está cadastrado!", vbOKOnly, "Biblioteca"
            flag = True
            txtLocatario.Text = ""
            btnLocalizar.Visible = True
            btnRecarregar.Visible = False
            txtLocatario.SetFocus
            Exit Sub
        End If
        
    If txtLocatario.Value = "" Then
        MsgBox "Campo LOCATÁRIO não pode ficar vazia!", vbOKOnly, "Entre com um nome"
        txtLocatario.SetFocus
        flag = True
        Exit Sub
    End If
    
    linha = linha + 1
    Wend
    
    If flag = False Then
        frm1.Visible = True
        btnLocalizar.Visible = False
        btnRecarregar.Visible = True
        ComboBox1.SetFocus
    End If

End Sub

Private Sub btnRecarregar_Click()

    frm1.Visible = False
    txtLocatario.Text = ""
    txtLocatario.SetFocus
    btnRecarregar.Visible = False
    btnLocalizar.Visible = True
    lblMensagem.Visible = False

End Sub

Private Sub optAluno_Click()

    ComboBox1.Visible = True
    ComboBox2.Visible = True
    
End Sub

Private Sub optProfessor_Click()

    ComboBox1.Visible = False
    ComboBox2.Visible = False
    
End Sub

Private Sub txtLocatario_Change()
lblMensagem.Visible = False
End Sub

Private Sub UserForm_Initialize()

    ComboBox1.AddItem "1º ano"
    ComboBox1.AddItem "2º ano"
    ComboBox1.AddItem "3º ano"
    ComboBox1.AddItem "4º ano"
    ComboBox1.AddItem "5º ano"
    ComboBox1.AddItem "6º ano"
    ComboBox1.AddItem "7º ano"
    ComboBox1.AddItem "8º ano"
    ComboBox1.AddItem "9º ano"
    ComboBox1.ListIndex = 0
    
    ComboBox2.AddItem "A"
    ComboBox2.AddItem "B"
    ComboBox2.AddItem "C"
    ComboBox2.ListIndex = 0
    
End Sub

Private Sub UserForm_Terminate()
Planilha1.Activate
End Sub
