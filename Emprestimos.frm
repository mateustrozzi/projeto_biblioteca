VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Emprestimos 
   Caption         =   "Emrpéstimos - Biblioteca"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12630
   OleObjectBlob   =   "Emprestimos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Emprestimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbbID_Change()

Dim linha As Long
Dim id As String
id = cbbId
linha = 2

Do Until Planilha3.Cells(linha, 2) = ""
    If Planilha3.Cells(linha, 1) = id Then
        cbbLocatario = Planilha3.Cells(linha, 2)
        lblSala.Caption = Planilha3.Cells(linha, 3)
        Exit Sub
    End If
    linha = linha + 1
Loop
End Sub

Private Sub cbbLivro_Change()

Dim linha As Long
Dim livro As String
livro = cbbLivro
linha = 2

Do Until Planilha2.Cells(linha, 2) = ""
    If Planilha2.Cells(linha, 2) = livro Then
        If Planilha2.Cells(linha, 8) > 1 Then
            lblExemplares.Caption = Planilha2.Cells(linha, 8) & " exemplares"
        ElseIf Planilha2.Cells(linha, 8) = 1 Then
            lblExemplares.Caption = Planilha2.Cells(linha, 8) & " exemplar"
        Else
            lblExemplares.Caption = "INDISPONÍVEL"
        End If
        Exit Sub
    End If
    linha = linha + 1
Loop

End Sub

Private Sub cbbLocatario_Change()

Dim linha As Long
Dim locatario As String
locatario = cbbLocatario
linha = 2

Do Until Planilha3.Cells(linha, 2) = ""
    If Planilha3.Cells(linha, 2) = locatario Then
        cbbId = Planilha3.Cells(linha, 1)
        lblSala.Caption = Planilha3.Cells(linha, 3)
        Exit Sub
    End If
    linha = linha + 1
Loop

        
End Sub

Private Sub CommandButton1_Click()

Dim linha As Long
Dim line As Long
Dim locatario As String
Dim data As Date
Dim cont As Integer
Dim ok As Boolean

ok = True
locatario = cbbLocatario
linha = 2
data = Date

If lblExemplares = "INDISPONÍVEL" Then
    MsgBox UCase(cbbLivro) & " está indisponível no momento..." & Chr(13) & _
        "Se você tem uma cópia em mãos, primeiro cadastre-a no sistema!", vbOKOnly, _
        "BIBLIOTECA"
        
    cbbId = ""
    cbbLivro = ""
    cbbLocatario = ""
    lblExemplares = ""
    lblSala = ""
    Exit Sub
End If
        
Do Until Planilha4.Cells(linha, 2) = ""
    If Planilha4.Cells(linha, 2) = locatario Then
        If Planilha4.Cells(linha, 7) <= data Then
            MsgBox UCase(locatario) & " tem um empréstimo vigente!", vbCritical, "BIBLIOTECA"
            ok = False
            Planilha4.Select
            Cells(linha, 7).EntireRow.Select
            Unload Me
            Exit Do
            Exit Sub
        End If
    End If
    linha = linha + 1
Loop

cont = 0
line = 2

Do Until Planilha4.Cells(line, 2) = ""
    If Planilha4.Cells(line, 2) = locatario And ok = True Then
        cont = cont + 1
        If cont > 4 Then
            MsgBox UCase(locatario) & " já atingiu a cota máxima de locações!", vbCritical, "BIBLIOTECA"
            ok = False
            Planilha4.Select
            Cells(line, 2).EntireRow.Select
            Unload Me
            Exit Do
            Exit Sub
        End If
    End If
    line = line + 1
Loop

If ok Then
    linha = Planilha4.Range("A1").End(xlDown).Row + 1
    Planilha4.Cells(linha, 1) = cbbId.Value
    Planilha4.Cells(linha, 2) = cbbLocatario
    Planilha4.Cells(linha, 3) = lblSala
    Planilha4.Cells(linha, 4) = cbbLivro
    Planilha4.Cells(linha, 6) = data
    Planilha4.Cells(linha, 7) = data + 7
    line = 2
    Do Until Planilha2.Cells(line, 2) = ""
        If Planilha2.Cells(line, 2) = cbbLivro Then
            Planilha4.Cells(linha, 5) = Planilha2.Cells(line, 1)
            Planilha2.Cells(line, 8) = Planilha2.Cells(line, 8) - 1
            If Planilha2.Cells(line, 8) < 1 Then
                Planilha2.Cells(line, 9) = "INDISPONÍVEL"
            End If
            Exit Do
        End If
        line = line + 1
    Loop
    
    
    MsgBox "Empréstimo registrado no sistema com êxito", vbOKOnly, "BIBLIOTECA"
    
    
    cbbId = ""
    cbbLivro = ""
    cbbLocatario = ""
    lblExemplares = ""
    lblSala = ""
    Exit Sub
End If
End Sub


Private Sub UserForm_Initialize()

Dim users As Range
Planilha3.Select
cbbLocatario.RowSource = Planilha3.Range("locatarios").Address
cbbId.RowSource = Planilha3.Range("Id").Address
Planilha2.Select
cbbLivro.RowSource = Planilha2.Range("livros").Address
End Sub

Private Sub UserForm_Terminate()
Planilha1.Activate
End Sub
