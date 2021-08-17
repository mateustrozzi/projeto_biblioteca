VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Devolucao 
   Caption         =   "Devolução - Biblioteca"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   OleObjectBlob   =   "Devolucao.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Devolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbbID_Change()
Dim id As String
Dim linha As Integer
linha = 2
id = cbbId

Planilha4.Select

Do Until Cells(linha, 1) = ""
    If Cells(linha, 1) = id Then
        cbbLocatario = Cells(linha, 2)
    End If
    linha = linha + 1
Loop
End Sub

Private Sub cbbLocatario_Change()
Dim locatario As String
Dim linha As Integer, n As Integer
linha = 2
locatario = cbbLocatario

Planilha4.Select
ListBox1.Clear

Do Until Cells(linha, 2) = ""
    If Cells(linha, 2) = locatario Then
        cbbId = Cells(linha, 1)
    End If
    linha = linha + 1
Loop

ListBox1.ColumnWidths = "350;120;50"
ListBox1.ColumnCount = 3
ListBox1.AddItem
ListBox1.List(0, 0) = "Livro"
ListBox1.List(0, 1) = "ISBN"
ListBox1.List(0, 2) = "Devolução"
n = ListBox1.ListCount

linha = 2
    
Do Until Planilha4.Cells(linha, 2) = ""
    If Planilha4.Cells(linha, 2) = locatario Then
        If Planilha4.Cells(linha, 7) < Date Then
            lblSituacao.Visible = True
        Else
            lblSituacao.Visible = False
        End If
        lblSala = Planilha4.Cells(linha, 3)
        With ListBox1
            .AddItem
            .List(n, 0) = Planilha4.Cells(linha, 4)
            .List(n, 1) = Planilha4.Cells(linha, 5)
            .List(n, 2) = Planilha4.Cells(linha, 7)
        End With
        n = ListBox1.ListCount - 1
    End If
    linha = linha + 1
Loop
End Sub

Private Sub ComboBox1_Change()
Dim livro As String
Dim linha As Integer
'Dim flag As Boolean

'flag = False
linha = 2
livro = ComboBox1

Planilha4.Select

Do Until Cells(linha, 4) = ""
    If Cells(linha, 4) = livro Then
        cbbId = Cells(linha, 1)
        cbbLocatario = Cells(linha, 2)
    End If
    linha = linha + 1
Loop
End Sub

Private Sub ListBox1_Click()
Dim n As Integer
Dim livro As String
Dim linha As Integer
Dim line As Integer
Dim isbn As String
Dim temp As Integer
Dim ok As Boolean
ok = False
n = ListBox1.ListIndex
livro = ListBox1.Text
linha = 2
line = 2
Do Until Planilha4.Cells(linha, 4) = ""
    If Planilha4.Cells(linha, 4) = livro Then
        isbn = Planilha4.Cells(linha, 5)
        Planilha4.Cells(linha, 4).EntireRow.Delete
        
        Do Until Planilha2.Cells(line, 2) = ""
            If Planilha2.Cells(line, 2) = livro And Planilha2.Cells(line, 1) = isbn Then
                Planilha2.Cells(line, 8).Value = Planilha2.Cells(line, 8).Value + 1
                ok = True
                If Planilha2.Cells(line, 8) > 0 Then
                    Planilha2.Cells(line, 9) = "DISPONÍVEL"
                End If
            End If
            line = line + 1
        Loop
    End If
    linha = linha + 1
Loop

If ok Then
    MsgBox UCase(livro) & " foi realocado no sistema!", vbOKOnly, "SUCESSO"
    ok = False
End If


End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' carrega locatarios e ids e seus empréstimos filtrando possíveis repetições '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets(4).Select                                                             '
Dim linha As Integer                                                         '
Range("empr_locatarios").Copy Range("XFD1")                                  '
Range("XFD1").End(xlDown).Select                                             '
linha = ActiveCell.Row                                                       '
Range("XFD1:XFD" & linha).RemoveDuplicates Columns:=1, Header:=xlNo          '
Range("XFD1").End(xlDown).Select                                             '
linha = ActiveCell.Row                                                       '
Range("XFD1:XFD" & linha).Select                                             '
ActiveWorkbook.Names.Add Name:="temp_locatarios", RefersTo:=Selection        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Range("empr_id").Copy Range("XFC1")                                          '
Range("XFC1").End(xlDown).Select                                             '
linha = ActiveCell.Row                                                       '
Range("XFC1:XFC" & linha).RemoveDuplicates Columns:=1, Header:=xlNo          '
Range("XFC1").End(xlDown).Select                                             '
linha = ActiveCell.Row                                                       '
Range("XFC1:XFC" & linha).Select                                             '
ActiveWorkbook.Names.Add Name:="temp_locatarios_id", RefersTo:=Selection     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
cbbLocatario.RowSource = Range("temp_locatarios").Address                    '
cbbId.RowSource = Range("temp_locatarios_id").Address                        '
Planilha4.Range("A2").EntireRow.Select 'leva o foco da planilha ao início    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub UserForm_Terminate()
Planilha1.Activate
End Sub
