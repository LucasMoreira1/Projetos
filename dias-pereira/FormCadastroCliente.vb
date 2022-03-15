

Private Sub imgAdicionarCliente_click()


Dim dds As Worksheet
Set dds = ThisWorkbook.Sheets("Dados")
Dim last_row As Long
last_row = Application.WorksheetFunction.CountA(dds.Range("A:A"))

dds.Range("A" & last_row + 1).Value = last_row

dds.Range("B" & last_row + 1).Value = Me.txtautor.Value
dds.Range("C" & last_row + 1).Value = Me.txtreu.Value

dds.Range("D" & last_row + 1).Value = Me.txttelefone1.Value
dds.Range("F" & last_row + 1).Value = Me.txtnascimento.Value
dds.Range("G" & last_row + 1).Value = Me.txtidprocesso.Value
dds.Range("H" & last_row + 1).Value = Me.txtprocesso.Value
dds.Range("I" & last_row + 1).Value = Me.cboxstatusprocesso.Value
dds.Range("J" & last_row + 1).Value = Me.txtnatprocesso.Value
dds.Range("K" & last_row + 1).Value = Me.txtassunto1.Value
dds.Range("L" & last_row + 1).Value = Me.txtassunto2.Value
dds.Range("M" & last_row + 1).Value = Me.txtassunto3.Value

dds.Range("N" & last_row + 1).Value = Me.txtdatapericia.Value
dds.Range("O" & last_row + 1).Value = Me.txtdataaudiencia.Value
dds.Range("P" & last_row + 1).Value = Me.txttipoaudiencia.Value
dds.Range("Q" & last_row + 1).Value = Me.txtobservacao.Value
'---------------------------------------------------
Call LimparDadosCadastro
'---------------------------------------------------
MsgBox "Cliente adicionado com sucesso"
Unload Me

Call PaginaInicial.Filtro


End Sub





Private Sub txtdataaudiencia_Change()
    'Formata : dd/mm/aaaa
    If Len(txtdataaudiencia) = 2 Or Len(txtdataaudiencia) = 5 Then
        txtdataaudiencia.Text = txtdataaudiencia.Text & "/"
        
    End If
End Sub

Private Sub txtdataaudiencia_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtdataaudiencia.MaxLength = 10
    
    'Para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtdatapericia_Change()
    'Formata : dd/mm/aaaa
    If Len(txtdatapericia) = 2 Or Len(txtdatapericia) = 5 Then
        txtdatapericia.Text = txtdatapericia.Text & "/"
        
    End If
End Sub

Private Sub txtdatapericia_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtdatapericia.MaxLength = 10
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtnascimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtnascimento.MaxLength = 10
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtnascimento_Change()
    'Formata : dd/mm/aaaa
    If Len(txtnascimento) = 2 Or Len(txtnascimento) = 5 Then
        txtnascimento.Text = txtnascimento.Text & "/"
        
    End If
  
End Sub

Private Sub txtprocesso_Change()
    
If Me.cboxtipoprocesso.Text = "Judicial" Then
    'Formata : 0000000-00.0000.0.00.0000
        If Len(txtprocesso) = 7 Then
            txtprocesso.Text = txtprocesso.Text & "-"
            
        End If
        If Len(txtprocesso) = 10 Then
            txtprocesso.Text = txtprocesso.Text & "."
           
        End If
         If Len(txtprocesso) = 15 Then
            txtprocesso.Text = txtprocesso.Text & "."
            
        End If
        If Len(txtprocesso) = 17 Then
            txtprocesso.Text = txtprocesso.Text & "."
            
        End If
        If Len(txtprocesso) = 20 Then
            txtprocesso.Text = txtprocesso.Text & "."
           
        End If
        
End If

End Sub


Private Sub txtprocesso_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtprocesso.MaxLength = 25
    
    'para permitir que apenas números sejam digitados
'    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
'        KeyAscii = 0
'    End If
End Sub

Private Sub txttelefone1_Change()
    'Formata : (00)00000-0000
    If Len(txttelefone1) = 1 Then
        txttelefone1.Text = "(" & txttelefone1.Text
        
    End If
    If Len(txttelefone1) = 3 Then
        txttelefone1.Text = txttelefone1.Text & ")"
        
    End If
     If Len(txttelefone1) = 9 Then
        txttelefone1.Text = txttelefone1.Text & "-"
    End If
End Sub

Private Sub txttelefone1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txttelefone1.MaxLength = 14
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Activate()
With Me.cboxtipoprocesso
    .Clear
    .AddItem "Judicial"
    .AddItem "Inquerito Policial"
    .AddItem "Administrativo"
End With

With Me.cboxstatusprocesso
    .Clear
    .AddItem "Arquivado"
    .AddItem "Em andamento"
    .AddItem "Suspenso"
End With
cboxtipoprocesso.Text = "Judicial"
cboxstatusprocesso.Text = "Em andamento"

Me.ScrollTop = 0

End Sub
'                                   -----------------------------
'                                   |       Ações e Funções     |
'                                   -----------------------------

'Função Busca CEP autor - CadastroCliente
Private Sub btnbuscacep_Click()
Call BuscaCEP_autor_cadastro
End Sub
'Função Busca CEP réu - CadastroCliente
Private Sub btnbuscacepreu_Click()
Call BuscaCEP_reu_cadastro
End Sub
'Limpar dados formulário
Private Sub imgSairCliente_click()
Call LimparDadosCadastro
End Sub

'                                   -----------------------------
'                                   | Efeitos visuais de botões |
'                                   -----------------------------

Private Sub imgAdicionarCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
imgAdicionarCliente.BorderStyle = fmBorderStyleSingle
imgAdicionarCliente.SpecialEffect = fmSpecialEffectSunken
imgAdicionarCliente.ControlTipText = "Adicionar Cliente"
End Sub
Private Sub imgSairCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
imgSairCliente.BorderStyle = fmBorderStyleSingle
imgSairCliente.SpecialEffect = fmSpecialEffectSunken
imgSairCliente.ControlTipText = "Limpar dados"
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call efeitos_icones
End Sub
