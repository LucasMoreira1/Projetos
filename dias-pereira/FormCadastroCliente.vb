'Ajustes
Private Sub txtcep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtcep.MaxLength = 8

    'Para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtrg_Change()
    'Formata : 00.000.000-0
    If Len(txtrg) = 2 Then
        txtrg.Text = txtrg.Text & "."
        
    End If
    If Len(txtrg) = 6 Then
        txtrg.Text = txtrg.Text & "."
        
    End If
     If Len(txtrg) = 10 Then
        txtrg.Text = txtrg.Text & "-"
    End If
End Sub
Private Sub txtrg_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtrg.MaxLength = 12

    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtcpf_Change()
    'Formata : 000.000.000-00
    If Len(txtcpf) = 3 Then
        txtcpf.Text = txtcpf.Text & "."
        
    End If
    If Len(txtcpf) = 7 Then
        txtcpf.Text = txtcpf.Text & "."
        
    End If
     If Len(txtcpf) = 11 Then
        txtcpf.Text = txtcpf.Text & "-"
    End If
End Sub
Private Sub txtcpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtcpf.MaxLength = 14
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
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
Private Sub txttelefone2_Change()
    'Formata : (00)00000-0000
    If Len(txttelefone2) = 1 Then
        txttelefone2.Text = "(" & txttelefone2.Text
        
    End If
    If Len(txttelefone2) = 3 Then
        txttelefone2.Text = txttelefone2.Text & ")"
        
    End If
     If Len(txttelefone2) = 9 Then
        txttelefone2.Text = txttelefone2.Text & "-"
    End If
End Sub
Private Sub txttelefone2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txttelefone2.MaxLength = 14
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtnumero_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtnumero.MaxLength = 5
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If

End Sub
Private Sub txtcnpjreu_Change()
    'Formata : 00.000.000/0000-00
    If Len(txtcnpjreu) = 2 Then
        txtcnpjreu.Text = txtcnpjreu.Text & "."
        
    End If
    If Len(txtcnpjreu) = 6 Then
        txtcnpjreu.Text = txtcnpjreu.Text & "."
        
    End If
     If Len(txtcnpjreu) = 10 Then
        txtcnpjreu.Text = txtcnpjreu.Text & "/"
    End If
    If Len(txtcnpjreu) = 15 Then
        txtcnpjreu.Text = txtcnpjreu.Text & "-"
    End If
End Sub
Private Sub txtcnpjreu_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtcnpjreu.MaxLength = 18
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtcepreu_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtcepreu.MaxLength = 8

    'Para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtnumeroreu_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txtnumeroreu.MaxLength = 5
    'Para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub
Private Sub txttelefonereu_Change()
    'Formata : (00)00000-0000
    If Len(txttelefonereu) = 1 Then
        txttelefonereu.Text = "(" & txttelefonereu.Text
        
    End If
    If Len(txttelefonereu) = 3 Then
        txttelefonereu.Text = txttelefonereu.Text & ")"
        
    End If
     If Len(txttelefonereu) = 9 Then
        txttelefonereu.Text = txttelefonereu.Text & "-"
    End If
End Sub

Private Sub txttelefonereu_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
    txttelefonereu.MaxLength = 14
    
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub


Private Sub imgAdicionarCliente_click()


Dim dds As Worksheet
Set dds = ThisWorkbook.Sheets("Dados")
Dim last_row As Long
last_row = Application.WorksheetFunction.CountA(dds.Range("A:A"))

dds.Range("A" & last_row + 1).Value = last_row

dds.Range("B" & last_row + 1).Value = Me.txtautor.Value
dds.Range("C" & last_row + 1).Value = Me.txtnacionalidade.Value
dds.Range("D" & last_row + 1).Value = Me.txtestadocivil.Value
dds.Range("E" & last_row + 1).Value = Me.txtprofissao.Value
dds.Range("F" & last_row + 1).Value = Me.txtrg.Value
dds.Range("G" & last_row + 1).Value = Me.txtcpf.Value
dds.Range("H" & last_row + 1).Value = Me.txtnascimento.Value
dds.Range("I" & last_row + 1).Value = Me.txtemail.Value
dds.Range("J" & last_row + 1).Value = Me.txttelefone1.Value
dds.Range("K" & last_row + 1).Value = Me.txttelefone2.Value
dds.Range("L" & last_row + 1).Value = Me.txtcep.Value
dds.Range("M" & last_row + 1).Value = Me.txtlogradouro.Value
dds.Range("N" & last_row + 1).Value = Me.txtnumero.Value
dds.Range("O" & last_row + 1).Value = Me.txtcomplemento.Value
dds.Range("P" & last_row + 1).Value = Me.txtbairro.Value
dds.Range("Q" & last_row + 1).Value = Me.txtcidade.Value
dds.Range("R" & last_row + 1).Value = Me.txtestado.Value
dds.Range("S" & last_row + 1).Value = Me.txtreu.Value
dds.Range("T" & last_row + 1).Value = Me.txtcnpjreu.Value
dds.Range("U" & last_row + 1).Value = Me.txttelefonereu.Value
dds.Range("V" & last_row + 1).Value = Me.txtcepreu.Value
dds.Range("W" & last_row + 1).Value = Me.txtlogradouroreu.Value
dds.Range("X" & last_row + 1).Value = Me.txtnumeroreu.Value
dds.Range("Y" & last_row + 1).Value = Me.txtcomplementoreu.Value
dds.Range("Z" & last_row + 1).Value = Me.txtbairroreu.Value
dds.Range("AA" & last_row + 1).Value = Me.txtcidadereu.Value
dds.Range("AB" & last_row + 1).Value = Me.txtestadoreu.Value
dds.Range("AC" & last_row + 1).Value = Me.txtprocesso.Value
dds.Range("AD" & last_row + 1).Value = Me.cboxtipoprocesso.Value
dds.Range("AE" & last_row + 1).Value = Me.txtidprocesso.Value
dds.Range("AF" & last_row + 1).Value = Me.cboxstatusprocesso.Value
dds.Range("AG" & last_row + 1).Value = Me.txtnatprocesso.Value
dds.Range("AH" & last_row + 1).Value = Me.txtassunto1.Value
dds.Range("AI" & last_row + 1).Value = Me.txtassunto2.Value
dds.Range("AJ" & last_row + 1).Value = Me.txtassunto3.Value
dds.Range("AK" & last_row + 1).Value = Me.txtdatapericia.Value
dds.Range("AL" & last_row + 1).Value = Me.txttipoaudiencia.Value
dds.Range("AM" & last_row + 1).Value = Me.txtdataaudiencia.Value
dds.Range("AN" & last_row + 1).Value = Me.txtobservacao.Value
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
