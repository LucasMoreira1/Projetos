Private Sub imgAtualizarCliente_click()

If PaginaInicial.TextBox8.Value = Me. "" Then
MsgBox "Selecione o usuário para editar."
Exit Sub
End If

Dim dds As Worksheet
Set dds = ThisWorkbook.Sheets("Dados")
Dim selected_row As Long
selected_row = Application.WorksheetFunction.Match(PaginaInicial.TextBox8.Value, dds.Range("B:B"), 0)


If selected_row <> 1 Then
    dds.Range("B" & selected_row).Value = Me.txtautor.Value
    dds.Range("C" & selected_row).Value = Me.txtnacionalidade.Value
    dds.Range("D" & selected_row).Value = Me.txtestadocivil.Value
    dds.Range("E" & selected_row).Value = Me.txtprofissao.Value
    dds.Range("F" & selected_row).Value = Me.txtrg.Value
    dds.Range("G" & selected_row).Value = Me.txtcpf.Value
    dds.Range("H" & selected_row).Value = Me.txtnascimento.Value
    dds.Range("I" & selected_row).Value = Me.txtemail.Value
    dds.Range("J" & selected_row).Value = Me.txttelefone1.Value
    dds.Range("K" & selected_row).Value = Me.txttelefone2.Value
    dds.Range("L" & selected_row).Value = Me.txtcep.Value
    dds.Range("M" & selected_row).Value = Me.txtlogradouro.Value
    dds.Range("N" & selected_row).Value = Me.txtnumero.Value
    dds.Range("O" & selected_row).Value = Me.txtcomplemento.Value
    dds.Range("P" & selected_row).Value = Me.txtbairro.Value
    dds.Range("Q" & selected_row).Value = Me.txtcidade.Value
    dds.Range("R" & selected_row).Value = Me.txtestado.Value
    dds.Range("S" & selected_row).Value = Me.txtreu.Value
    dds.Range("T" & selected_row).Value = Me.txtcnpjreu.Value
    dds.Range("U" & selected_row).Value = Me.txttelefonereu.Value
    dds.Range("V" & selected_row).Value = Me.txtcepreu.Value
    dds.Range("W" & selected_row).Value = Me.txtlogradouroreu.Value
    dds.Range("X" & selected_row).Value = Me.txtnumeroreu.Value
    dds.Range("Y" & selected_row).Value = Me.txtcomplementoreu.Value
    dds.Range("Z" & selected_row).Value = Me.txtbairroreu.Value
    dds.Range("AA" & selected_row).Value = Me.txtcidadereu.Value
    dds.Range("AB" & selected_row).Value = Me.txtestadoreu.Value
    dds.Range("AC" & selected_row).Value = Me.txtprocesso.Value
    dds.Range("AD" & selected_row).Value = Me.cboxtipoprocesso.Value
    dds.Range("AE" & selected_row).Value = Me.txtidprocesso.Value
    dds.Range("AF" & selected_row).Value = Me.cboxstatusprocesso.Value
    dds.Range("AG" & selected_row).Value = Me.txtnatprocesso.Value
    dds.Range("AH" & selected_row).Value = Me.txtassunto1.Value
    dds.Range("AI" & selected_row).Value = Me.txtassunto2.Value
    dds.Range("AJ" & selected_row).Value = Me.txtassunto3.Value
    dds.Range("AK" & selected_row).Value = Me.txtdatapericia.Value
    dds.Range("AL" & selected_row).Value = Me.txttipoaudiencia.Value
    dds.Range("AM" & selected_row).Value = Me.txtdataaudiencia.Value
    dds.Range("AN" & selected_row).Value = Me.txtobservacao.Value




Else
Exit Sub
End If

'---------------------------------------------------
'Limpar dados do formulário
Call LimparDadosEditar
'---------------------------------------------------
MsgBox "Cliente atualizado com sucesso"
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
    
    'para permitir que apenas números sejam digitados
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
Call BuscaCEP_autor_editar
End Sub
'Função Busca CEP réu - CadastroCliente
Private Sub btnbuscacepreu_Click()
Call BuscaCEP_reu_editar
End Sub
'Limpar dados formulário
Private Sub imgLimparCliente_click()
Call LimparDadosEditar
End Sub

'                                   -----------------------------
'                                   | Efeitos visuais de botões |
'                                   -----------------------------

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call efeitos_icones
End Sub
Private Sub imgLimparCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
imgLimparCliente.BorderStyle = fmBorderStyleSingle
imgLimparCliente.SpecialEffect = fmSpecialEffectSunken
imgLimparCliente.ControlTipText = "Limpar dados"
End Sub
Private Sub imgAtualizarCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
imgAtualizarCliente.BorderStyle = fmBorderStyleSingle
imgAtualizarCliente.SpecialEffect = fmSpecialEffectSunken
imgAtualizarCliente.ControlTipText = "Atualizar Cliente"
End Sub
