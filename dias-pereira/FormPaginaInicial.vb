''
'' Sistema para banco de dados de clientes e processos jurídicos. Com funções como: Mensagens via whatsapp automatizadas.
'' Criado para 'Dias Pereira Advocacia'
'' Software desenvolvido por Agility IT/Lucas Moreira (12)99732-9778 - Lucas_moreira@outlook.com
''
'' Revisão 11/03/2022
'' Modificações:
''      - Add 4 colunas: Reu, Assunto1, Assunto2 e Assunto3.
''      - Alteração na coluna "nome" para "autor" e devidos objetos.
''
'' Revisão 13/03/2022
'' Modificações:
''      - Alterado nome dos forms: "UserForm6" para "CadastroCliente" e "UserForm9" para "EditarCliente".
''      - Alteração no Layout das forms "CadastroCliente" e "EditarCliente".
''      - Criado Modulo "Functions" para centralizar funções, limpando o codigo dos Forms.
''
'' Revisão 15/03/2022
'' Modificações:
''      - Botão "Busca CEP" para preencher dados de endereço durante cadastro e edição de cliente.
''      - Adicionar colunas na base do Excel referente as colunas criadas na ultima alteração de layout.
''      - Criado funções "LimparDadosCliente" e "LimparDadosCadastro" no modulo "Functions".
''      - Ajustar Codigos para preencher colunas no Excel (CadastroCliente e EditarCliente).
''      - Ajustar Lista LV1 para mostrar corretamente os dados das colunas.
''
'' Revisão 17/03/2022
'' Modificações:
''      - Minimizar APP
''      - Ajustar campos númericos e de data nos formulários de Cadastro e Edição.
''      - Adicionado Label e função para mostrar idade ao preencher a data de nascimento
''      - Add Botões de Fluxo de Caixa e Botão para Criar petição.
''
''
'' A fazer:
''      - Personalizar Lista LV1 para mostrar apenas dados importantes. (Remover lista/Menos info possivel).
''          - Testar pesquisa (Filtro) com esse ajuste na Lista LV1.
''      - Continuar migrando funções para o modulo "Functions"
''
''

'                                   -----------------------------
'                                   |      Listas e ajustes     |
'                                   -----------------------------

Private Sub LV1_DblClick()
TextBox8.Value = LV1.SelectedItem.SubItems(1)

Call Editar_Cliente

End Sub
Private Sub imgEditarCliente_click()
TextBox8.Value = LV1.SelectedItem.SubItems(1)

Call Editar_Cliente

End Sub

Public Sub UserForm_Initialize()
'Chama função para habilitar botões de maximizar e minimizar
    SetMaxMin Me.Caption

MultiPage1.BackColor = vbWhite
'Refresh de dados
Call atualiza_usuarios
Call atualiza_mensagem
'Seleciona a planilha
ThisWorkbook.Sheets("Dados").Select

plan = ThisWorkbook.Sheets("Dados").Name
Dim D As Integer
    D = Range("Tabelas!B3").End(xlDown).Row
    ComboBox1.RowSource = "Tabelas!B3:B" & D
    ComboBox2.RowSource = "Tabelas!B3:B" & D
    ComboBox3.RowSource = "Tabelas!B3:B" & D
Dim li As ListItem
Dim lsi As ListSubItem
Dim r As Long
Dim C As Long

ultimacelula = Cells(ActiveSheet.UsedRange.Rows.Count, 1).Row
'Lista 1
With Me.LV1
    .View = lvwReport
    .Gridlines = True
    .HideColumnHeaders = False
End With
'Colunas
With Me.LV1
    For C = 0 To 39
     .ColumnHeaders.Add , , Cells(1, C + 1), Cells(1, C + 1).Width
    Next C
End With
'Campos das colunas
With Me.LV1
    For r = 2 To ultimacelula
        Set li = .ListItems.Add(, , Cells(r, 1)) 'id
        li.ListSubItems.Add , , Cells(r, 2) 'autor
        li.ListSubItems.Add , , Cells(r, 3) 'nacionalidade
        li.ListSubItems.Add , , Cells(r, 4) 'estadocivil
        li.ListSubItems.Add , , Cells(r, 5) 'profissao
        li.ListSubItems.Add , , Cells(r, 6) 'rg
        li.ListSubItems.Add , , Cells(r, 7) 'cpf
        li.ListSubItems.Add , , Cells(r, 8) 'nascimento
        li.ListSubItems.Add , , Cells(r, 9) 'email
        li.ListSubItems.Add , , Cells(r, 10) 'telefone1
        li.ListSubItems.Add , , Cells(r, 11) 'telefone2
        li.ListSubItems.Add , , Cells(r, 12) 'cep
        li.ListSubItems.Add , , Cells(r, 13) 'logradouro
        li.ListSubItems.Add , , Cells(r, 14) 'numero
        li.ListSubItems.Add , , Cells(r, 15) 'complemento
        li.ListSubItems.Add , , Cells(r, 16) 'bairro
        li.ListSubItems.Add , , Cells(r, 17) 'cidade
        li.ListSubItems.Add , , Cells(r, 18) 'estado
        li.ListSubItems.Add , , Cells(r, 19) 'reu
        li.ListSubItems.Add , , Cells(r, 20) 'cnpjreu
        li.ListSubItems.Add , , Cells(r, 21) 'telefonereu
        li.ListSubItems.Add , , Cells(r, 22) 'cepreu
        li.ListSubItems.Add , , Cells(r, 23) 'logradouroreu
        li.ListSubItems.Add , , Cells(r, 24) 'numeroreu
        li.ListSubItems.Add , , Cells(r, 25) 'complementoreu
        li.ListSubItems.Add , , Cells(r, 26) 'bairroreu
        li.ListSubItems.Add , , Cells(r, 27) 'cidadereu
        li.ListSubItems.Add , , Cells(r, 28) 'estadoreu
        li.ListSubItems.Add , , Cells(r, 29) 'processo
        li.ListSubItems.Add , , Cells(r, 30) 'idprocesso
        li.ListSubItems.Add , , Cells(r, 31) 'natprocesso
        li.ListSubItems.Add , , Cells(r, 32) 'assunto1
        li.ListSubItems.Add , , Cells(r, 33) 'assunto2
        li.ListSubItems.Add , , Cells(r, 34) 'assunto3
        li.ListSubItems.Add , , Cells(r, 35) 'datapericia
        li.ListSubItems.Add , , Cells(r, 36)
        li.ListSubItems.Add , , Cells(r, 37)
        li.ListSubItems.Add , , Cells(r, 38)
        li.ListSubItems.Add , , Cells(r, 39) 'tipoaudiencia
        li.ListSubItems.Add , , Cells(r, 40) 'dataaudiencia
       
    Next r
End With
End Sub
Private Sub imgExecutar_click()
On Error GoTo Erro
   'Declarar todas as nossas variaveis
        Dim driver As WebDriver
        Dim cd As New Selenium.ChromeDriver
        Dim por As by
        Dim keys As New Selenium.keys
        Dim nome, contato, arquivo1, arquivo2, Msg, URLweb, URLnumero As String
        Dim msgPerso As String
        Dim ultimacelula, i As Integer
        Dim FindBy As New Selenium.by
        Dim textBox As String
        Dim linha As Long
    'Selecionar a nossa folha de trabalho
        ThisWorkbook.Sheets("Dados").Select
        plan = ThisWorkbook.Sheets("Dados").Name
    'Atribuir valores as nossas variaveis
        URLweb = "https://web.whatsapp.com"
        textBox = PaginaInicial.TB1.Value
        Tabelas = Planilha2.Name
        linha = 2
        arquivo1 = ThisWorkbook.Sheets("Mensagens").Range("H2")
        'Inicio do codigo
        LV1.ListItems.Clear
        With cd
        '1. Abrir nosso navegador
            .Start
        '2. Entrar no site do whatsapp web: https://web.whatsapp.com
            .Get URLweb
        'Tempo para ler o QR code e carregar pagina do Whatsapp
            MsgBox "Clique aqui após Scanear o QR Code", vbSystemModal, "Leia o QR Code com o celular"
            Application.Wait Now + TimeValue("00:00:02")
            Application.SendKeys ("%{TAB}")
            DoEvents
        Select Case ComboBox1
            Case Planilha2.Range("B3").Text
            ColunaC1 = 1
            Case Planilha2.Range("B4").Text
            ColunaC1 = 2
            Case Planilha2.Range("B5").Text
            ColunaC1 = 3
            Case Planilha2.Range("B6").Text
            ColunaC1 = 4
            Case Planilha2.Range("B7").Text
            ColunaC1 = 5
            Case Planilha2.Range("B8").Text
            ColunaC1 = 6
            Case Planilha2.Range("B9").Text
            ColunaC1 = 7
            Case Planilha2.Range("B10").Text
            ColunaC1 = 8
            Case Planilha2.Range("B11").Text
            ColunaC1 = 9
            Case Planilha2.Range("B12").Text
            ColunaC1 = 10
            Case Planilha2.Range("B13").Text
            ColunaC1 = 11
            Case Planilha2.Range("B14").Text
            ColunaC1 = 12
            Case Planilha2.Range("B15").Text
            ColunaC1 = 13
            Case Planilha2.Range("B16").Text
            ColunaC1 = 14
            Case Planilha2.Range("B17").Text
            ColunaC1 = 15
            Case Planilha2.Range("B18").Text
            ColunaC1 = 16
            Case Planilha2.Range("B19").Text
            ColunaC1 = 17
            Case Planilha2.Range("B20").Text
            ColunaC1 = 18
            Case Planilha2.Range("B21").Text
            ColunaC1 = 19
            Case Planilha2.Range("B22").Text
            ColunaC1 = 20
            Case Planilha2.Range("B23").Text
            ColunaC1 = 21
            Case Planilha2.Range("B24").Text
            ColunaC1 = 22
            Case Planilha2.Range("B25").Text
            ColunaC1 = 23
            Case Planilha2.Range("B26").Text
            ColunaC1 = 24
            Case Planilha2.Range("B27").Text
            ColunaC1 = 25
            Case Planilha2.Range("B28").Text
            ColunaC1 = 26
            Case Planilha2.Range("B29").Text
            ColunaC1 = 27
            Case Planilha2.Range("B30").Text
            ColunaC1 = 28
            Case Planilha2.Range("B31").Text
            ColunaC1 = 29
            Case Planilha2.Range("B32").Text
            ColunaC1 = 30
            Case Planilha2.Range("B33").Text
            ColunaC1 = 31
            Case Planilha2.Range("B34").Text
            ColunaC1 = 32
            Case Planilha2.Range("B35").Text
            ColunaC1 = 33
            Case Planilha2.Range("B36").Text
            ColunaC1 = 34
            Case Planilha2.Range("B37").Text
            ColunaC1 = 35
            Case Planilha2.Range("B38").Text
            ColunaC1 = 36
            Case Planilha2.Range("B39").Text
            ColunaC1 = 37
            Case Planilha2.Range("B40").Text
            ColunaC1 = 38
            Case Planilha2.Range("B41").Text
            ColunaC1 = 39
            Case Planilha2.Range("B42").Text
            ColunaC1 = 40
        End Select

        If ComboBox1.Value = "Filtro 1" Then
            ColunaC1 = 1
        End If
        
        Select Case ComboBox2
            Case Planilha2.Range("B3").Text
            ColunaC2 = 1
            Case Planilha2.Range("B4").Text
            ColunaC2 = 2
            Case Planilha2.Range("B5").Text
            ColunaC2 = 3
            Case Planilha2.Range("B6").Text
            ColunaC2 = 4
            Case Planilha2.Range("B7").Text
            ColunaC2 = 5
            Case Planilha2.Range("B8").Text
            ColunaC2 = 6
            Case Planilha2.Range("B9").Text
            ColunaC2 = 7
            Case Planilha2.Range("B10").Text
            ColunaC2 = 8
            Case Planilha2.Range("B11").Text
            ColunaC2 = 9
            Case Planilha2.Range("B12").Text
            ColunaC2 = 10
            Case Planilha2.Range("B13").Text
            ColunaC2 = 11
            Case Planilha2.Range("B14").Text
            ColunaC2 = 12
            Case Planilha2.Range("B15").Text
            ColunaC2 = 13
            Case Planilha2.Range("B16").Text
            ColunaC2 = 14
            Case Planilha2.Range("B17").Text
            ColunaC2 = 15
            Case Planilha2.Range("B18").Text
            ColunaC2 = 16
            Case Planilha2.Range("B19").Text
            ColunaC2 = 17
            Case Planilha2.Range("B20").Text
            ColunaC2 = 18
            Case Planilha2.Range("B21").Text
            ColunaC2 = 19
            Case Planilha2.Range("B22").Text
            ColunaC2 = 20
            Case Planilha2.Range("B23").Text
            ColunaC2 = 21
            Case Planilha2.Range("B24").Text
            ColunaC2 = 22
            Case Planilha2.Range("B25").Text
            ColunaC2 = 23
            Case Planilha2.Range("B26").Text
            ColunaC2 = 24
            Case Planilha2.Range("B27").Text
            ColunaC2 = 25
            Case Planilha2.Range("B28").Text
            ColunaC2 = 26
            Case Planilha2.Range("B29").Text
            ColunaC2 = 27
            Case Planilha2.Range("B30").Text
            ColunaC2 = 28
            Case Planilha2.Range("B31").Text
            ColunaC2 = 29
            Case Planilha2.Range("B32").Text
            ColunaC2 = 30
            Case Planilha2.Range("B33").Text
            ColunaC2 = 31
            Case Planilha2.Range("B34").Text
            ColunaC2 = 32
            Case Planilha2.Range("B35").Text
            ColunaC2 = 33
            Case Planilha2.Range("B36").Text
            ColunaC2 = 34
            Case Planilha2.Range("B37").Text
            ColunaC2 = 35
            Case Planilha2.Range("B38").Text
            ColunaC2 = 36
            Case Planilha2.Range("B39").Text
            ColunaC2 = 37
            Case Planilha2.Range("B40").Text
            ColunaC2 = 38
            Case Planilha2.Range("B41").Text
            ColunaC2 = 39
            Case Planilha2.Range("B42").Text
            ColunaC2 = 40
       End Select
        
        If ComboBox2.Value = "Filtro 2" Then
            ColunaC2 = 1
        End If
        
        Select Case ComboBox3
            Case Planilha2.Range("B3").Text
            ColunaC3 = 1
            Case Planilha2.Range("B4").Text
            ColunaC3 = 2
            Case Planilha2.Range("B5").Text
            ColunaC3 = 3
            Case Planilha2.Range("B6").Text
            ColunaC3 = 4
            Case Planilha2.Range("B7").Text
            ColunaC3 = 5
            Case Planilha2.Range("B8").Text
            ColunaC3 = 6
            Case Planilha2.Range("B9").Text
            ColunaC3 = 7
            Case Planilha2.Range("B10").Text
            ColunaC3 = 8
            Case Planilha2.Range("B11").Text
            ColunaC3 = 9
            Case Planilha2.Range("B12").Text
            ColunaC3 = 10
            Case Planilha2.Range("B13").Text
            ColunaC3 = 11
            Case Planilha2.Range("B14").Text
            ColunaC3 = 12
            Case Planilha2.Range("B15").Text
            ColunaC3 = 13
            Case Planilha2.Range("B16").Text
            ColunaC3 = 14
            Case Planilha2.Range("B17").Text
            ColunaC3 = 15
            Case Planilha2.Range("B18").Text
            ColunaC3 = 16
            Case Planilha2.Range("B19").Text
            ColunaC3 = 17
            Case Planilha2.Range("B20").Text
            ColunaC3 = 18
            Case Planilha2.Range("B21").Text
            ColunaC3 = 19
            Case Planilha2.Range("B22").Text
            ColunaC3 = 20
            Case Planilha2.Range("B23").Text
            ColunaC3 = 21
            Case Planilha2.Range("B24").Text
            ColunaC3 = 22
            Case Planilha2.Range("B25").Text
            ColunaC3 = 23
            Case Planilha2.Range("B26").Text
            ColunaC3 = 24
            Case Planilha2.Range("B27").Text
            ColunaC3 = 25
            Case Planilha2.Range("B28").Text
            ColunaC3 = 26
            Case Planilha2.Range("B29").Text
            ColunaC3 = 27
            Case Planilha2.Range("B30").Text
            ColunaC3 = 28
            Case Planilha2.Range("B31").Text
            ColunaC3 = 29
            Case Planilha2.Range("B32").Text
            ColunaC3 = 30
            Case Planilha2.Range("B33").Text
            ColunaC3 = 31
            Case Planilha2.Range("B34").Text
            ColunaC3 = 32
            Case Planilha2.Range("B35").Text
            ColunaC3 = 33
            Case Planilha2.Range("B36").Text
            ColunaC3 = 34
            Case Planilha2.Range("B37").Text
            ColunaC3 = 35
            Case Planilha2.Range("B38").Text
            ColunaC3 = 36
            Case Planilha2.Range("B39").Text
            ColunaC3 = 37
            Case Planilha2.Range("B40").Text
            ColunaC3 = 38
            Case Planilha2.Range("B41").Text
            ColunaC3 = 39
            Case Planilha2.Range("B42").Text
            ColunaC3 = 40
        End Select
        
        If ComboBox3.Value = "Filtro 3" Then
            ColunaC3 = 1
        End If
            
        With Sheets(plan)
        ultimacelula = Cells(ActiveSheet.UsedRange.Rows.Count, 1).Row
        Do While linha <= ultimacelula
        
        If UCase(Cells(linha, ColunaC1)) Like UCase("*" & Trim(TextBox1) & "*") And _
            UCase(Cells(linha, ColunaC2)) Like UCase("*" & Trim(TextBox2) & "*") And _
            UCase(Cells(linha, ColunaC3)) Like UCase("*" & Trim(TextBox3) & "*") Then
        
        Set Item = LV1.ListItems.Add(Text:=.Cells(linha, 1))
            Item.SubItems(1) = .Cells(linha, 2)
            Item.SubItems(2) = .Cells(linha, 3)
            Item.SubItems(3) = .Cells(linha, 4)
            Item.SubItems(4) = .Cells(linha, 5)
            Item.SubItems(5) = .Cells(linha, 6)
            Item.SubItems(6) = .Cells(linha, 7)
            Item.SubItems(7) = .Cells(linha, 8)
            Item.SubItems(8) = .Cells(linha, 9)
            Item.SubItems(9) = .Cells(linha, 10)
            Item.SubItems(10) = .Cells(linha, 11)
            Item.SubItems(11) = .Cells(linha, 12)
            Item.SubItems(12) = .Cells(linha, 13)
            Item.SubItems(13) = .Cells(linha, 14)
            Item.SubItems(14) = .Cells(linha, 15)
            Item.SubItems(15) = .Cells(linha, 16)
            Item.SubItems(16) = .Cells(linha, 17)
            Item.SubItems(17) = .Cells(linha, 18)
            Item.SubItems(18) = .Cells(linha, 19)
            Item.SubItems(19) = .Cells(linha, 20)
            Item.SubItems(20) = .Cells(linha, 21)
            Item.SubItems(21) = .Cells(linha, 22)
            Item.SubItems(22) = .Cells(linha, 23)
            Item.SubItems(23) = .Cells(linha, 24)
            Item.SubItems(24) = .Cells(linha, 25)
            Item.SubItems(25) = .Cells(linha, 26)
            Item.SubItems(26) = .Cells(linha, 27)
            Item.SubItems(27) = .Cells(linha, 28)
            Item.SubItems(28) = .Cells(linha, 29)
            Item.SubItems(29) = .Cells(linha, 30)
            Item.SubItems(30) = .Cells(linha, 31)
            Item.SubItems(31) = .Cells(linha, 32)
            Item.SubItems(32) = .Cells(linha, 33)
            Item.SubItems(33) = .Cells(linha, 34)
            Item.SubItems(34) = .Cells(linha, 35)
            Item.SubItems(35) = .Cells(linha, 36)
            Item.SubItems(36) = .Cells(linha, 37)
            Item.SubItems(37) = .Cells(linha, 38)
            Item.SubItems(38) = .Cells(linha, 39)
            Item.SubItems(39) = .Cells(linha, 40)
            
            
            nome = Item.SubItems(1)
            contato = Item.SubItems(3)
            With New MSForms.DataObject
             .SetText TB1.Text
             .PutInClipboard
            
            msgfinal = "Abraços, Dra. Vanessa Dias."
            URLnumero = "https://web.whatsapp.com/send?phone=55" & contato
            If Planilha2.Range("D4").Value < 12 Then
            msgPerso = "Bom dia " & nome & "! Tudo bem?"
            Else
            msgPerso = "Boa tarde " & nome & "! Tudo bem?"
            End If
                     
            'Abrir whatsapp na tela do numero
            cd.Get URLnumero
            '.Get URLnumero
            Application.Wait (Now + TimeValue("00:00:35")) 'Tempo para esperar o comando ser realizado
            If Not cd.IsElementPresent(FindBy.Css("div[title='Digite uma mensagem']")) Then
                ThisWorkbook.Sheets("LogMensagem").Activate
                ActiveSheet.Cells(linha, 1).Value = "Erro, " & Now & ", " & PaginaInicial.TextBox6.Value
                ThisWorkbook.Sheets("Dados").Select
                Cells(linha, 5) = "Erro"
            Else
                cd.FindElementByCss("div[title='Digite uma mensagem']").Click
                cd.SendKeys msgPerso
                Application.Wait (Now + TimeValue("00:00:02")) 'Tempo para esperar o comando ser realizado
            '4.2 Clicar no botão para enviar a nossa mensagem
                cd.FindElementByCss("span[data-icon='send']").Click
                Application.Wait (Now + TimeValue("00:00:02")) 'Tempo para esperar o comando ser realizado
            '4.1 Escrever a mensagem que vamos enviar
                cd.FindElementByCss("div[title='Digite uma mensagem']").Click
                Call SendKeys("^v", True)
                Application.Wait (Now + TimeValue("00:00:02")) 'Tempo para esperar o comando ser realizado
            '4.2 Clicar no botão para enviar a nossa mensagem
                cd.FindElementByCss("span[data-icon='send']").Click
                Application.Wait (Now + TimeValue("00:00:02")) 'Tempo para esperar o comando ser realizado
            '5. Enviar o nosso arquivo
            'Arquivo 1
            '5.1 Clicando no botão anexo
                cd.FindElementByCss("span[data-testid='clip']").Click
                Application.Wait (Now + TimeValue("00:00:03")) 'Tempo para esperar o comando ser realizado
                '5.2 Digitar o endereço do arquivo
                cd.FindElementByCss("input[type='file']").SendKeys arquivo1
                Application.Wait (Now + TimeValue("00:00:05")) 'Tempo para esperar o comando ser realizado
                '5.3 Clicar para poder enviar o nosso arquivo
                cd.FindElementByCss("span[data-icon='send']").Click
                Application.Wait (Now + TimeValue("00:00:05")) 'Tempo para esperar o comando ser realizado
              '6. Enviar mensagem "assinatura"
                cd.FindElementByCss("div[title='Digite uma mensagem']").Click
                cd.SendKeys msgfinal
                Application.Wait (Now + TimeValue("00:00:02")) 'Tempo para esperar o comando ser realizado
                '4.2 Clicar no botão para enviar a nossa mensagem
                cd.FindElementByCss("span[data-icon='send']").Click
                Application.Wait (Now + TimeValue("00:00:02")) 'Tempo para esperar o comando ser realizado
                
                ThisWorkbook.Sheets("LogMensagem").Activate
                ActiveSheet.Cells(linha, 1).Value = "Mensagem enviada, " & Now & ", " & PaginaInicial.TextBox6.Value
                ThisWorkbook.Sheets("Dados").Select
                Cells(linha, 5) = "Mensagem enviada"
            
            End If
            End With
        
        End If
         linha = linha + 1
        Loop
        End With
        End With
Exit Sub
Erro:
MsgBox "Erro!", vbCritical
End Sub


Private Sub imgSenha_click()
Dim us As Worksheet
Set us = ThisWorkbook.Sheets("Usuarios")
Dim selected_row As Long
selected_row = Application.WorksheetFunction.Match(PaginaInicial.TextBox6.Value, us.Range("A:A"), 0)


UserForm3.TextBox1.Value = PaginaInicial.TextBox6.Value
UserForm3.Show
End Sub
Private Sub LB1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.TextBox5.Value = Me.LB1.List(Me.LB1.ListIndex, 0)

UserForm5.TextBox1.Value = Me.LB1.List(Me.LB1.ListIndex, 0)
UserForm5.TextBox2.Value = Me.LB1.List(Me.LB1.ListIndex, 2)
UserForm5.ComboBox1.Value = Me.LB1.List(Me.LB1.ListIndex, 1)
UserForm5.Show


End Sub
Private Sub LB2_Click()
Me.TextBox7.Value = Me.LB2.List(Me.LB2.ListIndex, 0)
PaginaInicial.TextBox4.Value = Me.LB2.List(Me.LB2.ListIndex, 1)
End Sub
Private Sub LB2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.TextBox7.Value = Me.LB2.List(Me.LB2.ListIndex, 0)

PaginaInicial.TextBox4.Value = Me.LB2.List(Me.LB2.ListIndex, 1)
PaginaInicial.TB1.Value = Me.LB2.List(Me.LB2.ListIndex, 1)
MultiPage1.Value = 0

End Sub



Public Sub Filtro()

ThisWorkbook.Sheets("Dados").Select
plan = ThisWorkbook.Sheets("Dados").Name

Tabelas = Planilha2.Name


Dim lsi As ListSubItem

Dim r As Long
Dim C As Long

Dim linha As Long
linha = 2

LV1.ListItems.Clear

    Select Case ComboBox1
        Case Planilha2.Range("B3").Text
        ColunaC1 = 1
        Case Planilha2.Range("B4").Text
        ColunaC1 = 2
        Case Planilha2.Range("B5").Text
        ColunaC1 = 3
        Case Planilha2.Range("B6").Text
        ColunaC1 = 4
        Case Planilha2.Range("B7").Text
        ColunaC1 = 5
        Case Planilha2.Range("B8").Text
        ColunaC1 = 6
        Case Planilha2.Range("B9").Text
        ColunaC1 = 7
        Case Planilha2.Range("B10").Text
        ColunaC1 = 8
        Case Planilha2.Range("B11").Text
        ColunaC1 = 9
        Case Planilha2.Range("B12").Text
        ColunaC1 = 10
        Case Planilha2.Range("B13").Text
        ColunaC1 = 11
        Case Planilha2.Range("B14").Text
        ColunaC1 = 12
        Case Planilha2.Range("B15").Text
        ColunaC1 = 13
        Case Planilha2.Range("B16").Text
        ColunaC1 = 14
        Case Planilha2.Range("B17").Text
        ColunaC1 = 15
        Case Planilha2.Range("B18").Text
        ColunaC1 = 16
        Case Planilha2.Range("B19").Text
        ColunaC1 = 17
        Case Planilha2.Range("B20").Text
        ColunaC1 = 18
        Case Planilha2.Range("B21").Text
        ColunaC1 = 19
        Case Planilha2.Range("B22").Text
        ColunaC1 = 20
        Case Planilha2.Range("B23").Text
        ColunaC1 = 21
        Case Planilha2.Range("B24").Text
        ColunaC1 = 22
        Case Planilha2.Range("B25").Text
        ColunaC1 = 23
        Case Planilha2.Range("B26").Text
        ColunaC1 = 24
        Case Planilha2.Range("B27").Text
        ColunaC1 = 25
        Case Planilha2.Range("B28").Text
        ColunaC1 = 26
        Case Planilha2.Range("B29").Text
        ColunaC1 = 27
        Case Planilha2.Range("B30").Text
        ColunaC1 = 28
        Case Planilha2.Range("B31").Text
        ColunaC1 = 29
        Case Planilha2.Range("B32").Text
        ColunaC1 = 30
        Case Planilha2.Range("B33").Text
        ColunaC1 = 31
        Case Planilha2.Range("B34").Text
        ColunaC1 = 32
        Case Planilha2.Range("B35").Text
        ColunaC1 = 33
        Case Planilha2.Range("B36").Text
        ColunaC1 = 34
        Case Planilha2.Range("B37").Text
        ColunaC1 = 35
        Case Planilha2.Range("B38").Text
        ColunaC1 = 36
        Case Planilha2.Range("B39").Text
        ColunaC1 = 37
        Case Planilha2.Range("B40").Text
        ColunaC1 = 38
        Case Planilha2.Range("B41").Text
        ColunaC1 = 39
        Case Planilha2.Range("B42").Text
        ColunaC1 = 40
    End Select

    If ComboBox1.Value = "Filtro 1" Then
        ColunaC1 = 1
    End If

    Select Case ComboBox2
        Case Planilha2.Range("B3").Text
        ColunaC2 = 1
        Case Planilha2.Range("B4").Text
        ColunaC2 = 2
        Case Planilha2.Range("B5").Text
        ColunaC2 = 3
        Case Planilha2.Range("B6").Text
        ColunaC2 = 4
        Case Planilha2.Range("B7").Text
        ColunaC2 = 5
        Case Planilha2.Range("B8").Text
        ColunaC2 = 6
        Case Planilha2.Range("B9").Text
        ColunaC2 = 7
        Case Planilha2.Range("B10").Text
        ColunaC2 = 8
        Case Planilha2.Range("B11").Text
        ColunaC2 = 9
        Case Planilha2.Range("B12").Text
        ColunaC2 = 10
        Case Planilha2.Range("B13").Text
        ColunaC2 = 11
        Case Planilha2.Range("B14").Text
        ColunaC2 = 12
        Case Planilha2.Range("B15").Text
        ColunaC2 = 13
        Case Planilha2.Range("B16").Text
        ColunaC2 = 14
        Case Planilha2.Range("B17").Text
        ColunaC2 = 15
        Case Planilha2.Range("B18").Text
        ColunaC2 = 16
        Case Planilha2.Range("B19").Text
        ColunaC2 = 17
        Case Planilha2.Range("B20").Text
        ColunaC2 = 18
        Case Planilha2.Range("B21").Text
        ColunaC2 = 19
        Case Planilha2.Range("B22").Text
        ColunaC2 = 20
        Case Planilha2.Range("B23").Text
        ColunaC2 = 21
        Case Planilha2.Range("B24").Text
        ColunaC2 = 22
        Case Planilha2.Range("B25").Text
        ColunaC2 = 23
        Case Planilha2.Range("B26").Text
        ColunaC2 = 24
        Case Planilha2.Range("B27").Text
        ColunaC2 = 25
        Case Planilha2.Range("B28").Text
        ColunaC2 = 26
        Case Planilha2.Range("B29").Text
        ColunaC2 = 27
        Case Planilha2.Range("B30").Text
        ColunaC2 = 28
        Case Planilha2.Range("B31").Text
        ColunaC2 = 29
        Case Planilha2.Range("B32").Text
        ColunaC2 = 30
        Case Planilha2.Range("B33").Text
        ColunaC2 = 31
        Case Planilha2.Range("B34").Text
        ColunaC2 = 32
        Case Planilha2.Range("B35").Text
        ColunaC2 = 33
        Case Planilha2.Range("B36").Text
        ColunaC2 = 34
        Case Planilha2.Range("B37").Text
        ColunaC2 = 35
        Case Planilha2.Range("B38").Text
        ColunaC2 = 36
        Case Planilha2.Range("B39").Text
        ColunaC2 = 37
        Case Planilha2.Range("B40").Text
        ColunaC2 = 38
        Case Planilha2.Range("B41").Text
        ColunaC2 = 39
        Case Planilha2.Range("B42").Text
        ColunaC2 = 40
   End Select

    If ComboBox2.Value = "Filtro 2" Then
        ColunaC2 = 1
    End If

    Select Case ComboBox3
        Case Planilha2.Range("B3").Text
        ColunaC3 = 1
        Case Planilha2.Range("B4").Text
        ColunaC3 = 2
        Case Planilha2.Range("B5").Text
        ColunaC3 = 3
        Case Planilha2.Range("B6").Text
        ColunaC3 = 4
        Case Planilha2.Range("B7").Text
        ColunaC3 = 5
        Case Planilha2.Range("B8").Text
        ColunaC3 = 6
        Case Planilha2.Range("B9").Text
        ColunaC3 = 7
        Case Planilha2.Range("B10").Text
        ColunaC3 = 8
        Case Planilha2.Range("B11").Text
        ColunaC3 = 9
        Case Planilha2.Range("B12").Text
        ColunaC3 = 10
        Case Planilha2.Range("B13").Text
        ColunaC3 = 11
        Case Planilha2.Range("B14").Text
        ColunaC3 = 12
        Case Planilha2.Range("B15").Text
        ColunaC3 = 13
        Case Planilha2.Range("B16").Text
        ColunaC3 = 14
        Case Planilha2.Range("B17").Text
        ColunaC3 = 15
        Case Planilha2.Range("B18").Text
        ColunaC3 = 16
        Case Planilha2.Range("B19").Text
        ColunaC3 = 17
        Case Planilha2.Range("B20").Text
        ColunaC3 = 18
        Case Planilha2.Range("B21").Text
        ColunaC3 = 19
        Case Planilha2.Range("B22").Text
        ColunaC3 = 20
        Case Planilha2.Range("B23").Text
        ColunaC3 = 21
        Case Planilha2.Range("B24").Text
        ColunaC3 = 22
        Case Planilha2.Range("B25").Text
        ColunaC3 = 23
        Case Planilha2.Range("B26").Text
        ColunaC3 = 24
        Case Planilha2.Range("B27").Text
        ColunaC3 = 25
        Case Planilha2.Range("B28").Text
        ColunaC3 = 26
        Case Planilha2.Range("B29").Text
        ColunaC3 = 27
        Case Planilha2.Range("B30").Text
        ColunaC3 = 28
        Case Planilha2.Range("B31").Text
        ColunaC3 = 29
        Case Planilha2.Range("B32").Text
        ColunaC3 = 30
        Case Planilha2.Range("B33").Text
        ColunaC3 = 31
        Case Planilha2.Range("B34").Text
        ColunaC3 = 32
        Case Planilha2.Range("B35").Text
        ColunaC3 = 33
        Case Planilha2.Range("B36").Text
        ColunaC3 = 34
        Case Planilha2.Range("B37").Text
        ColunaC3 = 35
        Case Planilha2.Range("B38").Text
        ColunaC3 = 36
        Case Planilha2.Range("B39").Text
        ColunaC3 = 37
        Case Planilha2.Range("B40").Text
        ColunaC3 = 38
        Case Planilha2.Range("B41").Text
        ColunaC3 = 39
        Case Planilha2.Range("B42").Text
        ColunaC3 = 40
    End Select

    If ComboBox3.Value = "Filtro 3" Then
        ColunaC3 = 1
    End If
    
    With Sheets(plan)
    ultimacelula = Cells(ActiveSheet.UsedRange.Rows.Count, 1).Row
    Do While linha <= ultimacelula
    
    If UCase(Cells(linha, ColunaC1)) Like UCase("*" & Trim(TextBox1) & "*") And _
        UCase(Cells(linha, ColunaC2)) Like UCase("*" & Trim(TextBox2) & "*") And _
        UCase(Cells(linha, ColunaC3)) Like UCase("*" & Trim(TextBox3) & "*") Then
    
    Set Item = LV1.ListItems.Add(Text:=.Cells(linha, 1))
        Item.SubItems(1) = .Cells(linha, 2)
        Item.SubItems(2) = .Cells(linha, 3)
        Item.SubItems(3) = .Cells(linha, 4)
        Item.SubItems(4) = .Cells(linha, 5)
        Item.SubItems(5) = .Cells(linha, 6)
        Item.SubItems(6) = .Cells(linha, 7)
        Item.SubItems(7) = .Cells(linha, 8)
        Item.SubItems(8) = .Cells(linha, 9)
        Item.SubItems(9) = .Cells(linha, 10)
        Item.SubItems(10) = .Cells(linha, 11)
        Item.SubItems(11) = .Cells(linha, 12)
        Item.SubItems(12) = .Cells(linha, 13)
        Item.SubItems(13) = .Cells(linha, 14)
        Item.SubItems(14) = .Cells(linha, 15)
        Item.SubItems(15) = .Cells(linha, 16)
        Item.SubItems(16) = .Cells(linha, 17)
        Item.SubItems(17) = .Cells(linha, 18)
        Item.SubItems(18) = .Cells(linha, 19)
        Item.SubItems(19) = .Cells(linha, 20)
        Item.SubItems(20) = .Cells(linha, 21)
        Item.SubItems(21) = .Cells(linha, 22)
        Item.SubItems(22) = .Cells(linha, 23)
        Item.SubItems(23) = .Cells(linha, 24)
        Item.SubItems(24) = .Cells(linha, 25)
        Item.SubItems(25) = .Cells(linha, 26)
        Item.SubItems(26) = .Cells(linha, 27)
        Item.SubItems(27) = .Cells(linha, 28)
        Item.SubItems(28) = .Cells(linha, 29)
        Item.SubItems(29) = .Cells(linha, 30)
        Item.SubItems(30) = .Cells(linha, 31)
        Item.SubItems(31) = .Cells(linha, 32)
        Item.SubItems(32) = .Cells(linha, 33)
        Item.SubItems(33) = .Cells(linha, 34)
        Item.SubItems(34) = .Cells(linha, 35)
        Item.SubItems(35) = .Cells(linha, 36)
        Item.SubItems(36) = .Cells(linha, 37)
        Item.SubItems(37) = .Cells(linha, 38)
        Item.SubItems(38) = .Cells(linha, 39)
        Item.SubItems(39) = .Cells(linha, 40)
    End If
    linha = linha + 1
    
    Loop
    
    End With
    
End Sub

Private Sub TextBox1_Change()
Call Filtro
End Sub

Private Sub TextBox2_Change()
Call Filtro
End Sub

Private Sub TextBox3_Change()
Call Filtro
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call efeitos_icones

End Sub

'Backup
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ThisWorkbook.Save

    ThisWorkbook.SaveCopyAs ("C:\Backup-Advocacia\backup.xlsm")
    Application.Quit
  
End Sub
Private Sub imgCriarCliente_click()
CadastroCliente.Show
End Sub




'                                   -----------------------------
'                                   | Efeitos visuais de botões |
'                                   -----------------------------

        Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call efeitos_icones
        End Sub
        Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call efeitos_icones
        End Sub
        Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call efeitos_icones
        End Sub
        Private Sub Frame4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call efeitos_icones
        End Sub
        Private Sub Frame5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call efeitos_icones
        End Sub
        Private Sub imgExecutar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgExecutar.BorderStyle = fmBorderStyleSingle
        imgExecutar.SpecialEffect = fmSpecialEffectSunken
        imgExecutar.ControlTipText = "Disparar mensagens"
        End Sub
        Private Sub imgAtualizar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgAtualizar.BorderStyle = fmBorderStyleSingle
        imgAtualizar.SpecialEffect = fmSpecialEffectSunken
        imgAtualizar.ControlTipText = "Atualizar dados"
        End Sub
        Private Sub imgExcel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgExcel.BorderStyle = fmBorderStyleSingle
        imgExcel.SpecialEffect = fmSpecialEffectSunken
        imgExcel.ControlTipText = "Abrir Excel"
        End Sub
        Private Sub imgSalvarMsg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgSalvarMsg.BorderStyle = fmBorderStyleSingle
        imgSalvarMsg.SpecialEffect = fmSpecialEffectSunken
        imgSalvarMsg.ControlTipText = "Salvar mensagem"
        End Sub
        Private Sub imgSelecionar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgSelecionar.BorderStyle = fmBorderStyleSingle
        imgSelecionar.SpecialEffect = fmSpecialEffectSunken
        imgSelecionar.ControlTipText = "Selecionar mensagem"
        End Sub
        Private Sub imgLimparMsg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgLimparMsg.BorderStyle = fmBorderStyleSingle
        imgLimparMsg.SpecialEffect = fmSpecialEffectSunken
        imgLimparMsg.ControlTipText = "Limpar mensagem"
        End Sub
        Private Sub imgInicio_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgInicio.BorderStyle = fmBorderStyleSingle
        imgInicio.SpecialEffect = fmSpecialEffectSunken
        imgInicio.ControlTipText = "Página Inicial"
        End Sub
        Private Sub imgEmotes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgEmotes.BorderStyle = fmBorderStyleSingle
        imgEmotes.SpecialEffect = fmSpecialEffectSunken
        imgEmotes.ControlTipText = "Emotes"
        End Sub
        Private Sub imgConfiguracao_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgConfiguracao.BorderStyle = fmBorderStyleSingle
        imgConfiguracao.SpecialEffect = fmSpecialEffectSunken
        imgConfiguracao.ControlTipText = "Configurações do sistema"
        End Sub
        Private Sub imgEditarMsg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgEditarMsg.BorderStyle = fmBorderStyleSingle
        imgEditarMsg.SpecialEffect = fmSpecialEffectSunken
        imgEditarMsg.ControlTipText = "Salvar edição"
        End Sub
        Private Sub imgEditarUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgEditarUser.BorderStyle = fmBorderStyleSingle
        imgEditarUser.SpecialEffect = fmSpecialEffectSunken
        imgEditarUser.ControlTipText = "Editar usuário"
        End Sub
        Private Sub imgExcluirMsg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgExcluirMsg.BorderStyle = fmBorderStyleSingle
        imgExcluirMsg.SpecialEffect = fmSpecialEffectSunken
        imgExcluirMsg.ControlTipText = "Excluir mensagem"
        End Sub
        Private Sub imgLimparImg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgLimparImg.BorderStyle = fmBorderStyleSingle
        imgLimparImg.SpecialEffect = fmSpecialEffectSunken
        imgLimparImg.ControlTipText = "Limpar Imagem"
        End Sub
        Private Sub imgNovaMsg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgNovaMsg.BorderStyle = fmBorderStyleSingle
        imgNovaMsg.SpecialEffect = fmSpecialEffectSunken
        imgNovaMsg.ControlTipText = "Nova Mensagem"
        End Sub
        Private Sub imgExcluirUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgExcluirUser.BorderStyle = fmBorderStyleSingle
        imgExcluirUser.SpecialEffect = fmSpecialEffectSunken
        imgExcluirUser.ControlTipText = "Excluir usuário"
        End Sub
        Private Sub imgSelecionarImg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgSelecionarImg.BorderStyle = fmBorderStyleSingle
        imgSelecionarImg.SpecialEffect = fmSpecialEffectSunken
        imgSelecionarImg.ControlTipText = "Selecionar Imagem"
        End Sub
        Private Sub imgSenha_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgSenha.BorderStyle = fmBorderStyleSingle
        imgSenha.SpecialEffect = fmSpecialEffectSunken
        imgSenha.ControlTipText = "Alterar Senha"
        End Sub
        Private Sub imgFluxoCaixa_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgFluxoCaixa.BorderStyle = fmBorderStyleSingle
        imgFluxoCaixa.SpecialEffect = fmSpecialEffectSunken
        imgFluxoCaixa.ControlTipText = "Fluxo de caixa"
        End Sub
        Private Sub imgMensagens_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgMensagens.BorderStyle = fmBorderStyleSingle
        imgMensagens.SpecialEffect = fmSpecialEffectSunken
        imgMensagens.ControlTipText = "Banco de Mensagens"
        End Sub
        Private Sub imgNovoUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgNovoUser.BorderStyle = fmBorderStyleSingle
        imgNovoUser.SpecialEffect = fmSpecialEffectSunken
        imgNovoUser.ControlTipText = "Novo usuário"
        End Sub
        Private Sub imgSalvarImg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgSalvarImg.BorderStyle = fmBorderStyleSingle
        imgSalvarImg.SpecialEffect = fmSpecialEffectSunken
        imgSalvarImg.ControlTipText = "Salvar Imagem"
        End Sub
        Private Sub imgCriarCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgCriarCliente.BorderStyle = fmBorderStyleSingle
        imgCriarCliente.SpecialEffect = fmSpecialEffectSunken
        imgCriarCliente.ControlTipText = "Criar cliente"
        End Sub
        Private Sub imgExcluirCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgExcluirCliente.BorderStyle = fmBorderStyleSingle
        imgExcluirCliente.SpecialEffect = fmSpecialEffectSunken
        imgExcluirCliente.ControlTipText = "Excluir cliente"
        End Sub
        Private Sub imgSalvar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgSalvar.BorderStyle = fmBorderStyleSingle
        imgSalvar.SpecialEffect = fmSpecialEffectSunken
        imgSalvar.ControlTipText = "Salvar dados"
        End Sub
        Private Sub imgUsuarios_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgUsuarios.BorderStyle = fmBorderStyleSingle
        imgUsuarios.SpecialEffect = fmSpecialEffectSunken
        imgUsuarios.ControlTipText = "Usuários"
        End Sub
        Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call efeitos_icones
        End Sub
        Private Sub imgEditarCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        imgEditarCliente.BorderStyle = fmBorderStyleSingle
        imgEditarCliente.SpecialEffect = fmSpecialEffectSunken
        imgEditarCliente.ControlTipText = "Editar cliente"
        End Sub
        
        
'                                   -----------------------------
'                                   |      Funções de click     |
'                                   -----------------------------

        Private Sub imgAtualizar_click()
        Call Filtro
        End Sub
        Private Sub imgSelecionar_click()
        MultiPage1.Value = 1
        End Sub
        Private Sub imgSalvarMsg_click()
        ThisWorkbook.Sheets("Mensagens").Select
        ActiveSheet.Range("C3").Value = TB1.Text
        End Sub
        Private Sub imgExcel_click()
        Application.Visible = True
        Me.Hide
        End Sub
        Private Sub imgLimparMsg_click()
        ThisWorkbook.Sheets("Mensagens").Select
        TB1.Text = ""
        End Sub
        Private Sub imgInicio_click()
        MultiPage1.Value = 0
        End Sub
        Private Sub abrir_excel_Click()
        Application.Visible = True
        Me.Hide
        End Sub
        Private Sub imgSalvar_click()
        ThisWorkbook.Save
        ThisWorkbook.SaveCopyAs ("C:\Backup-Advocacia\backup.xlsm")
        End Sub
        Private Sub botaoSalvarPlanilha_Click()
        ThisWorkbook.Save
        ThisWorkbook.SaveCopyAs ("C:\Backup-Advocacia\backup.xlsm")
        End Sub
        Private Sub CommandButton10_Click()
        CadastroCliente.Show
        End Sub
        Private Sub CommandButton5_Click()
        ThisWorkbook.Sheets("Mensagens").Select
        ActiveSheet.Range("C3").Value = TB1.Text
        End Sub
        Private Sub CommandButton7_Click()
        ThisWorkbook.Sheets("Mensagens").Select
        TB1.Text = ""
        End Sub
        Private Sub imgConfiguracao_click()
        MsgBox "Em breve"
        End Sub
        Private Sub imgUsuarios_click()
        MultiPage1.Value = 3
        End Sub
        Private Sub LB1_Click()
        Me.TextBox5.Value = Me.LB1.List(Me.LB1.ListIndex, 0)
        End Sub
        Private Sub imgLimparImg_click()
        MsgImagem.Picture = Nothing
        MultiPage1.Value = 1
        MultiPage1.Value = 0
        End Sub
        Private Sub imgMensagens_click()
        MultiPage1.Value = 1
        End Sub
        Private Sub imgNovaMsg_click()
        UserForm7.Show
        End Sub
        Private Sub imgNovoUser_click()
        UserForm4.Show
        End Sub
        Private Sub LV1_ItemClick(ByVal Item As MSComctlLib.ListItem)
        TextBox8.Value = LV1.SelectedItem.SubItems(1)
        End Sub


'                                   -----------------------------
'                                   |      Funções e ações      |
'                                   -----------------------------

        ' Editar Mensagem
        Private Sub imgEditarMsg_click()
            If Me.TextBox7.Value = "" Then
            MsgBox "Selecione a mensagem para excluir."
            Exit Sub
            End If
        Dim msgs As Worksheet
        Set msgs = ThisWorkbook.Sheets("Mensagens")
        Dim selected_row As Long
        selected_row = Application.WorksheetFunction.Match(PaginaInicial.TextBox7.Value, msgs.Range("A:A"), 0)
        msgs.Range("B" & selected_row).Value = Me.TextBox4.Value
        Call atualiza_mensagem
        End Sub
        ' Excluir mensagem
        Private Sub imgExcluirMsg_click()
            If Me.TextBox7.Value = "" Then
            MsgBox "Selecione a mensagem para excluir."
            Exit Sub
            End If
        Dim msgs As Worksheet
        Set msgs = ThisWorkbook.Sheets("Mensagens")
        Dim selected_row As Long
        selected_row = Application.WorksheetFunction.Match(PaginaInicial.TextBox7.Value, msgs.Range("A:A"), 0)
        msgs.Range("A" & selected_row).EntireRow.Delete
        Call atualiza_mensagem
        End Sub
        ' Editar Usuario
        Private Sub imgEditarUser_click()
            If Me.TextBox5.Value = "" Then
            MsgBox "Selecione o usuário para editar."
            Exit Sub
            End If
        UserForm5.TextBox1.Value = Me.LB1.List(Me.LB1.ListIndex, 0)
        UserForm5.TextBox2.Value = Me.LB1.List(Me.LB1.ListIndex, 2)
        UserForm5.ComboBox1.Value = Me.LB1.List(Me.LB1.ListIndex, 1)
        UserForm5.Show
        End Sub
        ' Excluir Usuário
        Private Sub imgExcluirUser_click()
            If Me.TextBox5.Value = "" Then
            MsgBox "Selecione o usuário para excluir."
            Exit Sub
            End If
        Dim us As Worksheet
        Set us = ThisWorkbook.Sheets("Usuarios")
        Dim selected_row As Long
        selected_row = Application.WorksheetFunction.Match(PaginaInicial.TextBox5.Value, us.Range("A:A"), 0)
        us.Range("A" & selected_row).EntireRow.Delete
        Call atualiza_usuarios
        End Sub
        ' Selecionar imagem
        Private Sub imgSelecionarImg_click()
        On Error GoTo Erro
        Dim imagem As String
        Dim img As String
        ThisWorkbook.Sheets("Mensagens").Range("H2") = Application.GetOpenFilename("Imagens JPG,*.jpg,", , "Selecione uma imagem JPG:")
        img = ThisWorkbook.Sheets("Mensagens").Range("H2")
            If img = "Falso" Then
            Exit Sub
            Else
            MsgImagem.Picture = LoadPicture(img)
            End If
        MultiPage1.Value = 1
        MultiPage1.Value = 0
        Exit Sub
Erro:
        MsgBox "Erro!", vbCritical
        End Sub
        ' Selecionar imagem
        Private Sub MsgImagem_click()
        On Error GoTo Erro
        Dim imagem As String
        Dim img As String
        ThisWorkbook.Sheets("Mensagens").Range("H2") = Application.GetOpenFilename("Imagens JPG,*.jpg,", , "Selecione uma imagem JPG:")
        img = ThisWorkbook.Sheets("Mensagens").Range("H2")
            If img = "Falso" Then
            Exit Sub
            Else
            MsgImagem.Picture = LoadPicture(img)
            End If
        MultiPage1.Value = 1
        MultiPage1.Value = 0
        Exit Sub
Erro:
        MsgBox "Erro!", vbCritical
        End Sub
        ' Atualiza mensagem
        Sub atualiza_mensagem()
        Dim msgs As Worksheet
        Set msgs = ThisWorkbook.Sheets("Mensagens")
        Dim last_row As Long
        last_row = Application.WorksheetFunction.CountA(msgs.Range("A:A"))
        With Me.LB2
            .ColumnHeads = True
            .ColumnCount = 2
            .ColumnWidths = "80,80"
            If last_row = 1 Then
            .RowSource = "Mensagens!A2:C2"
            Else
            .RowSource = "Mensagens!A2:C" & last_row
            End If
        End With
        End Sub
        ' Atualiza clientes
        Sub atualiza_usuarios()
        Dim us As Worksheet
        Set us = ThisWorkbook.Sheets("Usuarios")
        Dim last_row As Long
        last_row = Application.WorksheetFunction.CountA(us.Range("A:A"))
        With Me.LB1
            .ColumnHeads = True
            .ColumnCount = 2
            .ColumnWidths = "120,40"
            
            If last_row = 1 Then
            .RowSource = "Usuarios!A2:C2"
            Else
            .RowSource = "Usuarios!A2:C" & last_row
            End If
        End With
        End Sub
        ' Excluir Cliente
        Private Sub imgExcluirCliente_click()
            If PaginaInicial.TextBox8.Value = "" Then
            MsgBox "Selecione o usuário para editar."
            Exit Sub
            End If
        Dim dds As Worksheet
        Set dds = ThisWorkbook.Sheets("Dados")
        Dim selected_row As Long
        selected_row = Application.WorksheetFunction.Match(PaginaInicial.TextBox8.Value, dds.Range("B:B"), 0)
        'Confirmação de exclusão
        Dim Msg, STYLE, Title, Response
        Msg = "Tem certeza que deseja excluir essa linha?"
        STYLE = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Confirmação de exclusão"
        Response = MsgBox(Msg, STYLE, Title)
            If Response = vbYes Then
                MsgBox "Cliente excluído com sucesso"
                dds.Range("B" & selected_row).EntireRow.Delete
            Else
                MsgBox "Operação cancelada"
            End If
        'Atualiza a lista
        Call PaginaInicial.Filtro
        End Sub