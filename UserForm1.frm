VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Cadastro"
   ClientHeight    =   9360.001
   ClientLeft      =   -705
   ClientTop       =   -2625
   ClientWidth     =   15225
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ultimodado As Long
Function AbrirArquivoDados() As Workbook
    Dim wb As Workbook
    Dim filePath As String
    filePath = "C:\temp\bancoDados.xlsx" ' Atualize com o caminho correto
    
    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    
    If wb Is Nothing Then
        MsgBox "Erro ao abrir o arquivo de dados."
        Set AbrirArquivoDados = Nothing
    Else
        Set AbrirArquivoDados = wb
    End If
End Function

Sub CriarSaldo(ID As String, Nome As String, Saldo As Double)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    
    Set wb = AbrirArquivoDados()
    
    If Not wb Is Nothing Then
        Set ws = wb.Sheets("Saldos")
        ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        
        ws.Cells(ultimaLinha, 1).Value = ID
        ws.Cells(ultimaLinha, 2).Value = Nome
        ws.Cells(ultimaLinha, 3).Value = Saldo
        
        wb.Save
        wb.Close
        MsgBox "Registro criado com sucesso!"
    End If
End Sub
Private Sub BTAlterarRegistro_Click()

    Dim resposta As VbMsgBoxResult
    Dim valor As Long
    Dim fila As Range
    Dim linha As Long
    Dim ws As Worksheet
    
    ' Verifica se o ID foi informado
    If Me.TextBoxIDP.Value = "" Then
        MsgBox "Selecione um Cadastro para alterar!"
        Exit Sub
    End If
    
    ' Verifica se o ID inserido é um número
    If Not IsNumeric(Me.TextBoxIDP.Value) Then
        MsgBox "O ID deve ser um número!"
        Exit Sub
    End If
    
    valor = CLng(Me.TextBoxIDP.Value)
    
    ' Confirma se o usuário deseja alterar o cadastro
    resposta = MsgBox("Deseja alterar o cadastro de ID " & valor & "?", vbYesNo)
    
    If resposta = vbNo Then
        Exit Sub
    Else
        ' Define a planilha a ser usada
        Set ws = Sheets("DADOS")
        
        ' Procura o valor na coluna A
        Set fila = ws.Range("A:A").Find(valor, LookAt:=xlWhole)
        
        ' Verifica se encontrou o valor
        If fila Is Nothing Then
            MsgBox "ID não encontrado!"
            Exit Sub
        End If
        
        linha = fila.Row
        
        ' Atualiza os dados na planilha
        ws.Range("B" & linha).Value = Me.TextBoxProdutor.Value
        ws.Range("C" & linha).Value = Me.TextBoxEmail.Value
        ws.Range("D" & linha).Value = Me.TextBoxFazenda.Value
        ws.Range("E" & linha).Value = Me.TextBoxCPF.Value
        ws.Range("F" & linha).Value = Me.TextBoxCNPJ.Value
        ws.Range("G" & linha).Value = Me.TextBoxTelefone.Value
        ws.Range("H" & linha).Value = Me.TextBoxCEP.Value
        ws.Range("I" & linha).Value = Me.TextBoxRua.Value
        ws.Range("J" & linha).Value = Me.TextBoxNumero.Value
        ws.Range("K" & linha).Value = Me.TextBoxRegiao.Value
        ws.Range("L" & linha).Value = Me.TextBoxCidade.Value
        ws.Range("M" & linha).Value = Me.TextBoxEstado.Value
        ws.Range("N" & linha).Value = Me.TextBoxBairro.Value
        ws.Range("O" & linha).Value = CDate(Format(Me.TextBoxAbertura.Text, "dd/mm/yyyy"))
        ws.Range("P" & linha).Value = CDate(Format(Me.TextBoxFechamento.Text, "dd/mm/yyyy"))
        
        ' Atualiza os dados da amostra
        ws.Range("Q" & linha).Value = Me.TextBoxIDAmostra.Value
        ws.Range("R" & linha).Value = Me.TextBoxAmostra.Value
        ws.Range("S" & linha).Value = Me.TextBoxEspecie.Value
        ws.Range("T" & linha).Value = Me.TextBoxLote.Value
        ws.Range("U" & linha).Value = Me.TextBoxVariedade.Value
        ws.Range("V" & linha).Value = Me.TextBoxPadrao.Value
        ws.Range("W" & linha).Value = Me.TextBoxCultivar.Value
        ws.Range("X" & linha).Value = Me.TextBoxProcesso.Value
        ws.Range("Y" & linha).Value = Me.TextBoxSeca.Value
        ws.Range("Z" & linha).Value = Me.TextBoxAltitude.Value
        ws.Range("AA" & linha).Value = Me.TextBoxAspecto.Value
        ws.Range("AB" & linha).Value = Me.TextBoxGleba.Value
        ws.Range("AC" & linha).Value = Me.TextBoxSacas.Value
        ws.Range("AD" & linha).Value = Me.TextBoxSafraAno.Value
        ws.Range("AE" & linha).Value = Me.TextBoxLoteCorrido.Value
        ws.Range("AF" & linha).Value = Me.TextBoxCafe.Value
        ws.Range("AG" & linha).Value = Me.TextBoxBicaCorrida.Value
        ws.Range("AH" & linha).Value = Me.TextBoxCor.Value
        ws.Range("AI" & linha).Value = Me.TextBoxPVA.Value
        ws.Range("AJ" & linha).Value = Me.TextBoxUmidade.Value
        ws.Range("AK" & linha).Value = Me.TextBoxCata.Value
        ws.Range("AL" & linha).Value = Me.TextBoxObservacoes.Value
        
        ' Limpa os campos de entrada
        Me.TextBoxProdutor.Value = ""
        Me.TextBoxEmail.Value = ""
        Me.TextBoxFazenda.Value = ""
        Me.TextBoxCPF.Value = ""
        Me.TextBoxCNPJ.Value = ""
        Me.TextBoxTelefone.Value = ""
        Me.TextBoxCEP.Value = ""
        Me.TextBoxRua.Value = ""
        Me.TextBoxNumero.Value = ""
        Me.TextBoxRegiao.Value = ""
        Me.TextBoxCidade.Value = ""
        Me.TextBoxEstado.Value = ""
        Me.TextBoxBairro.Value = ""
        Me.TextBoxAbertura.Value = ""
        Me.TextBoxFechamento.Value = ""
        Me.TextBoxAmostra.Value = ""
        Me.TextBoxEspecie.Value = ""
        Me.TextBoxLote.Value = ""
        Me.TextBoxVariedade.Value = ""
        Me.TextBoxPadrao.Value = ""
        Me.TextBoxCultivar.Value = ""
        Me.TextBoxProcesso.Value = ""
        Me.TextBoxSeca.Value = ""
        Me.TextBoxAltitude.Value = ""
        Me.TextBoxAspecto.Value = ""
        Me.TextBoxGleba.Value = ""
        Me.TextBoxSacas.Value = ""
        Me.TextBoxSafraAno.Value = ""
        Me.TextBoxLoteCorrido.Value = ""
        Me.TextBoxCafe.Value = ""
        Me.TextBoxBicaCorrida.Value = ""
        Me.TextBoxCor.Value = ""
        Me.TextBoxPVA.Value = ""
        Me.TextBoxUmidade.Value = ""
        Me.TextBoxCata.Value = ""
        Me.TextBoxObservacoes.Value = ""
        
        ' Reseta o ListBox
        ListBox_2.ListIndex = -1
        
        ' Mensagem de confirmação
        MsgBox "Cadastro Alterado com sucesso!"
    End If

End Sub

Private Sub BTBuscaCEP_Click()

    Dim CEP As String
    Dim Rua As String
    Dim Bairro As String
    Dim Uf As String
    Dim Complemento As String
    Dim Cidade As String
    
    ' Verifica se o campo de CEP não está vazio
    If Me.TextBoxCEP.Value = "" Then
        MsgBox "Por favor, insira um CEP válido.", vbExclamation
        Exit Sub
    End If
    
    CEP = Me.TextBoxCEP.Value
    
    ' Chama a função de busca do CEP
    Call BuscaCEP(CEP, Rua, Bairro, Uf, Complemento, Cidade)
    
    ' Verifica se a busca foi bem-sucedida (se os valores foram preenchidos)
    If Rua = "" And Bairro = "" And Uf = "" And Cidade = "" Then
        MsgBox "CEP não encontrado. Verifique o CEP informado.", vbExclamation
        Exit Sub
    End If
    
    ' Atualiza os campos do formulário com os valores obtidos
    Me.TextBoxRua.Value = Rua
    Me.TextBoxBairro.Value = Bairro
    Me.TextBoxCidade.Value = Cidade
    Me.TextBoxEstado.Value = Uf

End Sub


    Private Sub BTDeletarl_Click()

    Dim linha As Range
    Dim resposta As VbMsgBoxResult
    Dim valor As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim registroEncontrado As Boolean
    
    ' Verificar se o TextBoxIDP está vazio
    If Me.TextBoxIDP.Value = "" Then
        MsgBox "Selecione um Cadastro antes de prosseguir com a exclusão!"
        Exit Sub
    End If
    
    ' Obter o valor do TextBoxIDP
    valor = CLng(Me.TextBoxIDP.Value)
    resposta = MsgBox("Deseja excluir o cadastro de ID " & valor & "?", vbYesNo)
    
    If resposta = vbNo Then
        Exit Sub
    End If
    
    ' Definir a planilha
    Set ws = Sheets("DADOS")
    
    ' Verificar se a resposta foi sim
    If resposta = vbYes Then
        With Me.ListBox_2
            registroEncontrado = False ' Inicializa como não encontrado
            
            ' Percorrer os itens da ListBox de trás para frente
            For i = .ListCount - 1 To 0 Step -1
                If .Selected(i) Then
                    ' Procurar a linha correspondente na coluna A
                    Set linha = ws.Range("A:A").Find(.List(i, 0), LookAt:=xlWhole)
                    
                    ' Verificar se a linha foi encontrada
                    If Not linha Is Nothing Then
                        ' Excluir a linha
                        linha.EntireRow.Delete
                        registroEncontrado = True ' Define como encontrado
                        
                        ' Remover o item da ListBox
                        .RemoveItem i
                        
                        ListBox_2.ListIndex = -1
    
                        ' Limpar campos Produtor
                        Me.TextBoxBuscaCPF.Value = ""
                        Me.TextBoxProdutor.Value = ""
                        Me.TextBoxEmail.Value = ""
                        Me.TextBoxFazenda.Value = ""
                        Me.TextBoxCPF.Value = ""
                        Me.TextBoxCNPJ.Value = ""
                        Me.TextBoxTelefone.Value = ""
                        Me.TextBoxCEP.Value = ""
                        Me.TextBoxRua.Value = ""
                        Me.TextBoxNumero.Value = ""
                        Me.TextBoxRegiao.Value = ""
                        Me.TextBoxCidade.Value = ""
                        Me.TextBoxEstado.Value = ""
                        Me.TextBoxBairro.Value = ""
                        Me.TextBoxAbertura.Value = ""
                        Me.TextBoxFechamento.Value = ""
                        
                        ' Limpar campos Amostra
                        Me.TextBoxIDAmostra.Value = ""
                        Me.TextBoxAmostra.Value = ""
                        Me.TextBoxEspecie.Value = ""
                        Me.TextBoxLote.Value = ""
                        Me.TextBoxVariedade.Value = ""
                        Me.TextBoxPadrao.Value = ""
                        Me.TextBoxCultivar.Value = ""
                        Me.TextBoxProcesso.Value = ""
                        Me.TextBoxSeca.Value = ""
                        Me.TextBoxAltitude.Value = ""
                        Me.TextBoxAspecto.Value = ""
                        Me.TextBoxGleba.Value = ""
                        Me.TextBoxSacas.Value = ""
                        Me.TextBoxSafraAno.Value = ""
                        Me.TextBoxLoteCorrido.Value = ""
                        Me.TextBoxCafe.Value = ""
                        Me.TextBoxBicaCorrida.Value = ""
                        Me.TextBoxCor.Value = ""
                        Me.TextBoxPVA.Value = ""
                        Me.TextBoxUmidade.Value = ""
                        Me.TextBoxRenda.Value = ""
                        Me.TextBoxCata.Value = ""
                        Me.TextBoxObservacoes.Value = ""
                        
                        MsgBox "Registro excluído com sucesso!"
                        Exit For ' Sai do loop após excluir o registro
                    End If
                End If
            Next i
            
            ' Exibe uma mensagem caso o registro não tenha sido encontrado
            If Not registroEncontrado Then
                MsgBox "Registro não encontrado na planilha!"
            End If
        End With
    End If

End Sub

Private Sub BTEditarExcluir_Click()

    Me.BTAlterarRegistro.Enabled = True
    Me.BTDeletarl.Enabled = True
    Me.BTSalvarNovo.Enabled = False

End Sub

Private Sub BTInserirNovol_Click()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DADOS")

    Me.BTAlterarRegistro.Enabled = False
    Me.BTDeletarl.Enabled = False
    Me.BTSalvarNovo.Enabled = True
      
    
    ' Limpar campos de Produtor
    Me.TextBoxProdutor.Value = ""
    Me.TextBoxEmail.Value = ""
    Me.TextBoxFazenda.Value = ""
    Me.TextBoxCPF.Value = ""
    Me.TextBoxCNPJ.Value = ""
    Me.TextBoxTelefone.Value = ""
    Me.TextBoxCEP.Value = ""
    Me.TextBoxRua.Value = ""
    Me.TextBoxNumero.Value = ""
    Me.TextBoxRegiao.Value = ""
    Me.TextBoxCidade.Value = ""
    Me.TextBoxEstado.Value = ""
    Me.TextBoxBairro.Value = ""
    Me.TextBoxAbertura.Value = ""
    Me.TextBoxFechamento.Value = ""
    
    ' Limpar campos de Amostra
    Me.TextBoxAmostra.Value = ""
    Me.TextBoxEspecie.Value = ""
    Me.TextBoxLote.Value = ""
    Me.TextBoxVariedade.Value = ""
    Me.TextBoxPadrao.Value = ""
    Me.TextBoxCultivar.Value = ""
    Me.TextBoxProcesso.Value = ""
    Me.TextBoxSeca.Value = ""
    Me.TextBoxAltitude.Value = ""
    Me.TextBoxAspecto.Value = ""
    Me.TextBoxGleba.Value = ""
    Me.TextBoxSacas.Value = ""
    Me.TextBoxSafraAno.Value = ""
    Me.TextBoxLoteCorrido.Value = ""
    Me.TextBoxCafe.Value = ""
    Me.TextBoxBicaCorrida.Value = ""
    Me.TextBoxCor.Value = ""
    Me.TextBoxPVA.Value = ""
    Me.TextBoxUmidade.Value = ""
    Me.TextBoxRenda.Value = ""
    Me.TextBoxCata.Value = ""
    Me.TextBoxObservacoes.Value = ""
    
    ' Definir o próximo ID automaticamente com base no maior valor da coluna A
    
    Me.TextBoxIDP.Value = ws.Range("Z1").Value
    Me.TextBoxIDAmostra.Value = ws.Range("Z1").Value
    Me.TextBoxIDClassificacao.Value = ws.Range("Z1").Value
    

End Sub

Private Sub BTLimpar_Click()


    ListBox_2.ListIndex = -1
    
    Me.TextBoxBuscaCPF.Value = ""
    
    Me.TextBoxProdutor.Value = ""
    Me.TextBoxEmail.Value = ""
    Me.TextBoxFazenda.Value = ""
    Me.TextBoxCPF.Value = ""
    Me.TextBoxCNPJ.Value = ""
    Me.TextBoxTelefone.Value = ""
    Me.TextBoxCEP.Value = ""
    Me.TextBoxRua.Value = ""
    Me.TextBoxNumero.Value = ""
    Me.TextBoxRegiao.Value = ""
    Me.TextBoxCidade.Value = ""
    Me.TextBoxEstado.Value = ""
    Me.TextBoxBairro.Value = ""
    Me.TextBoxAbertura.Value = ""
    Me.TextBoxFechamento.Value = ""
    
    'Amostra
    
    Me.TextBoxAmostra.Value = ""
    Me.TextBoxEspecie.Value = ""
    Me.TextBoxLote.Value = ""
    Me.TextBoxVariedade.Value = ""
    Me.TextBoxPadrao.Value = ""
    Me.TextBoxCultivar.Value = ""
    Me.TextBoxProcesso.Value = ""
    Me.TextBoxSeca.Value = ""
    Me.TextBoxAltitude.Value = ""
    Me.TextBoxAspecto.Value = ""
    Me.TextBoxGleba.Value = ""
    Me.TextBoxSacas.Value = ""
    Me.TextBoxSafraAno.Value = ""
    Me.TextBoxLoteCorrido.Value = ""
    Me.TextBoxCafe.Value = ""
    Me.TextBoxBicaCorrida.Value = ""
    Me.TextBoxCor.Value = ""
    Me.TextBoxPVA.Value = ""
    Me.TextBoxUmidade.Value = ""
    Me.TextBoxRenda.Value = ""
    Me.TextBoxCata.Value = ""
    Me.TextBoxObservacoes.Value = ""
    Me.TextBoxDefeitos.Value = ""
    Me.ComboBoxDefeitos = Empty
    
    Me.ListBox_2.RowSource = "DADOS!Produtor"
    Me.TextBoxBuscaCPF.SetFocus


End Sub

Private Sub BTPesquisarCPF_Click()

If Me.TextBoxBuscaCPF.Value = "" Then
 
    MsgBox "Insira o CPF para realizar a Busca."
    Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DADOS")
    
    ultimodado = ws.Range("E" & ws.Rows.Count).End(xlUp).Row
    
    List = Clear
    ListBox_2.RowSource = Clear
    
    y = 0
    
    For linha = 4 To ultimodado
    
        cpf = ws.Cells(linha, 5).Value
        
        If cpf Like Me.TextBoxBuscaCPF.Value Then
        
        Me.ListBox_2.List = ws.Range("A3:AZ4").Value
        Me.ListBox_2.Clear
        
       
            'produtor
            Me.ListBox_2.AddItem
            Me.ListBox_2.List(y, 0) = ws.Cells(linha, 1).Value
            Me.ListBox_2.List(y, 1) = ws.Cells(linha, 2).Value
            Me.ListBox_2.List(y, 2) = ws.Cells(linha, 3).Value
            Me.ListBox_2.List(y, 3) = ws.Cells(linha, 4).Value
            Me.ListBox_2.List(y, 4) = ws.Cells(linha, 5).Value
            Me.ListBox_2.List(y, 5) = ws.Cells(linha, 6).Value
            Me.ListBox_2.List(y, 6) = ws.Cells(linha, 7).Value
            Me.ListBox_2.List(y, 7) = ws.Cells(linha, 8).Value
            Me.ListBox_2.List(y, 8) = ws.Cells(linha, 9).Value
            Me.ListBox_2.List(y, 9) = ws.Cells(linha, 10).Value
            Me.ListBox_2.List(y, 10) = ws.Cells(linha, 11).Value
            Me.ListBox_2.List(y, 11) = ws.Cells(linha, 12).Value
            Me.ListBox_2.List(y, 12) = ws.Cells(linha, 13).Value
            Me.ListBox_2.List(y, 13) = ws.Cells(linha, 14).Value
            Me.ListBox_2.List(y, 14) = ws.Cells(linha, 15).Value
            Me.ListBox_2.List(y, 15) = ws.Cells(linha, 16).Value
            
            'Amostra
            
            Me.ListBox_2.List(y, 16) = ws.Cells(linha, 17).Value
            Me.ListBox_2.List(y, 17) = ws.Cells(linha, 18).Value
            Me.ListBox_2.List(y, 18) = ws.Cells(linha, 19).Value
            Me.ListBox_2.List(y, 19) = ws.Cells(linha, 20).Value
            Me.ListBox_2.List(y, 20) = ws.Cells(linha, 21).Value
            Me.ListBox_2.List(y, 21) = ws.Cells(linha, 22).Value
            Me.ListBox_2.List(y, 22) = ws.Cells(linha, 23).Value
            Me.ListBox_2.List(y, 23) = ws.Cells(linha, 24).Value
            Me.ListBox_2.List(y, 24) = ws.Cells(linha, 25).Value
            Me.ListBox_2.List(y, 25) = ws.Cells(linha, 26).Value
            Me.ListBox_2.List(y, 26) = ws.Cells(linha, 27).Value
            Me.ListBox_2.List(y, 27) = ws.Cells(linha, 28).Value
            Me.ListBox_2.List(y, 28) = ws.Cells(linha, 29).Value
            Me.ListBox_2.List(y, 29) = ws.Cells(linha, 30).Value
            Me.ListBox_2.List(y, 30) = ws.Cells(linha, 31).Value
            Me.ListBox_2.List(y, 31) = ws.Cells(linha, 32).Value
            Me.ListBox_2.List(y, 32) = ws.Cells(linha, 33).Value
            Me.ListBox_2.List(y, 33) = ws.Cells(linha, 34).Value
            Me.ListBox_2.List(y, 34) = ws.Cells(linha, 35).Value
            Me.ListBox_2.List(y, 35) = ws.Cells(linha, 36).Value
            Me.ListBox_2.List(y, 36) = ws.Cells(linha, 37).Value
            Me.ListBox_2.List(y, 37) = ws.Cells(linha, 38).Value
            
           
            
            y = y + 1
            End If
            Next
            
    ListBox_2.ListIndex = -1
    Me.TextBoxBuscaCPF.SetFocusf


End Sub

Private Sub BTSalvarDefeitos_Click()

Dim ws As Worksheet
Dim linhaAmostra As Long
Dim defeito As String
Dim quantidade As Long
Dim colunasDefeitos As Variant
Dim i As Long
Dim codigo As Long
Dim colunaID As Range
' Dim resultadoBusca As Range

Set ws = ThisWorkbook.Sheets("DADOS")


codigo = Me.TextBoxIDP.Value

Set colunaID = ws.Columns("A").Find(What:=codigo, LookIn:=xlValues, LookAt:=xlWhole)

If Not colunaID Is Nothing Then

    linhaAmostra = colunaID.Row
Else
    MsgBox "ID não encontrado. Verifique o ID inserido.", vbExclamation, "Erro"
    Exit Sub
End If


colunasDefeitos = Array("Preto", "Preto Verde", "Ardido", "Verde", "Quebrado", "Mau Granado / Chocho", "Brocado Limpo", "Brocado Sujo", "Coco", "Marinheiro", "Casca Gd", _
"Casca Md ou Pq", "Fragmentos de Casca", "Pau/Pedra/Torrão Gd", "Pau/Pedra/Torrão Md", "Pau/Pedra/Torrão Pq")


defeito = Me.ComboBoxDefeitos.Value
quantidade = Me.TextBoxDefeitos.Value


For i = LBound(colunasDefeitos) To UBound(colunasDefeitos)
    If defeito = colunasDefeitos(i) Then
        
        ws.Cells(linhaAmostra, 40 + i).Value = quantidade
        Exit For
    End If
Next i

If i > UBound(colunasDefeitos) Then
    MsgBox "Defeito não reconhecido. Verifique o valor selecionado.", vbExclamation, "Erro"
    Exit Sub
End If

Me.ComboBoxDefeitos.Value = ""
Me.TextBoxQuantidade.Value = ""


Dim resposta As VbMsgBoxResult
resposta = MsgBox("Deseja adicionar outro defeito?", vbYesNo + vbQuestion, "Adicionar Outro Defeito")

If resposta = vbNo Then
    MsgBox "Todos os defeitos foram salvos com sucesso.", vbInformation, "Sucesso"
    Exit Sub
End If

Me.ListBox_2.RowSource = "DADOS!Produtor"

End Sub
Private Sub BTSalvarNovo_Click()


    If Me.TextBoxProdutor = "" Or Me.TextBoxEmail = "" Or Me.TextBoxCPF = "" Or Me.TextBoxTelefone = "" Or _
    Me.TextBoxFazenda = "" Or Me.TextBoxCNPJ = "" Or Me.TextBoxRua = "" Or Me.TextBoxNumero = "" Or _
    Me.TextBoxBairro = "" Or Me.TextBoxCEP = "" Or Me.TextBoxCidade = "" Or Me.TextBoxRegiao = "" Or _
    Me.TextBoxAbertura = "" Or Me.TextBoxFechamento = "" Or Me.TextBoxEstado = "" Or _
    Me.TextBoxAmostra.Value = "" Or Me.TextBoxEspecie.Value = "" Or Me.TextBoxLote.Value = "" Or Me.TextBoxVariedade.Value = "" Or _
    Me.TextBoxPadrao.Value = "" Or Me.TextBoxCultivar.Value = "" Or Me.TextBoxProcesso.Value = "" Or Me.TextBoxSeca.Value = "" Or _
    Me.TextBoxAltitude.Value = "" Or Me.TextBoxAspecto.Value = "" Or Me.TextBoxGleba.Value = "" Or Me.TextBoxSacas.Value = "" Or _
    Me.TextBoxSafraAno.Value = "" Or Me.TextBoxLoteCorrido.Value = "" Or Me.TextBoxCafe.Value = "" Or Me.TextBoxBicaCorrida.Value = "" Or _
    Me.TextBoxCor.Value = "" Or Me.TextBoxPVA.Value = "" Or Me.TextBoxUmidade.Value = "" Or Me.TextBoxRenda.Value = "" Or _
    Me.TextBoxCata.Value = "" Or Me.TextBoxObservacoes.Value = "" Then
    
    MsgBox "Todos os campos devem conter valores!"
    
    Exit Sub
    End If
    
    If Me.TextBoxIDP.Value <> Me.TextBoxIDAmostra.Value Then
    
        MsgBox "Os ID não correspondem!", vbExclamation
        
        Exit Sub
        
    End If

    

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DADOS")

    ws.Range("A4").EntireRow.Insert
    
    ' Produtor
    ws.Range("A4").Value = Me.TextBoxIDP.Value
    ws.Range("B4").Value = Me.TextBoxProdutor.Value
    ws.Range("C4").Value = Me.TextBoxEmail.Value
    ws.Range("D4").Value = Me.TextBoxFazenda.Value
    ws.Range("E4").Value = Me.TextBoxCPF.Value
    ws.Range("F4").Value = Me.TextBoxCNPJ.Value
    ws.Range("G4").Value = Me.TextBoxTelefone.Value
    ws.Range("H4").Value = Me.TextBoxCEP.Value
    ws.Range("I4").Value = Me.TextBoxRua.Value
    ws.Range("J4").Value = Me.TextBoxNumero.Value
    ws.Range("K4").Value = Me.TextBoxRegiao.Value
    ws.Range("L4").Value = Me.TextBoxCidade.Value
    ws.Range("M4").Value = Me.TextBoxEstado.Value
    ws.Range("N4").Value = Me.TextBoxBairro.Value
    ws.Range("O4").Value = CDate(Format(Me.TextBoxAbertura.Text, "dd/mm/yyyy"))
    ws.Range("P4").Value = CDate(Format(Me.TextBoxFechamento.Text, "dd/mm/yyyy"))
    
    ' Amostra
    
    ws.Range("A4").Value = Me.TextBoxIDAmostra.Value
    ws.Range("Q4").Value = Me.TextBoxAmostra.Value
    ws.Range("R4").Value = Me.TextBoxEspecie.Value
    ws.Range("S4").Value = Me.TextBoxLote.Value
    ws.Range("T4").Value = Me.TextBoxVariedade.Value
    ws.Range("U4").Value = Me.TextBoxPadrao.Value
    ws.Range("V4").Value = Me.TextBoxCultivar.Value
    ws.Range("W4").Value = Me.TextBoxProcesso.Value
    ws.Range("X4").Value = Me.TextBoxSeca.Value
    ws.Range("Y4").Value = Me.TextBoxAltitude.Value
    ws.Range("Z4").Value = Me.TextBoxAspecto.Value
    ws.Range("AA4").Value = Me.TextBoxGleba.Value
    ws.Range("AB4").Value = Me.TextBoxSacas.Value
    ws.Range("AC4").Value = Me.TextBoxSafraAno.Value
    ws.Range("AD4").Value = Me.TextBoxLoteCorrido.Value
    ws.Range("AE4").Value = Me.TextBoxCafe.Value
    ws.Range("AF4").Value = Me.TextBoxBicaCorrida.Value
    ws.Range("AG4").Value = Me.TextBoxCor.Value
    ws.Range("AH4").Value = Me.TextBoxPVA.Value
    ws.Range("AI4").Value = Me.TextBoxUmidade.Value
    ws.Range("AJ4").Value = Me.TextBoxRenda.Value
    ws.Range("AK4").Value = Me.TextBoxCata.Value
    ws.Range("AL4").Value = Me.TextBoxObservacoes.Value
    
    ' produtor
    
    Me.TextBoxProdutor.Value = ""
    Me.TextBoxEmail.Value = ""
    Me.TextBoxFazenda.Value = ""
    Me.TextBoxCPF.Value = ""
    Me.TextBoxCNPJ.Value = ""
    Me.TextBoxTelefone.Value = ""
    Me.TextBoxCEP.Value = ""
    Me.TextBoxRua.Value = ""
    Me.TextBoxNumero.Value = ""
    Me.TextBoxRegiao.Value = ""
    Me.TextBoxCidade.Value = ""
    Me.TextBoxEstado.Value = ""
    Me.TextBoxBairro.Value = ""
    Me.TextBoxAbertura.Value = ""
    Me.TextBoxFechamento.Value = ""
    
    'Amostra
    
    Me.TextBoxAmostra.Value = ""
    Me.TextBoxEspecie.Value = ""
    Me.TextBoxLote.Value = ""
    Me.TextBoxVariedade.Value = ""
    Me.TextBoxPadrao.Value = ""
    Me.TextBoxCultivar.Value = ""
    Me.TextBoxProcesso.Value = ""
    Me.TextBoxSeca.Value = ""
    Me.TextBoxAltitude.Value = ""
    Me.TextBoxAspecto.Value = ""
    Me.TextBoxGleba.Value = ""
    Me.TextBoxSacas.Value = ""
    Me.TextBoxSafraAno.Value = ""
    Me.TextBoxLoteCorrido.Value = ""
    Me.TextBoxCafe.Value = ""
    Me.TextBoxBicaCorrida.Value = ""
    Me.TextBoxCor.Value = ""
    Me.TextBoxPVA.Value = ""
    Me.TextBoxUmidade.Value = ""
    Me.TextBoxRenda.Value = ""
    Me.TextBoxCata.Value = ""
    Me.TextBoxObservacoes.Value = ""
    
       
    Me.ListBox_2.RowSource = "DADOS!Produtor"
    
    Me.TextBoxIDP.Value = ws.Range("Z1").Value
    Me.TextBoxIDAmostra.Value = ws.Range("Z1").Value
    Me.TextBoxIDClassificacao.Value = ws.Range("Z1").Value
   
   
    
    
    MsgBox "Cadastro Concluído com Sucesso!"


End Sub



Private Sub ListBox_2_Click()


    Dim codigo As Long
    Dim Data_abertura As Date
    Dim Data_fechamento As Date
    
    On Error Resume Next
    codigo = ListBox_2.List(ListBox_2.ListIndex, 0)
 
    Me.TextBoxIDP.Value = codigo
    Me.TextBoxIDAmostra = Me.TextBoxIDP.Value
    Me.TextBoxIDClassificacao = Me.TextBoxIDP.Value
    
    
     
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Planilha4")
    
    On Error Resume Next
    
    Me.TextBoxAbertura.Value = CDate(Format(Me.TextBoxAbertura.Text, "dd/mm/yyyy"))
    Me.TextBoxFechamento.Value = CDate(Format(Me.TextBoxFechamento.Text, "dd/mm/yyyy"))
    
    Me.BTAlterarRegistro.Enabled = True
    Me.BTDeletarl.Enabled = True
    Me.BTSalvarNovo.Enabled = False
    
    Produtor = ListBox_2.List(ListBox_2.ListIndex, 1)
    Email = ListBox_2.List(ListBox_2.ListIndex, 2)
    Fazenda = ListBox_2.List(ListBox_2.ListIndex, 3)
    cpf = ListBox_2.List(ListBox_2.ListIndex, 4)
    Cnpj = ListBox_2.List(ListBox_2.ListIndex, 5)
    Telefone = ListBox_2.List(ListBox_2.ListIndex, 6)
    CEP = ListBox_2.List(ListBox_2.ListIndex, 7)
    Rua = ListBox_2.List(ListBox_2.ListIndex, 8)
    Numero = ListBox_2.List(ListBox_2.ListIndex, 9)
    Regiao = ListBox_2.List(ListBox_2.ListIndex, 10)
    Cidade = ListBox_2.List(ListBox_2.ListIndex, 11)
    Estado = ListBox_2.List(ListBox_2.ListIndex, 12)
    Bairro = ListBox_2.List(ListBox_2.ListIndex, 13)
    Data_abertura = ListBox_2.List(ListBox_2.ListIndex, 14)
    Data_fechamento = ListBox_2.List(ListBox_2.ListIndex, 15)
    
    Amostra = ListBox_2.List(ListBox_2.ListIndex, 0)
    especie = ListBox_2.List(ListBox_2.ListIndex, 17)
    N_lote = ListBox_2.List(ListBox_2.ListIndex, 18)
    Variedade = ListBox_2.List(ListBox_2.ListIndex, 19)
    Padrao = ListBox_2.List(ListBox_2.ListIndex, 20)
    Cultivar = ListBox_2.List(ListBox_2.ListIndex, 21)
    Processo = ListBox_2.List(ListBox_2.ListIndex, 22)
    Seca = ListBox_2.List(ListBox_2.ListIndex, 23)
    Altitude = ListBox_2.List(ListBox_2.ListIndex, 24)
    Aspecto = ListBox_2.List(ListBox_2.ListIndex, 25)
    Gleba = ListBox_2.List(ListBox_2.ListIndex, 26)
    Sacas = ListBox_2.List(ListBox_2.ListIndex, 27)
    SafraAno = ListBox_2.List(ListBox_2.ListIndex, 28)
    LoteCorrido = ListBox_2.List(ListBox_2.ListIndex, 29)
    Cafe = ListBox_2.List(ListBox_2.ListIndex, 30)
    BicaCorrida = ListBox_2.List(ListBox_2.ListIndex, 31)
    PVA = ListBox_2.List(ListBox_2.ListIndex, 33)
    Cor = ListBox_2.List(ListBox_2.ListIndex, 32)
    Umidade = ListBox_2.List(ListBox_2.ListIndex, 34)
    renda = ListBox_2.List(ListBox_2.ListIndex, 35)
    Cata = ListBox_2.List(ListBox_2.ListIndex, 36)
    Observacoes = ListBox_2.List(ListBox_2.ListIndex, 37)
    
    
    ws.Range("L3").Value = codigo
    ws.Range("B5").Value = Produtor
    ws.Range("I7").Value = Telefone
    ws.Range("I5").Value = cpf
    ws.Range("B6").Value = Fazenda
    ws.Range("I6").Value = Cnpj
    ws.Range("I8").Value = Cidade
    ws.Range("I9").Value = Regiao
    ws.Range("B9").Value = Data_abertura
    ws.Range("E9").Value = Data_fechamento
    ws.Range("B7").Value = Email
    ws.Range("B8").Value = Endereço
    ws.Range("E8").Value = N
    ws.Range("L9").Value = Estado
    
    ws.Range("L1").Value = Amostra
    ws.Range("I11").Value = especie
    ws.Range("L11").Value = N_lote
    ws.Range("B13").Value = Variedade
    ws.Range("B14").Value = Padrao
    ws.Range("B15").Value = Cultivar
    ws.Range("B16").Value = Processo
    ws.Range("B17").Value = Seca
    ws.Range("B18").Value = Altitude
    ws.Range("B19").Value = Aspecto
    ws.Range("I13").Value = Gleba
    ws.Range("I14").Value = Sacas
    ws.Range("I15").Value = SafraAno
    ws.Range("I16").Value = LoteCorrido
    
    ws.Range("I17").Value = Cafe
    ws.Range("I18").Value = BicaCorrida
    ws.Range("I19").Value = Cor
    ws.Range("B20").Value = PVA
    ws.Range("D20").Value = Umidade
    ws.Range("I20").Value = renda
    ws.Range("L20").Value = Cata
    ws.Range("C21").Value = Observacoes

End Sub



Private Sub TextBoxAbertura_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.TextBoxAbertura.MaxLength = 10
    
    If Len(Me.TextBoxAbertura) = 2 Then
    Me.TextBoxAbertura.Text = Me.TextBoxAbertura.Text & "/"
    Me.TextBoxAbertura.SelStart = Len(Me.TextBoxAbertura)
    End If
    
    
    If Len(Me.TextBoxAbertura) = 5 Then
    Me.TextBoxAbertura.Text = Me.TextBoxAbertura.Text & "/"
    Me.TextBoxAbertura.SelStart = Len(Me.TextBoxAbertura)
    End If



End Sub



Private Sub TextBoxBuscaCPF_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.TextBoxBuscaCPF.MaxLength = 14
    
    If Len(Me.TextBoxBuscaCPF) = 3 Then
    Me.TextBoxBuscaCPF.Text = Me.TextBoxBuscaCPF.Text & "."
    Me.TextBoxBuscaCPF.SelStart = Len(Me.TextBoxBuscaCPF)
    End If
    
    If Len(Me.TextBoxBuscaCPF) = 7 Then
    Me.TextBoxBuscaCPF.Text = Me.TextBoxBuscaCPF.Text & "."
    Me.TextBoxBuscaCPF.SelStart = Len(Me.TextBoxBuscaCPF)
    End If
    
    If Len(Me.TextBoxBuscaCPF) = 11 Then
    Me.TextBoxBuscaCPF.Text = Me.TextBoxBuscaCPF.Text & "-"
    Me.TextBoxBuscaCPF.SelStart = Len(Me.TextBoxBuscaCPF)
    End If

End Sub



Private Sub TextBoxCEP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

  Me.TextBoxCEP.MaxLength = 9

    
    If Len(Me.TextBoxCEP) = 5 Then
    Me.TextBoxCEP.Text = Me.TextBoxCEP.Text & "-"
    Me.TextBoxCEP.SelStart = Len(Me.TextBoxCEP)
    End If


End Sub

Private Sub TextBoxCNPJ_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.TextBoxCNPJ.MaxLength = 18
    
    If Len(Me.TextBoxCNPJ) = 2 Then
    Me.TextBoxCNPJ.Text = Me.TextBoxCNPJ.Text & "."
    Me.TextBoxCNPJ.SelStart = Len(Me.TextBoxCNPJ)
    End If
    
    If Len(Me.TextBoxCNPJ) = 6 Then
    Me.TextBoxCNPJ.Text = Me.TextBoxCNPJ.Text & "."
    Me.TextBoxCNPJ.SelStart = Len(Me.TextBoxCNPJ)
    End If
    
    If Len(Me.TextBoxCNPJ) = 10 Then
    Me.TextBoxCNPJ.Text = Me.TextBoxCNPJ.Text & "/"
    Me.TextBoxCNPJ.SelStart = Len(Me.TextBoxCNPJ)
    End If
    
    If Len(Me.TextBoxCNPJ) = 15 Then
    Me.TextBoxCNPJ.Text = Me.TextBoxCNPJ.Text & "-"
    Me.TextBoxCNPJ.SelStart = Len(Me.TextBoxCNPJ)
    End If



End Sub

Private Sub TextBoxCPF_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.TextBoxCPF.MaxLength = 14
    
    If Len(Me.TextBoxCPF) = 3 Then
    Me.TextBoxCPF.Text = Me.TextBoxCPF.Text & "."
    Me.TextBoxCPF.SelStart = Len(Me.TextBoxCPF)
    End If
    
    If Len(Me.TextBoxCPF) = 7 Then
    Me.TextBoxCPF.Text = Me.TextBoxCPF.Text & "."
    Me.TextBoxCPF.SelStart = Len(Me.TextBoxCPF)
    End If
    
    If Len(Me.TextBoxCPF) = 11 Then
    Me.TextBoxCPF.Text = Me.TextBoxCPF.Text & "-"
    Me.TextBoxCPF.SelStart = Len(Me.TextBoxCPF)
    End If

End Sub



Private Sub TextBoxFechamento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Me.TextBoxFechamento.MaxLength = 10
    
    If Len(Me.TextBoxFechamento) = 2 Then
    Me.TextBoxFechamento.Text = Me.TextBoxFechamento.Text & "/"
    Me.TextBoxFechamento.SelStart = Len(Me.TextBoxFechamento)
    End If
    
    If Len(Me.TextBoxFechamento) = 5 Then
    Me.TextBoxFechamento.Text = Me.TextBoxFechamento.Text & "/"
    Me.TextBoxFechamento.SelStart = Len(Me.TextBoxFechamento)
    End If
    
    

End Sub

Private Sub TextBoxIDP_Change()

    If IsNumeric(TextBoxIDP.Value) = True Then
    
        Dim codigo As Long
        codigo = TextBoxIDP.Value
       
    
        Else
        
        End If
        
    On Error Resume Next
    
    'Produtor
    
    Me.TextBoxProdutor = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 2, 0)
    Me.TextBoxEmail = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 3, 0)
    Me.TextBoxFazenda = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 4, 0)
    Me.TextBoxCPF = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 5, 0)
    Me.TextBoxCNPJ = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 6, 0)
    Me.TextBoxTelefone = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 7, 0)
    Me.TextBoxCEP = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 8, 0)
    Me.TextBoxRua = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 9, 0)
    Me.TextBoxNumero = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 10, 0)
    Me.TextBoxRegiao = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 11, 0)
    Me.TextBoxCidade = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 12, 0)
    Me.TextBoxEstado = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 13, 0)
    Me.TextBoxBairro = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 14, 0)
    Me.TextBoxAbertura = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 15, 0)
    Me.TextBoxFechamento = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 16, 0)
    
    'Amostra
    
    Me.TextBoxAmostra = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 17, 0)
    Me.TextBoxEspecie = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 18, 0)
    Me.TextBoxLote = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 19, 0)
    Me.TextBoxVariedade = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 20, 0)
    Me.TextBoxPadrao = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 21, 0)
    Me.TextBoxCultivar = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 22, 0)
    Me.TextBoxProcesso = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 23, 0)
    Me.TextBoxSeca = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 24, 0)
    Me.TextBoxAltitude = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 25, 0)
    Me.TextBoxAspecto = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 26, 0)
    Me.TextBoxGleba = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 27, 0)
    Me.TextBoxSacas = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 28, 0)
    Me.TextBoxSafraAno = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 29, 0)
    Me.TextBoxLoteCorrido = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 30, 0)
    Me.TextBoxCafe = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 31, 0)
    Me.TextBoxBicaCorrida = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 32, 0)
    Me.TextBoxCor = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 33, 0)
    Me.TextBoxPVA = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 34, 0)
    Me.TextBoxUmidade = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 35, 0)
    Me.TextBoxRenda = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 36, 0)
    Me.TextBoxCata = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 37, 0)
    Me.TextBoxObservacoes = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 38, 0)
    Me.TextBoxPesoDefeitos = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 39, 0)
    
    Select Case (Me.ComboBoxDefeitos.Value)
    
    Case "Preto"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 40, 0)
    Case "Preto Verde"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 41, 0)
    Case "Ardido"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 42, 0)
    Case "Verde"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 43, 0)
    Case "Quebrado"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 44, 0)
    Case "Mau Granado / Chocho"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 45, 0)
    Case "Brocado Limpo"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 46, 0)
    Case "Brocado Sujo"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 47, 0)
    Case "Coco"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 48, 0)
    Case "Marinheiro"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 49, 0)
    Case "Casca Gd"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 50, 0)
    Case "Casca Md ou Pd"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 51, 0)
    Case "Fragmentos de Casca"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 52, 0)
    Case "Pau/Pedra/Torrão Gd"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 53, 0)
    Case "Pau/Pedra/Torrão Md"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 54, 0)
    Case "Pau/Pedra/Torrão Pq"
        Me.TextBoxDefeitos.Value = Application.WorksheetFunction.VLookup(codigo, Sheets("DADOS").Range("A:BF"), 55, 0)
    Case Else
        MsgBox "Defeito selecionado não é válido!"
    End Select
    
End Sub

Private Sub TextBoxTelefone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


    Me.TextBoxTelefone.MaxLength = 15
    
    If Len(Me.TextBoxTelefone) = 0 Then
    Me.TextBoxTelefone.Text = Me.TextBoxTelefone.Text & "("
    Me.TextBoxTelefone.SelStart = Len(Me.TextBoxTelefone)
    End If
    
    If Len(Me.TextBoxTelefone) = 3 Then
    Me.TextBoxTelefone.Text = Me.TextBoxTelefone.Text & ") "
    Me.TextBoxTelefone.SelStart = Len(Me.TextBoxTelefone)
    End If
    
    If Len(Me.TextBoxTelefone) = 10 Then
    Me.TextBoxTelefone.Text = Me.TextBoxTelefone.Text & "-"
    Me.TextBoxTelefone.SelStart = Len(Me.TextBoxTelefone)
    End If


End Sub

Private Sub UserForm_Activate()
    

    
    Me.ListBox_2.RowSource = "DADOS!Produtor"  'comentario teste
    Me.ListBox_2.ColumnCount = 55
    Me.ListBox_2.ColumnHeads = True
    Me.ListBox_2.ColumnWidths = "30;170;0;150;85;0;70;0;0;0;0;0;0;0;0;0;40;50;30;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
    
    Me.ComboBoxDefeitos.RowSource = "Defeitos"
    
    Me.Height = 610
    Me.Width = 960
    
    
End Sub


Private Sub UserForm_Initialize()

     Me.Height = 630
     Me.Width = 960
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
End Sub

