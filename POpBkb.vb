Imports Excel = Microsoft.Office.Interop.Excel

Public Class form_POpBkb

    Private Sub form_POpBkb_Load(sender As Object, e As EventArgs) Handles Me.Load

        Label1.Select()

    End Sub

    Private Sub btn_Fechar_Click(sender As Object, e As EventArgs) Handles btn_Fechar.Click

        Label1.Select()
        Application.Exit()

    End Sub

    Private Sub btn_Minimizar_Click(sender As Object, e As EventArgs) Handles btn_Minimizar.Click

        Label1.Select()
        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub btn_Home_Click(sender As Object, e As EventArgs) Handles btn_Home.Click

        Label1.Select()
        Me.Hide()
        MainMenu.Show()

    End Sub

    Private Sub btn_ArqPOpBkb_Click(sender As Object, e As EventArgs) Handles btn_ArqPOpBkb.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "POpBkb"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqNotas_Click(sender As Object, e As EventArgs) Handles btn_ArqNotas.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "Notas"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqComent_Click(sender As Object, e As EventArgs) Handles btn_ArqComent.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "Coment"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqAuv_Click(sender As Object, e As EventArgs) Handles btn_ArqAuv.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "Auv"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqTmaMtd_Click(sender As Object, e As EventArgs) Handles btn_ArqTmaMtd.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "TmaMtd"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqTmaD_Click(sender As Object, e As EventArgs) Handles btn_ArqTmaD.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "TmaD"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqTmaHH_Click(sender As Object, e As EventArgs) Handles btn_ArqTmaHH.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "TmaHH"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqDrive_Click(sender As Object, e As EventArgs) Handles btn_ArqDrive.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "Drive"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqVendaMtd_Click(sender As Object, e As EventArgs) Handles btn_ArqVendaMtd.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "VendaMtd"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqVendaD_Click(sender As Object, e As EventArgs) Handles btn_ArqVendaD.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "VendaD"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqRkDia_Click(sender As Object, e As EventArgs) Handles btn_ArqRkDia.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "RkDia"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqHistDisp_Click(sender As Object, e As EventArgs) Handles btn_ArqHistDisp.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "HistDisp"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqPDisp_Click(sender As Object, e As EventArgs) Handles btn_ArqPDisp.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "PDisp"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqBDVenda_Click(sender As Object, e As EventArgs) Handles btn_ArqBDVenda.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "BDVenda"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqCBK_Click(sender As Object, e As EventArgs) Handles btn_ArqCBK.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "CBK"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_DataPainel_Click(sender As Object, e As EventArgs) Handles btn_DataPainel.Click

        form_DataPOpsBKB.Show()

    End Sub

    Private Sub btn_ArqSand_Click(sender As Object, e As EventArgs) Handles btn_ArqSand.Click

        '------------------------------------------------------------------------------------------------------------------
        'Seta o Título do Open File Dialog
        '------------------------------------------------------------------------------------------------------------------
        ofd = Me.ofd1
        ofd.Title = "Selecione o " + DirectCast(sender, Button).Text

        '------------------------------------------------------------------------------------------------------------------
        'Chama função para selecionar Arquivo
        '------------------------------------------------------------------------------------------------------------------
        caminho = ""
        SelecionarArq(ofd, caminho)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------
        Dim nome As String

        formNome = Me
        nome = "Sand"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_GerarPainel_Click(sender As Object, e As EventArgs) Handles btn_GerarPainel.Click

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todas as planilhas abertas
        '------------------------------------------------------------------------------------------------------------------
        FecharWbksAbertos()

        '------------------------------------------------------------------------------------------------------------------
        'Definição de Variáveis
        '------------------------------------------------------------------------------------------------------------------
#Region "Definição de Variáveis"
        Dim vazios As Integer
        Dim nome As New ArrayList
        Dim nomes As New ArrayList
        Dim priLinha As Long
        Dim ultLinha As Long
        Dim ultLinha2 As Long
        Dim priColuna As String
        Dim ultColuna As String
        Dim ultColuna2 As String
        Dim priRange As String
        Dim segRange As String
        Dim terRange As String
        Dim quaRange As String
        Dim aba As String
        Dim salvar As String
        Dim dtPainel As String
        Dim lbNome As String
        Dim lbId As String
        Dim verExcel As Boolean
        Dim exportReport As Boolean
        Dim processo As String
        Dim inicio As DateTime
        Dim fim As DateTime
        Dim total As TimeSpan
        Dim processoPainel As String
        Dim inicioPainel As DateTime
        Dim fimPainel As DateTime
        Dim totalPainel As TimeSpan
        Dim dgvRelatorio As DataGridView
        Dim painel As String

        dgvRelatorio = Desenvolvedor.dgv_POpBkb 'DataGridView para salvar as infos de tempo de execução
        formNome = Me 'Nome do Formulário
        verExcel = Desenvolvedor.cbbox_Excel.Checked 'verificar status para ver o excel durante a execução do programa
        exportReport = Desenvolvedor.cbbox_Exportar.Checked 'verificar o status para saber se exporta o excel no fim da execução
#End Region

        '------------------------------------------------------------------------------------------------------------------
        'Atualizar Painel Operacional BKB
        '------------------------------------------------------------------------------------------------------------------
#Region "Atualizar Painel Operacional BKB"

        '------------------------------------------------------------------------------------------------------------------
        'Verifica se foi selecionada uma data
        '------------------------------------------------------------------------------------------------------------------
        If lbl_DataPainel.Text = "lbl_DataPainel" Then

            titulo = "Data Vazia"
            aviso = "Por favor selecione uma data para continuar."
            tipo = "Alerta"
            escolha = "Ok"

            Mensagem(aviso, tipo, titulo, escolha)

            lbNome_DataPainel.Text = "Selecione uma Data"
            lbNome_DataPainel.ForeColor = Color.Red

            GoTo fim

        End If

        '------------------------------------------------------------------------------------------------------------------
        'Verifica se todas as Labels de caminho estão preenchidas
        'chama a função checar caminhos que está no Arquivo.vb
        '------------------------------------------------------------------------------------------------------------------
#Region "Checar Caminhos"
        vazios = 0

        ChecarCaminhos(formNome, vazios, nome)
#End Region

        If vazios <> 0 Then

            '------------------------------------------------------------------------------------------------------------------
            'Se vazios é diferente de 0 ele lista todas as bases que falta selecionar
            '------------------------------------------------------------------------------------------------------------------
#Region "Mostrar Vazios"
            MostraVazios(formNome, vazios, nome, nomes)

            Dim bases As String = Join(nomes.ToArray, Chr(13))

            titulo = "Bases não selecionadas"
            aviso = "As bases abaixo não foram selecionadas:" & Chr(13) & Chr(13) & bases &
                   Chr(13) & Chr(13) & "Por favor selecione-as para continuar."
            tipo = "Erro"
            escolha = "Ok"

            Mensagem(aviso, tipo, titulo, escolha)
#End Region

        Else

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempo - Painel Operacional
            '------------------------------------------------------------------------------------------------------------------
            processoPainel = "Atualizar Painel Operacional BKB"
            inicioPainel = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Esconde o formulário Painel Operacional BKB, Abre o formulário de Loading e seta o valor da progress bar
            '------------------------------------------------------------------------------------------------------------------
            Me.Hide()
            form_Loading.ProgressBar1.Maximum = 48
            form_Loading.Show()

            '------------------------------------------------------------------------------------------------------------------
            'Se vazios é igual a 0 ele executa a rotina de colagem
            'Declaração das Panilhas Workbooks
            '------------------------------------------------------------------------------------------------------------------
#Region "Declaração de Planilhas"
            Dim VendaD, PDisp, HistDisp, POpBkb, Notas, Coment, Auv, Drive,
            TmaD, TmaMtd, TmaHH, VendaMtd, RkDia, BDVenda, CBK, Sand As Excel.Workbook

            VendaD = Nothing
            PDisp = Nothing
            HistDisp = Nothing
            POpBkb = Nothing
            Notas = Nothing
            Coment = Nothing
            Auv = Nothing
            Drive = Nothing
            TmaD = Nothing
            TmaMtd = Nothing
            TmaHH = Nothing
            VendaMtd = Nothing
            RkDia = Nothing
            BDVenda = Nothing
            CBK = Nothing
            Sand = Nothing
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Disponibilidade
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Venda D-1 -> Painel Disponibilidade"
            processo = "Venda D-1 -> Painel Disponibilidade"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Venda D-1
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Venda D-1"
            lbNome = "lb_VendaD"
            lbId = "lbId_VendaD"

            AbrirArq(formNome, VendaD, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If VendaD Is Nothing Then

                lbNome_VendaD.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel Disponibilidade."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Painel Disponibilidade"
            lbNome = "lb_PDisp"
            lbId = "lbId_PDisp"

            AbrirArq(formNome, PDisp, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If PDisp Is Nothing Then

                lbNome_PDisp.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados da Base de Dados do Painel Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados da Base de Dados do Painel Disponibilidade"
            aba = "Base de Dados"
            priLinha = 11
            priColuna = "A"
            ultColuna = "N"

            Apagar(PDisp, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Disponibilidade."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Vendas D-1
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Vendas D-1"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "N"

            Copiar(VendaD, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados da Base de Dados do Painel Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados da Base de Dados do Painel Disponibilidade"
            aba = "Base de Dados"
            priLinha = 11
            priColuna = "A"

            Colar(PDisp, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas Bases de Dados
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "Base de Dados"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "P"
            ultColuna2 = "AB"

            PropFormulas(PDisp, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Venda D-1
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Venda D-1"
            salvar = "N"
            lbId = "lbId_VendaD"

            FecharWbks(formNome, VendaD, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Realcula todas as abas e atualiza as dinâmicas das abas no Painel Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
            With PDisp

                .RefreshAll()
                .Application.Calculate()
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas Colar em Consolidado
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "Colar em Consolidado"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "AF"
            ultColuna2 = "AI"

            PropFormulas(PDisp, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas Colar em Disp PDV
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "Colar em Disp PDV"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "AG"
            ultColuna2 = "AH"

            PropFormulas(PDisp, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Histórico Disponibilidade
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Painel Disponibilidade -> Histórico Disponibilidade"
            processo = "Painel Disponibilidade -> Histórico Disponibilidade"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Histórico Disponibilidade."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Histórico Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Histórico Disponibilidade"
            lbNome = "lb_histDisp"
            lbId = "lbId_HistDisp"

            AbrirArq(formNome, HistDisp, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If HistDisp Is Nothing Then

                lbNome_HistDisp.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Histórico Disponibilidade."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Painel Disp. para o Hist Disp - PDV Abertos
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Painel Disp. para o Hist Disp - PDV Abertos"
            aba = "PDV Abertos"
            priLinha = 8
            ultLinha = 0
            priColuna = "H"
            ultColuna = "R"

            Copiar(PDisp, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Acha a Coluna com a data selecionada no Histórico Disponibilidade para colar as informações - Histórico PDVs
            '------------------------------------------------------------------------------------------------------------------
#Region "Acha a Coluna com a data selecionada no Histórico Disponibilidade para colar as informações - Histórico PDVs"
            aba = "Histórico Balcão"
            priLinha = 3
            ultLinha = 6
            dtPainel = lbl_DataPainel.Text

            ColarCriterioDt(HistDisp, aba, priLinha, ultLinha, dtPainel)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Realcula todas as abas e atualiza as dinâmicas das abas no Histórico Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
            With HistDisp

                .RefreshAll()
                .Application.Calculate()
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel Operacional BKB."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Painel Operacional BKB"
            lbNome = "lb_pOpBkb"
            lbId = "lbId_POpBkb"

            AbrirArq(formNome, POpBkb, lbNome, lbId, verExcel)

            Threading.Thread.Sleep(5000)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If POpBkb Is Nothing Then

                lbNome_POpBkb.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com as informações do Histórico Disponibilidade.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Histórico Disponibilidade -> Painel Operacional BKB"
            processo = "Histórico Disponibilidade -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com as informações do Histórico Disponibilidade."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza as Informações de Disponibilidade no Painel Operacional
            'Copia os dados do Painel Disp. para o Hist Disp - Disp KSK
            '------------------------------------------------------------------------------------------------------------------
#Region "Colar aba Painel Operacional"
            aba = "Colar aba Painel Operacional"
            priLinha = 4
            ultLinha = 0
            priColuna = "B"
            ultColuna = "N"

            Copiar(HistDisp, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados no 9.Disponibilidade do Painel Op BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados no 9.Disponibilidade do Painel Op BKB"
            aba = "9. Disponibilidade MTD"
            priLinha = 4
            priColuna = "B"

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Histórico Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Histórico Disponibilidade"
            salvar = "N"
            lbId = "lbId_HistDisp"

            FecharWbks(formNome, HistDisp, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com as informações do Painel Disponibilidade.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Painel Disponibilidade -> Painel Operacional BKB"
            processo = "Painel Disponibilidade -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados da Disp PDV do Painel Op BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados da Disp PDV do Painel Op BKB"
            aba = "Disp. PDV"
            priLinha = 6
            priColuna = "A"
            ultColuna = "AH"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com as informações do Painel Disponibilidade."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Colar em Disp PDV
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Colar em Disp PDV"
            aba = "Colar em Disp PDV"
            priLinha = 7
            ultLinha = 0
            priColuna = "A"
            ultColuna = "AH"

            Copiar(PDisp, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados no Colar em Disp. PDV
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados no Colar em Disp. PDV"
            aba = "Disp. PDV"
            priColuna = "A"
            priLinha = 6

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar formatação condicional
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar formatação condicional na Disp. PDV"
            aba = "Disp. PDV"
            priColuna = "A"
            ultColuna = "AH"
            priLinha = 6
            ultLinha = 0

            PropFormatacao(POpBkb, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do Disp. PDV Consolid.
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do Disp. PDV Consolid."
            aba = "Disp. PDV Consolid."
            priLinha = 6
            priColuna = "B"
            ultColuna = "AM"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Colar em Consolidado
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Colar em Consolidado"
            aba = "Colar em Consolidado"
            priLinha = 7
            ultLinha = 0
            priColuna = "A"
            ultColuna = "AL"

            Copiar(PDisp, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em Disp. PDV Consolid.
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em Disp. PDV Consolid."
            aba = "Disp. PDV Consolid."
            priColuna = "B"
            priLinha = 6

            ColarTexto(POpBkb, aba, priLinha, priColuna)

            POpBkb.Worksheets(aba).calculate
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar formatação condicional
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar formatação condicional na Disp. PDV Consolid."
            aba = "Disp. PDV Consolid."
            priColuna = "B"
            ultColuna = "AM"
            priLinha = 6
            ultLinha = 0

            PropFormatacao(POpBkb, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Painel Disponibilidade
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Painel Disponibilidade"
            salvar = "N"
            lbId = "lbId_PDisp"

            FecharWbks(formNome, PDisp, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpBkb

                .Application.CalculateBeforeSave = False
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Ranking Diário."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Ranking Diário"
            lbNome = "lb_RkDia"
            lbId = "lbId_RkDia"

            AbrirArq(formNome, RkDia, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If RkDia Is Nothing Then

                lbNome_RkDia.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com as informações do Ranking Diário.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Ranking Diário -> Painel Operacional BKB"
            processo = "Ranking Diário -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com as informações do Ranking Diário."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 6. Delivery Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 6. Delivery Comentários - Painel Operacional BKB"
            aba = "6. Delivery Comentários"
            priLinha = 2
            priColuna = "E"
            ultColuna = "L"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Comentários - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Comentários - Ranking Diário"
            aba = "Comentários"
            priLinha = 2
            ultLinha = 0
            priColuna = "B"
            ultColuna = "B"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6. Delivery Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional BKB"
            aba = "6. Delivery Comentários"
            priColuna = "E"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Apagar coluna Mês Aba Comentários - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Apagar coluna Mês Aba Comentários - Ranking Diário"
            aba = "Comentários"
            priColuna = "K"
            ultColuna = "K"

            ApagarColuna(RkDia, aba, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Comentários - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Comentários - Ranking Diário"
            aba = "Comentários"
            priLinha = 2
            ultLinha = 0
            priColuna = "G"
            ultColuna = "K"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6. Delivery Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional BKB"
            aba = "6. Delivery Comentários"
            priColuna = "F"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Comentários - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Comentários - Ranking Diário"
            aba = "Comentários"
            priLinha = 2
            ultLinha = 0
            priColuna = "M"
            ultColuna = "M"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6. Delivery Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional BKB"
            aba = "6. Delivery Comentários"
            priColuna = "K"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Comentários - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Comentários - Ranking Diário"
            aba = "Comentários"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "A"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6. Delivery Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional BKB"
            aba = "6. Delivery Comentários"
            priColuna = "L"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "6. Delivery Comentários"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "E"
            ultColuna = "A"
            ultColuna2 = "D"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 6.1 Delivery Tempo e Notas - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 6.1 Delivery Tempo e Notas - Painel Operacional BKB"
            aba = "6.1 Delivery Tempo e Notas"
            priRange = "F6:AJ22"
            With POpBkb.Worksheets(aba)
                ultLinha = .range("A" & .rows.count).end(Excel.XlDirection.xlUp).row
            End With
            segRange = "A42:AJ" & ultLinha
            With POpBkb.Worksheets(aba)
                ultLinha2 = .range("AL" & .rows.count).end(Excel.XlDirection.xlUp).row
            End With
            terRange = "AL6:BQ" & ultLinha2
            quaRange = "F29:AJ33"

            ApagarDlvTN(POpBkb, aba, priRange, segRange, terRange, quaRange)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Reg - Rest - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Reg - Rest - Ranking Diário"
            aba = "Reg - Rest"
            priLinha = 6
            ultLinha = 22
            priColuna = "F"
            ultColuna = "AJ"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "F"
            priLinha = 6

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Reg - Rest - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Reg - Rest - Ranking Diário"
            aba = "Reg - Rest"
            priLinha = 29
            ultLinha = 33
            priColuna = "F"
            ultColuna = "AJ"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "F"
            priLinha = 29

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Reg - Rest - Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Reg - Rest - Ranking Diário"
            aba = "Reg - Rest"
            priLinha = 40
            ultLinha = 0
            priColuna = "A"
            ultColuna = "AJ"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "A"
            priLinha = 42

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Setor - Ranking Diário - 1ª PARTE
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Setor - Ranking Diário"
            aba = "Setor"
            priLinha = 5
            ultLinha = 0
            priColuna = "B"
            ultColuna = "C"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB - 1ª PARTE
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "AL"
            priLinha = 6

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Setor - Ranking Diário - 2ª PARTE
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Setor - Ranking Diário"
            aba = "Setor"
            priLinha = 5
            ultLinha = 0
            priColuna = "Q"
            ultColuna = "AT"

            Copiar(RkDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB - 2ª PARTE
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional BKB"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "AN"
            priLinha = 6

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Ranking Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Ranking Diário"
            salvar = "N"
            lbId = "lbId_RkDia"

            FecharWbks(formNome, RkDia, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional BKB."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpBkb

                .Application.CalculateBeforeSave = False
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Venda MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Venda MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Venda MTD"
            lbNome = "lb_VendaMtd"
            lbId = "lbId_VendaMtd"

            AbrirArq(formNome, VendaMtd, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If VendaMtd Is Nothing Then

                lbNome_VendaMtd.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Base Dinâmica Vendas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Base Dinâmica Vendas"
            lbNome = "lb_BDVenda"
            lbId = "lbId_BDVenda"

            AbrirArq(formNome, BDVenda, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If BDVenda Is Nothing Then

                lbNome_BDVenda.ForeColor = Color.Red
                GoTo fim

            End If

#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Base Dinâmica Vendas com informações do Venda MTD.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Venda MTD -> Base Dinâmica Vendas"
            processo = "Venda MTD -> Base Dinâmica Vendas"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Base Dinâmica Vendas com informações do Venda MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 1. COLAR BASE SAP - Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 1. COLAR BASE SAP - Base Dinâmica Vendas"
            aba = "1. COLAR BASE SAP"
            priLinha = 4
            priColuna = "B"
            ultColuna = "G"

            Apagar(BDVenda, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Venda MTD - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Venda MTD - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "F"

            Copiar(VendaMtd, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 1. COLAR BASE SAP - Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 1. COLAR BASE SAP - Base Dinâmica Vendas"
            aba = "1. COLAR BASE SAP"
            priColuna = "B"
            priLinha = 4

            Colar(BDVenda, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "1. COLAR BASE SAP"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "B"
            ultColuna = "H"
            ultColuna2 = "N"

            PropFormulas(BDVenda, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Fecha Venda MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha Venda MTD"
            salvar = "N"
            lbId = "lbId_VendaMtd"

            FecharWbks(formNome, VendaMtd, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Salva e Atualiza o Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
            With BDVenda

                .RefreshAll()

                Threading.Thread.Sleep(15000)

                .Application.Calculate()

                '------------------------------------------------------------------------------------------------------------------
                'Repete o Processo para atualizar as duas planilhas
                '------------------------------------------------------------------------------------------------------------------
                form_Loading.lbl_Loading.Text = "Atualizando Base Dinâmica Vendas"
                form_Loading.ProgressBar1.PerformStep()

                .RefreshAll()

                Threading.Thread.Sleep(15000)

                .Application.Calculate()

                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)

#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do Base Dinâmica Vendas.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Base Dinâmica Vendas -> Painel Operacional BKB"
            processo = "Base Dinâmica Vendas -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do Base Dinâmica Vendas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 1. Base SAP - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 1. Base SAP - Painel Operacional BKB"
            aba = "1. Base SAP"
            priLinha = 3
            priColuna = "A"
            ultColuna = "M"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do 1. COLAR BASE SAP - Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do 1. COLAR BASE SAP - Base Dinâmica Vendas"
            aba = "1. COLAR BASE SAP"
            priLinha = 4
            ultLinha = 0
            priColuna = "B"
            ultColuna = "N"

            Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 1. Base SAP - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 1. Base SAP - Painel Operacional BKB"
            aba = "1. Base SAP"
            priColuna = "A"
            priLinha = 3

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 8. Base dinâmica vendas - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 8. Base dinâmica vendas - Painel Operacional BKB"
            aba = "8. Base dinâmica vendas"
            priLinha = 2
            priColuna = "A"
            ultColuna = "M"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do 3. ATUALIZAR TRATAMENTO - Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do 3. ATUALIZAR TRATAMENTO - Base Dinâmica Vendas"
            aba = "3. ATUALIZAR TRATAMENTO"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 8. Base dinâmica vendas - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 8. Base dinâmica vendas - Painel Operacional BKB"
            aba = "8. Base dinâmica vendas"
            priColuna = "A"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Base Dinâmica Vendas"
            salvar = "N"
            lbId = "lbId_BDVenda"

            FecharWbks(formNome, BDVenda, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional BKB."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpBkb

                .Application.CalculateBeforeSave = False
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo B3 Notas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir B3 - Notas
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir B3 - Notas"
            lbNome = "lb_Notas"
            lbId = "lbId_Notas"

            AbrirArq(formNome, Notas, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Notas Is Nothing Then

                lbNome_Notas.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do B3 - Notas.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "B3 - Notas -> Painel Operacional BKB"
            processo = "B3 - Notas -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do B3 - Notas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 2. GT Notas B3 - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 2. GT Notas B3 - Painel Operacional BKB"
            aba = "2. GT Notas B3"
            priLinha = 3
            priColuna = "A"
            ultColuna = "C"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da B3 - Notas - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da B3 - Notas - 1ª aba"
            aba = "1"
            priLinha = 3
            ultLinha = 0
            priColuna = "A"
            ultColuna = "C"

            Copiar(Notas, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 2. GT Notas B3 - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 2. GT Notas B3 - Painel Operacional BKB"
            aba = "2. GT Notas B3"
            priColuna = "A"
            priLinha = 3

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "2. GT Notas B3"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "D"
            ultColuna2 = "G"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fecha B3 - Notas
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha B3 - Notas"
            salvar = "N"
            lbId = "lbId_Notas"

            FecharWbks(formNome, Notas, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo B3 Comentários."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir B3 - Comentários
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir B3 - Comentários"
            lbNome = "lb_Coment"
            lbId = "lbId_Coment"

            AbrirArq(formNome, Coment, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Coment Is Nothing Then

                lbNome_Coment.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do B3 - Comentários.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "B3 - Comentários -> Painel Operacional BKB"
            processo = "B3 - Comentários -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do B3 - Comentários."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 2.1 GT - Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 2.1 GT - Comentários - Painel Operacional BKB"
            aba = "2.1 GT - Comentários"
            priLinha = 3
            priColuna = "A"
            ultColuna = "E"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da B3 - Comentários - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da B3 - Comentários - 1ª aba"
            aba = "1"
            priLinha = 3
            ultLinha = 0
            priColuna = "A"
            ultColuna = "E"

            Copiar(Coment, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 2.1 GT - Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 2.1 GT - Comentários - Painel Operacional BKB"
            aba = "2.1 GT - Comentários"
            priColuna = "A"
            priLinha = 3

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "2.1 GT - Comentários"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "F"
            ultColuna2 = "L"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fecha B3 - Comentários
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha B3 - Comentários"
            salvar = "N"
            lbId = "lbId_Coment"

            FecharWbks(formNome, Coment, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo AUV."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir AUV
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir AUV"
            lbNome = "lb_Auv"
            lbId = "lbId_Auv"

            AbrirArq(formNome, Auv, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Auv Is Nothing Then

                lbNome_Auv.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do AUV.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "AUV -> Painel Operacional BKB"
            processo = "AUV -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 3. AUV - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 3. AUV - Painel Operacional BKB"
            aba = "3. AUV"
            priLinha = 2
            priColuna = "A"
            ultColuna = "D"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do AUV."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da AUV - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da AUV - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "D"

            Copiar(Auv, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 3. AUV - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 3. AUV - Painel Operacional BKB"
            aba = "3. AUV"
            priColuna = "A"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "3. AUV"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "F"
            ultColuna2 = "H"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar AUV
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar AUV"
            salvar = "N"
            lbId = "lbId_Auv"

            FecharWbks(formNome, Auv, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Drive."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Drive
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Drive"
            lbNome = "lb_Drive"
            lbId = "lbId_Drive"

            AbrirArq(formNome, Drive, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Drive Is Nothing Then

                lbNome_Drive.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do Drive.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Drive -> Painel Operacional BKB"
            processo = "Drive -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do Drive."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 4. Drive - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 4. Drive - Painel Operacional BKB"
            aba = "4. Drive"
            priLinha = 3
            priColuna = "D"
            ultColuna = "L"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Drive - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Drive - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "I"

            Copiar(Drive, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 4. Drive - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 4. Drive - Painel Operacional BKB"
            aba = "4. Drive"
            priColuna = "D"
            priLinha = 3

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "4. Drive"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "D"
            ultColuna = "A"
            ultColuna2 = "B"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fecha Drive
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha Drive"
            salvar = "N"
            lbId = "lbId_Drive"

            FecharWbks(formNome, Drive, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional BKB."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpBkb

                .Application.CalculateBeforeSave = False
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo TMA D-1."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir TMA D-1
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir TMA D-1"
            lbNome = "lb_TmaD"
            lbId = "lbId_TmaD"

            AbrirArq(formNome, TmaD, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If TmaD Is Nothing Then

                lbNome_TmaD.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do TMA D-1.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "TMA D-1 -> Painel Operacional BKB"
            processo = "TMA D-1 -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do TMA D-1."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 7. TMA D-1 - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 7. TMA D-1 - Painel Operacional BKB"
            aba = "7. TMA D-1"
            priLinha = 3
            priColuna = "A"
            ultColuna = "H"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da TMA D-1 - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da TMA D-1 - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "H"

            Copiar(TmaD, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 7. TMA D-1 - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 7. TMA D-1 - Painel Operacional BKB"
            aba = "7. TMA D-1"
            priColuna = "A"
            priLinha = 3

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "7. TMA D-1"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "L"
            ultColuna2 = "T"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fecha TMA D-1
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha TMA D-1"
            salvar = "N"
            lbId = "lbId_TmaD"

            FecharWbks(formNome, TmaD, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo TMA MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir TMA MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir TMA MTD"
            lbNome = "lb_TmaMtd"
            lbId = "lbId_TmaMtd"

            AbrirArq(formNome, TmaMtd, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If TmaMtd Is Nothing Then

                lbNome_TmaMtd.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do TMA MTD.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "TMA MTD -> Painel Operacional BKB"
            processo = "TMA MTD -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do TMA MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 7.1 TMA MTD - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 7.1 TMA MTD - Painel Operacional BKB"
            aba = "7.1 TMA MTD"
            priLinha = 3
            priColuna = "A"
            ultColuna = "H"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da TMA MTD - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da TMA MTD - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "H"

            Copiar(TmaMtd, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 7.1 TMA MTD - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 7.1 TMA MTD - Painel Operacional BKB"
            aba = "7.1 TMA MTD"
            priColuna = "A"
            priLinha = 3

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "7.1 TMA MTD"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "L"
            ultColuna2 = "T"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Venda MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Venda MTD"
            salvar = "N"
            lbId = "lbId_TmaMtd"

            FecharWbks(formNome, TmaMtd, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo TMA Hora Hora."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir TMA Hora Hora
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir TMA Hora Hora"
            lbNome = "lb_TmaHH"
            lbId = "lbId_TmaHH"

            AbrirArq(formNome, TmaHH, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If TmaHH Is Nothing Then

                lbNome_TmaHH.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do TMA Hora Hora.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "TMA Hora Hora -> Painel Operacional BKB"
            processo = "TMA Hora Hora -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do TMA Hora Hora."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 7.2 TMA Hr a Hr - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 7.2 TMA Hr a Hr - Painel Operacional BKB"
            aba = "7.2 TMA Hr a Hr"
            priLinha = 2
            priColuna = "D"
            ultColuna = "AC"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da TMA HH - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da TMA HH - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "Z"

            Copiar(TmaHH, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 7.2 TMA Hr a Hr - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 7.2 TMA Hr a Hr - Painel Operacional BKB"
            aba = "7.2 TMA Hr a Hr"
            priColuna = "D"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "7.2 TMA Hr a Hr"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "D"
            ultColuna = "A"
            ultColuna2 = "B"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fecha TMA Hora Hora
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha TMA Hora Hora"
            salvar = "N"
            lbId = "lbId_TmaHH"

            FecharWbks(formNome, TmaHH, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Clube BK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Clube BK
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Clube BK"
            lbNome = "lb_CBK"
            lbId = "lbId_CBK"

            AbrirArq(formNome, CBK, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If CBK Is Nothing Then

                lbNome_CBK.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações do Clube BK.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Clube BK -> Painel Operacional BKB"
            processo = "Clube BK -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações do Clube BK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 10. Clube BK - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 10. Clube BK - Painel Operacional BKB"
            aba = "10. Clube BK"
            priLinha = 2
            priColuna = "A"
            ultColuna = "C"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Clube BK - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Clube BK - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "C"

            Copiar(CBK, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 10. Clube BK - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 10. Clube BK - Painel Operacional BKB"
            aba = "10. Clube BK"
            priColuna = "A"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "10. Clube BK"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "E"
            ultColuna2 = "G"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fecha 10. Clube BK
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha 10. Clube BK"
            salvar = "N"
            lbId = "lbId_CBK"

            FecharWbks(formNome, CBK, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Qtd. Sanduiche."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Qtd. Sanduiche
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Qtd. Sanduiche"
            lbNome = "lb_Sand"
            lbId = "lbId_Sand"

            AbrirArq(formNome, Sand, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Sand Is Nothing Then

                lbNome_Sand.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional BKB com informações de Qtd. Sanduiche.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Qtd. Sanduiche -> Painel Operacional BKB"
            processo = "Qtd. Sanduiche -> Painel Operacional BKB"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional BKB com informações de Qtd. Sanduiche."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 11. Qtd. Sanduiche - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 11. Qtd. Sanduiche - Painel Operacional BKB"
            aba = "11. Qtd. Sanduiche"
            priLinha = 2
            priColuna = "A"
            ultColuna = "D"

            Apagar(POpBkb, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Qtd. Sanduiche - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Qtd. Sanduiche - 1ª aba"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "D"

            Copiar(Sand, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 11. Qtd. Sanduiche - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 11. Qtd. Sanduiche - Painel Operacional BKB"
            aba = "11. Qtd. Sanduiche"
            priColuna = "A"
            priLinha = 2

            Colar(POpBkb, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "11. Qtd. Sanduiche"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "F"
            ultColuna2 = "H"

            PropFormulas(POpBkb, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fecha Qtd. Sanduiche
            '------------------------------------------------------------------------------------------------------------------
#Region "Fecha Qtd. Sanduiche"
            salvar = "N"
            lbId = "lbId_Sand"

            FecharWbks(formNome, Sand, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso e Abrindo o Painel Operacional BKB."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpBkb

                .Application.CalculateBeforeSave = False
                .Save()

                '------------------------------------------------------------------------------------------------------------------
                'Final do código
                '------------------------------------------------------------------------------------------------------------------
                .Application.CalculateBeforeSave = True
                .Application.Visible = True

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos - Painel Operacional
            '------------------------------------------------------------------------------------------------------------------
            fimPainel = Now.ToLongTimeString
            totalPainel = fimPainel.Subtract(inicioPainel)

            GravarHoras(dgvRelatorio, processoPainel, inicioPainel, fimPainel, totalPainel)

        End If

fim:

        Me.Show()
        form_Loading.Hide()

#End Region

        '------------------------------------------------------------------------------------------------------------------
        'Exporta um Report com os tempos de cada Painel
        '------------------------------------------------------------------------------------------------------------------
#Region "Reports"
        If exportReport = True Then

            painel = "Operacional BKB"

            ExportarHoras(dgvRelatorio, painel)

        End If
#End Region

    End Sub

End Class