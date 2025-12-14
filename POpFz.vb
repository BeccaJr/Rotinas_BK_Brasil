Imports Excel = Microsoft.Office.Interop.Excel

Public Class form_POpFz

    Private Sub form_POpFz_Load(sender As Object, e As EventArgs) Handles Me.Load

        Label2.Select()

    End Sub

    Private Sub btn_Fechar_Click_1(sender As Object, e As EventArgs) Handles btn_Fechar.Click

        Label2.Select()
        Application.Exit()

    End Sub

    Private Sub btn_Minimizar_Click_1(sender As Object, e As EventArgs) Handles btn_Minimizar.Click

        Label2.Select()
        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub btn_Home_Click_1(sender As Object, e As EventArgs) Handles btn_Home.Click

        Label2.Select()
        Me.Hide()
        MainMenu.Show()

    End Sub

    Private Sub btn_ArqPOpFz_Click(sender As Object, e As EventArgs) Handles btn_ArqPOpFz.Click

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
        nome = "POpFz"

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

    Private Sub btn_ArqVendaDFz1_Click(sender As Object, e As EventArgs) Handles btn_ArqVendaDFz1.Click

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
        nome = "VendaDFz1"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqVendaDFz2_Click(sender As Object, e As EventArgs) Handles btn_ArqVendaDFz2.Click

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
        nome = "VendaDFz2"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqPHoras_Click(sender As Object, e As EventArgs) Handles btn_ArqPHoras.Click

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
        nome = "PHoras"

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

    Private Sub btn_GerarPainel_Click(sender As Object, e As EventArgs) Handles btn_GerarPainel.Click

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todas as planilhas abertas
        '------------------------------------------------------------------------------------------------------------------
        FecharWbksAbertos()

        '------------------------------------------------------------------------------------------------------------------
        'Zerar Progresso do Loading
        '------------------------------------------------------------------------------------------------------------------
        form_Loading.ProgressBar1.Value = 0

        '------------------------------------------------------------------------------------------------------------------
        'Definição de Variáveis
        '------------------------------------------------------------------------------------------------------------------
#Region "Definição de Variáveis"
        Dim vazios As Integer
        Dim nome As New ArrayList
        Dim nomes As New ArrayList
        Dim i As Long
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

        dgvRelatorio = Desenvolvedor.dgv_POpFz 'DataGridView para salvar as infos de tempo de execução
        formNome = Me 'Nome do Formulário
        verExcel = Desenvolvedor.cbbox_Excel.Checked 'verificar status para ver o excel durante a execução do programa
        exportReport = Desenvolvedor.cbbox_Exportar.Checked 'verificar o status para saber se exporta o excel no fim da execução
#End Region

        '------------------------------------------------------------------------------------------------------------------
        'Atualizar Painel Operacional FZ
        '------------------------------------------------------------------------------------------------------------------
#Region "Atualizar Painel Operacional FZ"

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
            processoPainel = "Atualizar Painel Operacional FZ"
            inicioPainel = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Esconde o formulário Painel Operacional FZ, Abre o formulário de Loading e seta o valor da progress bar
            '------------------------------------------------------------------------------------------------------------------
            Me.Hide()
            form_Loading.ProgressBar1.Maximum = 39
            form_Loading.Show()

            '------------------------------------------------------------------------------------------------------------------
            'Declaração das Panilhas Workbooks
            '------------------------------------------------------------------------------------------------------------------
#Region "Declaração de Planilhas"
            Dim POpFz, RkDia, BDVenda, TmaMtd, TmaD, VendaDFz1, VendaDFz2,
                PHoras, Auv, Notas, Coment, Drive As Excel.Workbook

            POpFz = Nothing
            RkDia = Nothing
            BDVenda = Nothing
            TmaMtd = Nothing
            TmaD = Nothing
            VendaDFz1 = Nothing
            VendaDFz2 = Nothing
            PHoras = Nothing
            Auv = Nothing
            Notas = Nothing
            Coment = Nothing
            Drive = Nothing
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel de Horas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Venda D-1 - FZ1 e FZ2 -> Painel de Horas"
            processo = "Venda D-1 - FZ1 e FZ2 -> Painel de Horas"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Venda D-1 FZ1
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Venda D-1 FZ1"
            lbNome = "lb_VendaDFz1"
            lbId = "lbId_VendaDFz1"

            AbrirArq(formNome, VendaDFz1, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If VendaDFz1 Is Nothing Then

                lbNome_VendaDFz1.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Venda D-1 FZ2
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Venda D-1 FZ2"
            lbNome = "lb_VendaDFz2"
            lbId = "lbId_VendaDFz2"

            AbrirArq(formNome, VendaDFz2, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If VendaDFz2 Is Nothing Then

                lbNome_VendaDFz2.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel de Horas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Painel de Horas"
            lbNome = "lb_PHoras"
            lbId = "lbId_PHoras"

            AbrirArq(formNome, PHoras, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If PHoras Is Nothing Then

                lbNome_PHoras.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados da Base de Dados do Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados da Base de Dados do Painel de Horas"
            aba = "Base de Dados"
            priLinha = 14
            priColuna = "A"
            ultColuna = "M"

            Apagar(PHoras, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel de Horas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Vendas D-1 FZ1
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Vendas D-1 FZ1"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(VendaDFz1, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados da Base de Dados do Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados da Base de Dados do Painel de Horas"
            aba = "Base de Dados"
            priLinha = 14
            priColuna = "A"

            Colar(PHoras, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Vendas D-1 FZ2
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Vendas D-1 FZ2"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(VendaDFz2, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados da Base de Dados do Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados da Base de Dados do Painel de Horas"
            aba = "Base de Dados"
            priColuna = "A"
            priLinha = PHoras.Worksheets(aba).range(priColuna & PHoras.Worksheets(aba).rows.count).end(Excel.XlDirection.xlUp).row + 1

            Colar(PHoras, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "Base de Dados"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "O"
            ultColuna2 = "Z"

            PropFormulas(PHoras, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Venda D-1 FZ1 e FZ2
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Venda D-1 FZ1 e FZ2"
            salvar = "N"
            lbId = "lbId_VendaDFz1"

            FecharWbks(formNome, VendaDFz1, salvar, lbId)

            salvar = "N"
            lbId = "lbId_VendaDFz2"

            FecharWbks(formNome, VendaDFz2, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Realcula todas as abas e atualiza as dinâmicas das abas no Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
            With PHoras

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

            PropFormulas(PHoras, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
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

            PropFormulas(PHoras, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Base Dinâmica Vendas com Painel de Horas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Painel de Horas -> Base Dinâmica Vendas"
            processo = "Painel de Horas -> Base Dinâmica Vendas"
            inicio = Now.ToLongTimeString

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
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atulizando Base Dinâmica Vendas com informações do Painel de Horas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Painel de Horas"
            aba = "BASE SAP VENDA"
            priLinha = 5
            ultLinha = -1
            priColuna = "A"
            ultColuna = "E"

            Copiar(PHoras, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na Base Dinâmica Vendas"
            aba = "1. COLAR BASE SAP"
            priColuna = "B"
            priLinha = BDVenda.Worksheets(aba).range(priColuna & BDVenda.Worksheets(aba).rows.count).end(Excel.XlDirection.xlUp).row + 1

            Colar(BDVenda, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Transforma os BKNs em Números
            '------------------------------------------------------------------------------------------------------------------
            With BDVenda.Worksheets("1. COLAR BASE SAP").range("C:C")

                .numberformat = "General"
                Threading.Thread.Sleep(2500)
                .formulaLocal = .value

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "1. COLAR BASE SAP"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "B"
            ultColuna = "G"
            ultColuna2 = "M"

            PropFormulas(BDVenda, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando dados do Base Dinâmica Vendas"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salva, Atualiza e Cola Valor nas Fórmulas - Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
            With BDVenda

                .RefreshAll()

                Threading.Thread.Sleep(15000)

                .Application.Calculate()

                '------------------------------------------------------------------------------------------------------------------
                'Repete o Processo para atualizar as duas planilhas
                '------------------------------------------------------------------------------------------------------------------
                form_Loading.ProgressBar1.PerformStep()

                .RefreshAll()

                Threading.Thread.Sleep(15000)

                .Application.Calculate()

                '------------------------------------------------------------------------------------------------------------------
                'Colar valor nas fórmulas
                '------------------------------------------------------------------------------------------------------------------
#Region "Colar valor nas fórmulas"
                aba = "1. COLAR BASE SAP"
                priLinha = 4
                ultLinha = -1
                priColuna = "G"
                ultColuna = "M"

                Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)

                aba = "1. COLAR BASE SAP"
                priLinha = 4
                priColuna = "G"

                Colar(BDVenda, aba, priLinha, priColuna)
#End Region

                .Save()

            End With

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Operacional FZ com Painel de Horas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Painel de Horas -> Painel Operacional FZ"
            processo = "Painel de Horas -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel Operacional FZ."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Painel Operacional FZ"
            lbNome = "lb_POpFz"
            lbId = "lbId_POpFz"

            AbrirArq(formNome, POpFz, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If POpFz Is Nothing Then

                lbNome_BDVenda.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atulizando Painel Operacional FZ com informações do Painel de Horas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do Disp. PDV do Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do Disp. PDV do Painel Operacional FZ"
            aba = "Disp. PDV"
            priLinha = 6
            priColuna = "A"
            ultColuna = "AH"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Painel de Horas"
            aba = "Colar em Disp PDV"
            priLinha = 8
            ultLinha = 0
            priColuna = "A"
            ultColuna = "AH"

            Copiar(PHoras, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados no Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na Disp. PDV"
            aba = "Disp. PDV"
            priColuna = "A"
            priLinha = 6

            Colar(POpFz, aba, priLinha, priColuna)
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

            PropFormatacao(POpFz, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "Disp. PDV"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "AJ"
            ultColuna2 = "AJ"

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados da Disp. PDV Consolid. do Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados da Disp. PDV Consolid. do Painel Operacional FZ"
            aba = "Disp. PDV Consolid."
            priLinha = 6
            priColuna = "B"
            ultColuna = "AJ"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Painel de Horas"
            aba = "Colar em Consolidado"
            priLinha = 7
            ultLinha = 0
            priColuna = "A"
            ultColuna = "AI"

            Copiar(PHoras, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados no Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na Disp. PDV Consolid."
            aba = "Disp. PDV Consolid."
            priColuna = "B"
            priLinha = 6

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar formatação condicional
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar formatação condicional na Disp. PDV Consolid."
            aba = "Disp. PDV Consolid."
            priColuna = "B"
            ultColuna = "AJ"
            priLinha = 6
            ultLinha = 0

            PropFormatacao(POpFz, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Painel de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Painel de Horas"
            salvar = "N"
            lbId = "lbId_PHoras"

            FecharWbks(formNome, PHoras, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional FZ."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpFz

                .Application.CalculateBeforeSave = False
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Operacional FZ com Base Dinâmica Vendas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Base Dinâmica Vendas -> Painel Operacional FZ"
            processo = "Base Dinâmica Vendas -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atulizando Painel Operacional FZ com informações do Base Dinâmica Vendas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 1. Base SAP do Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 1. Base SAP do Painel Operacional FZ"
            aba = "1. Base SAP"
            priLinha = 3
            priColuna = "A"
            ultColuna = "L"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Base Dinâmica Vendas"
            aba = "1. COLAR BASE SAP"
            priLinha = 4
            ultLinha = 0
            priColuna = "B"
            ultColuna = "M"

            Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados no Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na 1. Base SAP"
            aba = "1. Base SAP"
            priColuna = "A"
            priLinha = 3

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados da Disp. PDV Consolid. do Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados da 8. Base dinâmica vendas do Painel Operacional FZ"
            aba = "8. Base dinâmica vendas"
            priLinha = 2
            priColuna = "A"
            ultColuna = "N"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Base Dinâmica Vendas"
            aba = "3. ATUALIZAR TRATAMENTO"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "N"

            Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados no Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na 8. Base dinâmica vendas"
            aba = "8. Base dinâmica vendas"
            priColuna = "A"
            priLinha = 2

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Base Dinâmica Vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Base Dinâmica Vendas"
            salvar = "N"
            lbId = "lbId_BDVenda"

            FecharWbks(formNome, BDVenda, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional FZ."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpFz

                .Application.CalculateBeforeSave = False
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Ranking Diário
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Ranking Diário."
            form_Loading.ProgressBar1.PerformStep()

#Region "Abrir Ranking Diário"
            lbNome = "lb_RkDia"
            lbId = "lbId_RkDia"

            AbrirArq(formNome, RkDia, lbNome, lbId, verExcel)

            Threading.Thread.Sleep(5000)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If RkDia Is Nothing Then

                lbNome_RkDia.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizando Painel Operacional FZ com as informações do Ranking Diário.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Ranking Diário -> Painel Operacional FZ"
            processo = "Ranking Diário -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional FZ com as informações do Ranking Diário."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 6. Delivery Comentários - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 6. Delivery Comentários - Painel Operacional FZ"
            aba = "6. Delivery Comentários"
            priLinha = 2
            priColuna = "D"
            ultColuna = "K"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
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
            'colar os dados em 6. Delivery Comentários - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional FZ"
            aba = "6. Delivery Comentários"
            priColuna = "D"
            priLinha = 2

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

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
            'colar os dados em 6. Delivery Comentários - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional FZ"
            aba = "6. Delivery Comentários"
            priColuna = "E"
            priLinha = 2

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

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
            'colar os dados em 6. Delivery Comentários - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional FZ"
            aba = "6. Delivery Comentários"
            priColuna = "J"
            priLinha = 2

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

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
            'colar os dados em 6. Delivery Comentários - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6. Delivery Comentários - Painel Operacional FZ"
            aba = "6. Delivery Comentários"
            priColuna = "K"
            priLinha = 2

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "6. Delivery Comentários"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "D"
            ultColuna = "A"
            ultColuna2 = "C"

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 6.1 Delivery Tempo e Notas - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 6.1 Delivery Tempo e Notas - Painel Operacional FZ"
            aba = "6.1 Delivery Tempo e Notas"
            priRange = "F6:AJ22"
            With POpFz.Worksheets(aba)
                ultLinha = .range("A" & .rows.count).end(Excel.XlDirection.xlUp).row
            End With
            segRange = "A42:AJ" & ultLinha
            With POpFz.Worksheets(aba)
                ultLinha2 = .range("AL" & .rows.count).end(Excel.XlDirection.xlUp).row
            End With
            terRange = "AL6:BQ" & ultLinha2
            quaRange = "F29:AJ33"

            ApagarDlvTN(POpFz, aba, priRange, segRange, terRange, quaRange)
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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional FZ"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "F"
            priLinha = 6

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional FZ"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "F"
            priLinha = 29

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional FZ"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "A"
            priLinha = 42

            Colar(POpFz, aba, priLinha, priColuna)
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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK - 1ª PARTE
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "AL"
            priLinha = 6

            Colar(POpFz, aba, priLinha, priColuna)
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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK - 2ª PARTE
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "AN"
            priLinha = 6

            Colar(POpFz, aba, priLinha, priColuna)
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
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional FZ."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpFz

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
            'Atualizando Painel Operacional FZ com informações do B3 - Notas.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "B3 - Notas -> Painel Operacional FZ"
            processo = "B3 - Notas -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional FZ com informações do B3 - Notas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 2. GT Notas B3 - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 2. GT Notas B3 - Painel Operacional FZ"
            aba = "2. GT Notas B3"
            priLinha = 3
            priColuna = "A"
            ultColuna = "C"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
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
            'colar os dados em 2. GT Notas B3 - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 2. GT Notas B3 - Painel Operacional FZ"
            aba = "2. GT Notas B3"
            priColuna = "A"
            priLinha = 3

            Colar(POpFz, aba, priLinha, priColuna)
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

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
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
            'Atualizando Painel Operacional FZ com informações do B3 - Comentários.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "B3 - Comentários -> Painel Operacional FZ"
            processo = "B3 - Comentários -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional FZ com informações do B3 - Comentários."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 2.1 GT - Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 2.1 GT - Comentários - Painel Operacional FZ"
            aba = "2.1 GT - Comentários"
            priLinha = 3
            priColuna = "A"
            ultColuna = "D"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da B3 - Comentários - 1ª aba
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da B3 - Comentários - 1ª aba"
            aba = "1"
            priLinha = 3
            ultLinha = 0
            priColuna = "A"
            ultColuna = "D"

            Copiar(Coment, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados em 2.1 GT - Comentários - Painel Operacional BKB
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 2.1 GT - Comentários - Painel Operacional BKB"
            aba = "2.1 GT - Comentários"
            priColuna = "A"
            priLinha = 3

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "2.1 GT - Comentários"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "E"
            ultColuna2 = "I"

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
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
            'Atualizando Painel Operacional FZ com informações do AUV.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "AUV -> Painel Operacional FZ"
            processo = "AUV -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 3. AUV - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 3. AUV - Painel Operacional FZ"
            aba = "3. AUV"
            priLinha = 2
            priColuna = "A"
            ultColuna = "D"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional FZ com informações do AUV."
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
            'colar os dados em 3. AUV - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 3. AUV - Painel Operacional FZ"
            aba = "3. AUV"
            priColuna = "A"
            priLinha = 2

            Colar(POpFz, aba, priLinha, priColuna)
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

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
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
            'Atualizando Painel Operacional FZ com informações do Drive.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Drive -> Painel Operacional FZ"
            processo = "Drive -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional FZ com informações do Drive."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 4. Drive - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 4. Drive - Painel Operacional FZ"
            aba = "4. Drive"
            priLinha = 3
            priColuna = "C"
            ultColuna = "K"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
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
            'colar os dados em 4. Drive - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 4. Drive - Painel Operacional FZ"
            aba = "4. Drive"
            priColuna = "C"
            priLinha = 3

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "4. Drive"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "C"
            ultColuna = "A"
            ultColuna2 = "A"

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
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
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional FZ."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpFz

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
            'Atualizando Painel Operacional FZ com informações do TMA D-1.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "TMA D-1 -> Painel Operacional FZ"
            processo = "TMA D-1 -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional FZ com informações do TMA D-1."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 7. TMA D-1 - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 7. TMA D-1 - Painel Operacional FZ"
            aba = "7. TMA D-1"
            priLinha = 3
            priColuna = "A"
            ultColuna = "H"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
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
            'colar os dados em 7. TMA D-1 - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 7. TMA D-1 - Painel Operacional FZ"
            aba = "7. TMA D-1"
            priColuna = "A"
            priLinha = 3

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Apagar Zeros em 7. TMA D-1 - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apagar Zeros em 7. TMA D-1 - Painel Operacional FZ"
            aba = "7. TMA D-1"
            priColuna = "A"
            ultColuna = "B"
            ultColuna2 = "H"
            i = 3
            priLinha = 0
            ultLinha = 0

            ApagarZeros(POpFz, aba, i, priLinha, ultLinha, priColuna, ultColuna, ultColuna2)
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

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
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
            'Atualizando Painel Operacional FZ com informações do TMA MTD.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "TMA MTD -> Painel Operacional FZ"
            processo = "TMA MTD -> Painel Operacional FZ"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional FZ com informações do TMA MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 7. TMA MTD - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 7. TMA MTD - Painel Operacional FZ"
            aba = "7. TMA MTD"
            priLinha = 3
            priColuna = "A"
            ultColuna = "H"

            Apagar(POpFz, aba, priLinha, priColuna, ultColuna)
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
            'colar os dados em 7. TMA MTD - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 7. TMA MTD - Painel Operacional FZ"
            aba = "7. TMA MTD"
            priColuna = "A"
            priLinha = 3

            Colar(POpFz, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Apagar Zeros em 7. TMA MTD - Painel Operacional FZ
            '------------------------------------------------------------------------------------------------------------------
#Region "Apagar Zeros em 7. TMA MTD - Painel Operacional FZ"
            aba = "7. TMA MTD"
            priColuna = "A"
            ultColuna = "B"
            ultColuna2 = "H"
            i = 3
            priLinha = 0
            ultLinha = 0

            ApagarZeros(POpFz, aba, i, priLinha, ultLinha, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "7. TMA MTD"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "L"
            ultColuna2 = "T"

            PropFormulas(POpFz, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
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
            form_Loading.lbl_Loading.Text = "Salvando progresso e Abrindo o Painel Operacional FZ."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpFz

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

            painel = "Operacional FZ"

            ExportarHoras(dgvRelatorio, painel)

        End If
#End Region

    End Sub

End Class