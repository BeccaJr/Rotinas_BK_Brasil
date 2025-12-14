Imports Excel = Microsoft.Office.Interop.Excel

Public Class form_LaborPlk

    Private Sub form_LaborPlk_Load(sender As Object, e As EventArgs) Handles Me.Load

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

    Private Sub btn_ArqPLbPlk_Click(sender As Object, e As EventArgs) Handles btn_ArqPLbPlk.Click

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
        nome = "PLbPlk"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqRel2Hrs_Click(sender As Object, e As EventArgs) Handles btn_ArqRel2Hrs.Click

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
        nome = "Rel2Hrs"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqCadFunc_Click(sender As Object, e As EventArgs) Handles btn_ArqCadFunc.Click

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
        nome = "CadFunc"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqEstMarc_Click(sender As Object, e As EventArgs) Handles btn_ArqEstMarc.Click

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
        nome = "EstMarc"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqPtDia_Click(sender As Object, e As EventArgs) Handles btn_ArqPtDia.Click

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
        nome = "PtDia"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqRel11Hrs_Click(sender As Object, e As EventArgs) Handles btn_ArqRel11Hrs.Click

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
        nome = "Rel11Hrs"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqBHoras_Click(sender As Object, e As EventArgs) Handles btn_ArqBHoras.Click

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
        nome = "BHoras"

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

        dgvRelatorio = Desenvolvedor.dgv_PLbPlk 'DataGridView para salvar as infos de tempo de execução
        formNome = Me 'Nome do Formulário
        verExcel = Desenvolvedor.cbbox_Excel.Checked 'verificar status para ver o excel durante a execução do programa
        exportReport = Desenvolvedor.cbbox_Exportar.Checked 'verificar o status para saber se exporta o excel no fim da execução
#End Region

        '------------------------------------------------------------------------------------------------------------------
        'Atualizar Painel Labor PLK
        '------------------------------------------------------------------------------------------------------------------
#Region "Atualizar Painel Labor PLK"

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
            'Seta as váriaveis para gravar Tempo - Painel Labor PLK
            '------------------------------------------------------------------------------------------------------------------
            processoPainel = "Atualizar Painel Labor PLK"
            inicioPainel = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Esconde o formulário Painel Labor PLK, Abre o formulário de Loading e seta o valor da progress bar
            '------------------------------------------------------------------------------------------------------------------
            Me.Hide()
            form_Loading.ProgressBar1.Maximum = 14
            form_Loading.Show()

            '------------------------------------------------------------------------------------------------------------------
            'Se vazios é igual a 0 ele executa a rotina de colagem
            'Declaração das Panilhas Workbooks
            '------------------------------------------------------------------------------------------------------------------
#Region "Declaração de Planilhas"
            Dim PLbPlk, Rel2Hrs, CadFunc, EstMarc, PtDia, Rel11Hrs, BHoras As Excel.Workbook

            PLbPlk = Nothing
            Rel2Hrs = Nothing
            CadFunc = Nothing
            EstMarc = Nothing
            PtDia = Nothing
            Rel11Hrs = Nothing
            BHoras = Nothing
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel Labor PLK
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel Labor PLK."
            form_Loading.ProgressBar1.PerformStep()

#Region "Abrir Painel Labor PLK"
            lbNome = "lb_PLbPlk"
            lbId = "lbId_PLbPlk"

            AbrirArq(formNome, PLbPlk, lbNome, lbId, verExcel)

            Threading.Thread.Sleep(5000)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If PLbPlk Is Nothing Then

                lbNome_PLbPlk.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Labor PLK com informações do Relatório 2 Horas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Relatório 2 Horas -> Painel Labor PLK"
            processo = "Relatório 2 Horas -> Painel Labor PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Relatório 2 Horas"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Relatório 2 Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Relatório 2 Horas"
            lbNome = "lb_Rel2Hrs"
            lbId = "lbId_Rel2Hrs"

            AbrirArq(formNome, Rel2Hrs, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Rel2Hrs Is Nothing Then

                lbNome_Rel2Hrs.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Labor PLK com informações do Relatório 2 Horas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados MAIS DE 2H
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados MAIS DE 2H"
            aba = "MAIS DE 2H"
            priLinha = 2
            priColuna = "A"
            ultColuna = "N"

            Apagar(PLbPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Relatório 2 Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Relatório 2 Horas"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "N"

            Copiar(Rel2Hrs, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na MAIS DE 2H
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na MAIS DE 2H"
            aba = "MAIS DE 2H"
            priColuna = "A"
            priLinha = 2

            Colar(PLbPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "MAIS DE 2H"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "O"
            ultColuna2 = "R"

            PropFormulas(PLbPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Relatório 2 Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Relatório 2 Horas"
            salvar = "N"
            lbId = "lbId_Rel2Hrs"

            FecharWbks(formNome, Rel2Hrs, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Labor PLK com informações do Relatório 11 Horas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Relatório 11 Horas -> Painel Labor PLK"
            processo = "Relatório 11 Horas -> Painel Labor PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Relatório 11 Horas"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Relatório 11 Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Relatório 11 Horas"
            lbNome = "lb_Rel11Hrs"
            lbId = "lbId_Rel11Hrs"

            AbrirArq(formNome, Rel11Hrs, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Rel11Hrs Is Nothing Then

                lbNome_Rel11Hrs.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Labor PLK com informações do Relatório 11 Horas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados 11 HORAS DESCANSO
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados 11 HORAS DESCANSO"
            aba = "11 HORAS DESCANSO"
            priLinha = 2
            priColuna = "A"
            ultColuna = "S"

            Apagar(PLbPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Relatório 11 Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Relatório 11 Horas"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "S"

            Copiar(Rel11Hrs, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na 11 HORAS DESCANSO
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na 11 HORAS DESCANSO"
            aba = "11 HORAS DESCANSO"
            priColuna = "A"
            priLinha = 2

            Colar(PLbPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Relatório 11 Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Relatório 11 Horas"
            salvar = "N"
            lbId = "lbId_Rel11Hrs"

            FecharWbks(formNome, Rel11Hrs, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Labor PLK com informações do Resumo Banco de Horas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Resumo Banco de Horas -> Painel Labor PLK"
            processo = "Resumo Banco de Horas -> Painel Labor PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Resumo Banco de Horas"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Resumo Banco de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Resumo Banco de Horas"
            lbNome = "lb_BHoras"
            lbId = "lbId_BHoras"

            AbrirArq(formNome, BHoras, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If BHoras Is Nothing Then

                lbNome_BHoras.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Labor PLK com informações do Resumo Banco de Horas."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados RESUMO BANCO DE HORAS
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados RESUMO BANCO DE HORAS"
            aba = "RESUMO BANCO DE HORAS"
            priLinha = 2
            priColuna = "A"
            ultColuna = "L"

            Apagar(PLbPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Resumo Banco de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Resumo Banco de Horas"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "L"

            Copiar(BHoras, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na RESUMO BANCO DE HORAS
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na RESUMO BANCO DE HORAS"
            aba = "RESUMO BANCO DE HORAS"
            priColuna = "A"
            priLinha = 2

            Colar(PLbPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "RESUMO BANCO DE HORAS"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "M"
            ultColuna2 = "Q"

            PropFormulas(PLbPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Resumo Banco de Horas
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Resumo Banco de Horas"
            salvar = "N"
            lbId = "lbId_BHoras"

            FecharWbks(formNome, BHoras, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Labor PLK com informações do Cadastro Funcionário
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Cadastro Funcionário -> Painel Labor PLK"
            processo = "Cadastro Funcionário -> Painel Labor PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Cadastro Funcionário"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Cadastro Funcionário
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Cadastro Funcionário"
            lbNome = "lb_CadFunc"
            lbId = "lbId_CadFunc"

            AbrirArq(formNome, CadFunc, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If CadFunc Is Nothing Then

                lbNome_CadFunc.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Labor PLK com informações do Cadastro Funcionário."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados CAD_FUNC
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados CAD_FUNC"
            aba = "CAD_FUNC"
            priLinha = 2
            priColuna = "A"
            ultColuna = "AN"

            Apagar(PLbPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Cadastro Funcionário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Cadastro Funcionário"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "AN"

            Copiar(CadFunc, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na CAD_FUNC
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na CAD_FUNC"
            aba = "CAD_FUNC"
            priColuna = "A"
            priLinha = 2

            Colar(PLbPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "CAD_FUNC"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "AO"
            ultColuna2 = "CC"

            PropFormulas(PLbPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Cadastro Funcionário
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Cadastro Funcionário"
            salvar = "N"
            lbId = "lbId_CadFunc"

            FecharWbks(formNome, CadFunc, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Labor PLK com informações do Ponto Diário
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Ponto Diário -> Painel Labor PLK"
            processo = "Ponto Diário -> Painel Labor PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Ponto Diário"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Ponto Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Ponto Diário"
            lbNome = "lb_PtDia"
            lbId = "lbId_PtDia"

            AbrirArq(formNome, PtDia, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If PtDia Is Nothing Then

                lbNome_PtDia.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Labor PLK com informações do Ponto Diário."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados Ponto_Diario
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados Ponto_Diario"
            aba = "Ponto_Diario"
            priLinha = 2
            priColuna = "A"
            ultColuna = "O"

            Apagar(PLbPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Ponto Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Ponto Diário"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "O"

            Copiar(PtDia, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na Ponto_Diario
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na Ponto_Diario"
            aba = "Ponto_Diario"
            priColuna = "A"
            priLinha = 2

            Colar(PLbPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "Ponto_Diario"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "P"
            ultColuna2 = "P"

            PropFormulas(PLbPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Ponto Diário
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Ponto Diário"
            salvar = "N"
            lbId = "lbId_PtDia"

            FecharWbks(formNome, PtDia, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Labor PLK com informações do Estatística das Marcações
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Estatística das Marcações -> Painel Labor PLK"
            processo = "Estatística das Marcações -> Painel Labor PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Estatística das Marcações"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Estatística das Marcações
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Estatística das Marcações"
            lbNome = "lb_EstMarc"
            lbId = "lbId_EstMarc"

            AbrirArq(formNome, EstMarc, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If EstMarc Is Nothing Then

                lbNome_EstMarc.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Labor PLK com informações do Estatística das Marcações."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados ESTATISTICA DAS MARCACOES
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados ESTATISTICA DAS MARCACOES"
            aba = "ESTATISTICA DAS MARCACOES"
            priLinha = 2
            priColuna = "A"
            ultColuna = "E"

            Apagar(PLbPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados do Estatística das Marcações
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados do Estatística das Marcações"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "E"

            Copiar(EstMarc, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na ESTATISTICA DAS MARCACOES
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na ESTATISTICA DAS MARCACOES"
            aba = "ESTATISTICA DAS MARCACOES"
            priColuna = "A"
            priLinha = 2

            Colar(PLbPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "ESTATISTICA DAS MARCACOES"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "F"
            ultColuna2 = "J"

            PropFormulas(PLbPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Estatística das Marcações
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Estatística das Marcações"
            salvar = "N"
            lbId = "lbId_EstMarc"

            FecharWbks(formNome, EstMarc, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso e Abrindo o Painel Labor PLK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With PLbPlk

                .Application.CalculateBeforeSave = False
                .Save()

                '------------------------------------------------------------------------------------------------------------------
                'Final do código
                '------------------------------------------------------------------------------------------------------------------
                .Application.CalculateBeforeSave = True
                .Application.Visible = True

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos - Painel Labor PLK
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

            painel = "Labor PLK"

            ExportarHoras(dgvRelatorio, painel)

        End If
#End Region

    End Sub

End Class