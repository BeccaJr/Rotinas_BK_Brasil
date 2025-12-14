Imports Excel = Microsoft.Office.Interop.Excel

Public Class form_CmvPlk

    Private Sub form_CmvPlk_Load(sender As Object, e As EventArgs) Handles Me.Load

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

    Private Sub btn_ArqPCmvPlk_Click(sender As Object, e As EventArgs) Handles btn_ArqPCmvPlk.Click

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
        nome = "PCmvPlk"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqDC_Click(sender As Object, e As EventArgs) Handles btn_ArqDC.Click

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
        nome = "DC"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqDI_Click(sender As Object, e As EventArgs) Handles btn_ArqDI.Click

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
        nome = "DI"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqPControl_Click(sender As Object, e As EventArgs) Handles btn_ArqPControl.Click

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
        nome = "PControl"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqRelRest_Click(sender As Object, e As EventArgs) Handles btn_ArqRelRest.Click

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
        nome = "RelRest"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqDispMtd_Click(sender As Object, e As EventArgs) Handles btn_ArqDispMtd.Click

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
        nome = "DispMtd"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqRelLogCont_Click(sender As Object, e As EventArgs) Handles btn_ArqRelLogCont.Click
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
        nome = "RelLogCont"

        SalvarCaminho(formNome, caminho, nome)
    End Sub

    Private Sub btn_ArqRelReceb_Click(sender As Object, e As EventArgs) Handles btn_ArqRelReceb.Click

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
        nome = "RelReceb"

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

        dgvRelatorio = Desenvolvedor.dgv_PCmvPlk 'DataGridView para salvar as infos de tempo de execução
        formNome = Me 'Nome do Formulário
        verExcel = Desenvolvedor.cbbox_Excel.Checked 'verificar status para ver o excel durante a execução do programa
        exportReport = Desenvolvedor.cbbox_Exportar.Checked 'verificar o status para saber se exporta o excel no fim da execução
#End Region

        '------------------------------------------------------------------------------------------------------------------
        'Atualizar Painel CMV PLK
        '------------------------------------------------------------------------------------------------------------------
#Region "Atualizar Painel CMV PLK"

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
            'Seta as váriaveis para gravar Tempo - Painel CMV PLK
            '------------------------------------------------------------------------------------------------------------------
            processoPainel = "Atualizar Painel CMV PLK"
            inicioPainel = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Esconde o formulário Painel CMV PLK, Abre o formulário de Loading e seta o valor da progress bar
            '------------------------------------------------------------------------------------------------------------------
            Me.Hide()
            form_Loading.ProgressBar1.Maximum = 16
            form_Loading.Show()

            '------------------------------------------------------------------------------------------------------------------
            'Se vazios é igual a 0 ele executa a rotina de colagem
            'Declaração das Panilhas Workbooks
            '------------------------------------------------------------------------------------------------------------------
#Region "Declaração de Planilhas"
            Dim PCmvPlk, Dc, Di, PControl, RelRest, DispMtd, RelLogCont, RelReceb As Excel.Workbook

            PCmvPlk = Nothing
            Dc = Nothing
            Di = Nothing
            PControl = Nothing
            RelRest = Nothing
            DispMtd = Nothing
            RelLogCont = Nothing
            RelReceb = Nothing
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel CMV PLK
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel CMV PLK."
            form_Loading.ProgressBar1.PerformStep()

#Region "Abrir Painel CMV PLK"
            lbNome = "lb_PCmvPlk"
            lbId = "lbId_PCmvPlk"

            AbrirArq(formNome, PCmvPlk, lbNome, lbId, verExcel)

            Threading.Thread.Sleep(5000)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If PCmvPlk Is Nothing Then

                lbNome_PCmvPlk.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel CMV PLK com informações do Relatório Restaurante
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Relatório Restaurante -> Painel CMV PLK"
            processo = "Relatório Restaurante -> Painel CMV PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Relatório Restaurante"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Relatório Restaurante
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Relatório Restaurante"
            lbNome = "lb_RelRest"
            lbId = "lbId_RelRest"

            AbrirArq(formNome, RelRest, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If RelRest Is Nothing Then

                lbNome_RelRest.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel CMV PLK com informações do Relatório Restaurante."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados RELATÓRIO DE VENDA
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados RELATÓRIO DE VENDA"
            aba = "RELATÓRIO DE VENDA"
            priLinha = 2
            priColuna = "A"
            ultColuna = "M"

            Apagar(PCmvPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Relatório Restaurante
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Relatório Restaurante"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(RelRest, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na RELATÓRIO DE VENDA
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na RELATÓRIO DE VENDA"
            aba = "RELATÓRIO DE VENDA"
            priColuna = "A"
            priLinha = 2

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "RELATÓRIO DE VENDA"
            ultLinha = 0
            ultLinha2 = 2
            priColuna = "A"
            ultColuna = "P"
            ultColuna2 = "O"

            PropFormulas(PCmvPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Relatório Restaurante
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Relatório Restaurante"
            salvar = "N"
            lbId = "lbId_RelRest"

            FecharWbks(formNome, RelRest, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel CMV PLK com informações do Relatório Recebimentos
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Relatório Recebimentos -> Painel CMV PLK"
            processo = "Relatório Recebimentos -> Painel CMV PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Relatório Recebimentos"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Relatório Recebimentos
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Relatório Recebimentos"
            lbNome = "lb_RelReceb"
            lbId = "lbId_RelReceb"

            AbrirArq(formNome, RelReceb, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If RelReceb Is Nothing Then

                lbNome_RelReceb.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel CMV PLK com informações do Relatório Recebimentos."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados PENDÊNCIAS NFs
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados PENDÊNCIAS NFs"
            aba = "PENDÊNCIAS NFs"
            priLinha = 2
            priColuna = "A"
            ultColuna = "R"

            Apagar(PCmvPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Relatório Recebimentos
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Relatório Recebimentos"
            aba = "1"
            priLinha = 8
            ultLinha = 0
            priColuna = "A"
            ultColuna = "R"

            Copiar(RelReceb, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na PENDÊNCIAS NFs
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na PENDÊNCIAS NFs"
            aba = "PENDÊNCIAS NFs"
            priColuna = "A"
            priLinha = 2

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "PENDÊNCIAS NFs"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "T"
            ultColuna2 = "AA"

            PropFormulas(PCmvPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Relatório Recebimentos
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Relatório Recebimentos"
            salvar = "N"
            lbId = "lbId_RelReceb"

            FecharWbks(formNome, RelReceb, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel CMV PLK com informações do Relatório DC DI
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Relatório DC DI -> Painel CMV PLK"
            processo = "Relatório DC DI -> Painel CMV PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Relatório DC"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Relatório DC
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Relatório DC"
            lbNome = "lb_DC"
            lbId = "lbId_DC"

            AbrirArq(formNome, Dc, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Dc Is Nothing Then

                lbNome_DC.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel CMV PLK com informações do Relatório DC DI."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados RELATÓRIO DCDI
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados RELATÓRIO DCDI"
            aba = "RELATÓRIO DCDI"
            priLinha = 2
            priColuna = "A"
            ultColuna = "M"

            Apagar(PCmvPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Relatório DC
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Relatório DC"
            aba = "2"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(Dc, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na RELATÓRIO DCDI
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na RELATÓRIO DCDI"
            aba = "RELATÓRIO DCDI"
            priColuna = "A"
            priLinha = 2

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Relatório DC
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Relatório DC"
            salvar = "N"
            lbId = "lbId_DC"

            FecharWbks(formNome, Dc, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Relatório DI"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Relatório DI
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Relatório DI"
            lbNome = "lb_DI"
            lbId = "lbId_DI"

            AbrirArq(formNome, Di, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If Di Is Nothing Then

                lbNome_DI.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel CMV PLK com informações do Relatório DI."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Relatório DI
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Relatório DI"
            aba = "2"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(Di, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na RELATÓRIO DCDI
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na RELATÓRIO DCDI"
            aba = "RELATÓRIO DCDI"
            priColuna = "A"
            priLinha = PCmvPlk.Worksheets(aba).range(priColuna & PCmvPlk.Worksheets(aba).rows.count).end(Excel.XlDirection.xlUp).row + 1

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Relatório DI
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Relatório DI"
            salvar = "N"
            lbId = "lbId_DI"

            FecharWbks(formNome, Di, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "RELATÓRIO DCDI"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "O"
            ultColuna2 = "T"

            PropFormulas(PCmvPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel CMV PLK com Log Contagem
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Log Contagem -> Painel CMV PLK"
            processo = "Log Contagem -> Painel CMV PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Log Contagem."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Log Contagem
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Log Contagem"
            lbNome = "lb_RelLogCont"
            lbId = "lbId_RelLogCont"

            AbrirArq(formNome, RelLogCont, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If RelLogCont Is Nothing Then

                lbNome_RelLogCont.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel CMV PLK com informações do Log Contagem."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados HORA DA CONTAGEM
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados HORA DA CONTAGEM"
            aba = "HORA DA CONTAGEM"
            priLinha = 2
            priColuna = "A"
            ultColuna = "H"

            Apagar(PCmvPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Log Contagem
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Log Contagem"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "H"

            Copiar(RelLogCont, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na HORA DA CONTAGEM
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na HORA DA CONTAGEM"
            aba = "HORA DA CONTAGEM"
            priColuna = "A"
            priLinha = 2

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Log Contagem
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Log Contagem"
            salvar = "N"
            lbId = "lbId_RelLogCont"

            FecharWbks(formNome, RelLogCont, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "HORA DA CONTAGEM"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "J"
            ultColuna2 = "N"

            PropFormulas(PCmvPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel CMV PLK com informações do Painel de Controle
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Painel de Controle -> Painel CMV PLK"
            processo = "Painel de Controle -> Painel CMV PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel de Controle"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel de Controle
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Painel de Controle"
            lbNome = "lb_PControl"
            lbId = "lbId_PControl"

            AbrirArq(formNome, PControl, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If PControl Is Nothing Then

                lbNome_PControl.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel CMV PLK com informações do Painel de Controle."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados ROTINAS
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados ROTINAS"
            aba = "ROTINAS"
            priLinha = 2
            priColuna = "A"
            ultColuna = "G"

            Apagar(PCmvPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Aba Break
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Aba Break"
            aba = "1"
            priLinha = 10
            ultLinha = 0
            priColuna = "B"
            ultColuna = "H"

            Copiar(PControl, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na Rotinas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na ROTINAS"
            aba = "ROTINAS"
            priColuna = "A"
            priLinha = 2

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Aba DC
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Aba DC"
            aba = "2"
            priLinha = 10
            ultLinha = 0
            priColuna = "B"
            ultColuna = "H"

            Copiar(PControl, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na Rotinas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na ROTINAS"
            aba = "ROTINAS"
            priColuna = "A"
            priLinha = PCmvPlk.Worksheets(aba).range(priColuna & PCmvPlk.Worksheets(aba).rows.count).end(Excel.XlDirection.xlUp).row + 1

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Aba DIC
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Aba DIC"
            aba = "3"
            priLinha = 10
            ultLinha = 0
            priColuna = "B"
            ultColuna = "H"

            Copiar(PControl, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na Rotinas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na ROTINAS"
            aba = "ROTINAS"
            priColuna = "A"
            priLinha = PCmvPlk.Worksheets(aba).range(priColuna & PCmvPlk.Worksheets(aba).rows.count).end(Excel.XlDirection.xlUp).row + 1

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Aba Contagem
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Aba Contagem"
            aba = "4"
            priLinha = 10
            ultLinha = 0
            priColuna = "B"
            ultColuna = "H"

            Copiar(PControl, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na Rotinas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na ROTINAS"
            aba = "ROTINAS"
            priColuna = "A"
            priLinha = PCmvPlk.Worksheets(aba).range(priColuna & PCmvPlk.Worksheets(aba).rows.count).end(Excel.XlDirection.xlUp).row + 1

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "ROTINAS"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "I"
            ultColuna2 = "N"

            PropFormulas(PCmvPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Painel de Controle
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Painel de Controle"
            salvar = "N"
            lbId = "lbId_PControl"

            FecharWbks(formNome, PControl, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel CMV PLK com Dispersão MTD
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Dispersão MTD -> Painel CMV PLK"
            processo = "Dispersão MTD -> Painel CMV PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Dispersão MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Dispersão MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Dispersão MTD"
            lbNome = "lb_DispMtd"
            lbId = "lbId_DispMtd"

            AbrirArq(formNome, DispMtd, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If DispMtd Is Nothing Then

                lbNome_DispMtd.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel CMV PLK com informações do Dispersão MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados BASE PLK OFFICE
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados BASE PLK OFFICE"
            aba = "BASE PLK OFFICE"
            priLinha = 2
            priColuna = "A"
            ultColuna = "X"

            Apagar(PCmvPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Dispersão MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Dispersão MTD"
            aba = "1"
            priLinha = 11
            ultLinha = 0
            priColuna = "B"
            ultColuna = "Y"

            Copiar(DispMtd, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na BASE PLK OFFICE
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na BASE PLK OFFICE"
            aba = "BASE PLK OFFICE"
            priColuna = "A"
            priLinha = 2

            Colar(PCmvPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Dispersão MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Dispersão MTD"
            salvar = "N"
            lbId = "lbId_DispMtd"

            FecharWbks(formNome, DispMtd, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "BASE PLK OFFICE"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "Z"
            ultColuna2 = "AK"

            PropFormulas(PCmvPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso e Abrindo o Painel CMV PLK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With PCmvPlk

                .Application.CalculateBeforeSave = False
                .Save()

                '------------------------------------------------------------------------------------------------------------------
                'Final do código
                '------------------------------------------------------------------------------------------------------------------
                .Application.CalculateBeforeSave = True
                .Application.Visible = True

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Seta as váriaveis para gravar Tempos - Painel CMV PLK
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

            painel = "CMV PLK"

            ExportarHoras(dgvRelatorio, painel)

        End If
#End Region

    End Sub

End Class