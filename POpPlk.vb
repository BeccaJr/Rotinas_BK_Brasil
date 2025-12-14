Imports Excel = Microsoft.Office.Interop.Excel

Public Class form_POpPlk

    Private Sub form_POpPlk_Load(sender As Object, e As EventArgs) Handles Me.Load

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

    Private Sub btn_ArqPOpPlk_Click(sender As Object, e As EventArgs) Handles btn_ArqPOpPlk.Click

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
        nome = "POpPlk"

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

    Private Sub btn_ArqVendaDPlk_Click(sender As Object, e As EventArgs) Handles btn_ArqVendaDPlk.Click

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
        nome = "VendaDPlk"

        SalvarCaminho(formNome, caminho, nome)

    End Sub

    Private Sub btn_ArqVendaDMt_Click(sender As Object, e As EventArgs) Handles btn_ArqVendaDMt.Click

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
        nome = "VendaDMt"

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

        dgvRelatorio = Desenvolvedor.dgv_POpPlk 'DataGridView para salvar as infos de tempo de execução
        formNome = Me 'Nome do Formulário
        verExcel = Desenvolvedor.cbbox_Excel.Checked 'verificar status para ver o excel durante a execução do programa
        exportReport = Desenvolvedor.cbbox_Exportar.Checked 'verificar o status para saber se exporta o excel no fim da execução
#End Region

        '------------------------------------------------------------------------------------------------------------------
        'Atualizar Painel Operacional PLK
        '------------------------------------------------------------------------------------------------------------------
#Region "Atualizar Painel Operacional PLK"

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
            processoPainel = "Atualizar Painel Operacional PLK"
            inicioPainel = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Esconde o formulário Painel Operacional PLK, Abre o formulário de Loading e seta o valor da progress bar
            '------------------------------------------------------------------------------------------------------------------
            Me.Hide()
            form_Loading.ProgressBar1.Maximum = 24
            form_Loading.Show()

            '------------------------------------------------------------------------------------------------------------------
            'Se vazios é igual a 0 ele executa a rotina de colagem
            'Declaração das Panilhas Workbooks
            '------------------------------------------------------------------------------------------------------------------
#Region "Declaração de Planilhas"
            Dim POpPlk, TmaMtd, VendaDPlk, VendaDMt, RkDia, BDVenda, Auv As Excel.Workbook

            POpPlk = Nothing
            TmaMtd = Nothing
            VendaDPlk = Nothing
            VendaDMt = Nothing
            RkDia = Nothing
            BDVenda = Nothing
            Auv = Nothing
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Base Dinâmica Vendas com Venda D-1 PLK e Matias
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Venda D-1 PLK e Matias -> Base Dinâmica Vendas"
            processo = "Venda D-1 PLK e Matias -> Base Dinâmica Vendas"
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
            form_Loading.lbl_Loading.Text = "Abrindo Venda D-1 PLK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Venda PLK D-1
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Venda PLK D-1"
            lbNome = "lb_VendaDPlk"
            lbId = "lbId_VendaDPlk"

            AbrirArq(formNome, VendaDPlk, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If VendaDPlk Is Nothing Then

                lbNome_VendaDPlk.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Base Dinâmica Vendas com informações do Venda D-1 de PLK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Venda D-1 PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Venda D-1 PLK"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(VendaDPlk, aba, priLinha, ultLinha, priColuna, ultColuna)
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
            'Fechar Venda D-1 PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Venda D-1 PLK"
            salvar = "N"
            lbId = "lbId_VendaDPlk"

            FecharWbks(formNome, VendaDPlk, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Base Dinâmica Vendas com Venda D-1 Matias
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Venda D-1 Matias"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Venda D-1 Matias
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir Venda D-1 Matias"
            lbNome = "lb_VendaDMt"
            lbId = "lbId_VendaDMt"

            AbrirArq(formNome, VendaDMt, lbNome, lbId, verExcel)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If VendaDMt Is Nothing Then

                lbNome_VendaDMt.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atulizando Base Dinâmica Vendas com informações do Venda D-1 de Matias."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga coluna de Divisão da Base de Matias
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga coluna de Divisão da Base de Matias"
            aba = "1"
            priColuna = "D"
            ultColuna = "D"

            ApagarColuna(VendaDMt, aba, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da Venda D-1 Matias
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da Venda D-1 Matias"
            aba = "1"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(VendaDMt, aba, priLinha, ultLinha, priColuna, ultColuna)
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
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "1. COLAR BASE SAP"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "B"
            ultColuna = "O"
            ultColuna2 = "AC"

            PropFormulas(BDVenda, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar Venda D-1 Matias
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar Venda D-1 Matias"
            salvar = "N"
            lbId = "lbId_VendaDMt"

            FecharWbks(formNome, VendaDMt, salvar, lbId)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando dados do Base Dinâmica Vendas"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salva, Atualiza e Cola Valor nas fórmulas o Base Dinâmica Vendas
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
                priLinha = 2
                ultLinha = -1
                priColuna = "O"
                ultColuna = "AC"

                Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)

                aba = "1. COLAR BASE SAP"
                priLinha = 2
                priColuna = "O"

                Colar(BDVenda, aba, priLinha, priColuna)
#End Region

                .Save()

            End With

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Abrir Painel Operacional PLK
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo Painel Operacional PLK."
            form_Loading.ProgressBar1.PerformStep()

#Region "Abrir Painel Operacional PLK"
            lbNome = "lb_POpPlk"
            lbId = "lbId_POpPlk"

            AbrirArq(formNome, POpPlk, lbNome, lbId, verExcel)

            Threading.Thread.Sleep(5000)

            '------------------------------------------------------------------------------------------------------------------
            'Se tiver algum erro para abrir o arquivo ele sai da rotina
            '------------------------------------------------------------------------------------------------------------------
            If POpPlk Is Nothing Then

                lbNome_POpPlk.ForeColor = Color.Red
                GoTo fim

            End If
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Operacional PLK com Base Dinâmica Vendas
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Base Dinâmica Vendas -> Painel Operacional PLK"
            processo = "Base Dinâmica Vendas -> Painel Operacional PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizar Painel Operacional PLK com Base Dinâmica Vendas"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga informações do 1. Base Vendas PLKO
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga informações do 1. Base Vendas PLKO"
            aba = "1. Base Vendas PLKO"
            priLinha = 2
            priColuna = "A"
            ultColuna = "AB"

            Apagar(POpPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da 1. COLAR BASE SAP
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da 1. COLAR BASE SAP"
            aba = "1. COLAR BASE SAP"
            priLinha = 2
            ultLinha = 0
            priColuna = "B"
            ultColuna = "AC"

            Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na 1. Base Vendas PLKO
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na 1. Base Vendas PLKO"
            aba = "1. Base Vendas PLKO"
            priColuna = "A"
            priLinha = 2

            Colar(POpPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga informações do 4. Base dinâmica vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga informações do 4. Base dinâmica vendas"
            aba = "4. Base dinâmica vendas"
            priLinha = 2
            priColuna = "A"
            ultColuna = "M"

            Apagar(POpPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da 3. ATUALIZAR TRATAMENTO
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da 3. ATUALIZAR TRATAMENTO"
            aba = "3. ATUALIZAR TRATAMENTO"
            priLinha = 2
            ultLinha = 0
            priColuna = "A"
            ultColuna = "M"

            Copiar(BDVenda, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na 4. Base dinâmica vendas
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na 4. Base dinâmica vendas"
            aba = "4. Base dinâmica vendas"
            priColuna = "A"
            priLinha = 2

            Colar(POpPlk, aba, priLinha, priColuna)
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
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional PLK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpPlk

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
            'Atualizando Painel Operacional PLK com as informações do Ranking Diário.
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "Ranking Diário -> Painel Operacional PLK"
            processo = "Ranking Diário -> Painel Operacional PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional PLK com as informações do Ranking Diário."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 5. Delivery Comentários - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 5. Delivery Comentários - Painel Operacional PLK"
            aba = "5. Delivery Comentários"
            priLinha = 2
            priColuna = "D"
            ultColuna = "K"

            Apagar(POpPlk, aba, priLinha, priColuna, ultColuna)
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
            'colar os dados em 5. Delivery Comentários - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 5. Delivery Comentários - Painel Operacional PLK"
            aba = "5. Delivery Comentários"
            priColuna = "D"
            priLinha = 2

            Colar(POpPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Apagar coluna Mês (K) Aba Comentários - Ranking Diário
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
            'colar os dados em 5. Delivery Comentários - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 5. Delivery Comentários - Painel Operacional PLK"
            aba = "5. Delivery Comentários"
            priColuna = "E"
            priLinha = 2

            Colar(POpPlk, aba, priLinha, priColuna)
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
            'colar os dados em 5. Delivery Comentários - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 5. Delivery Comentários - Painel Operacional PLK"
            aba = "5. Delivery Comentários"
            priColuna = "J"
            priLinha = 2

            Colar(POpPlk, aba, priLinha, priColuna)
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
            'colar os dados em 5. Delivery Comentários - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 5. Delivery Comentários - Painel Operacional PLK"
            aba = "5. Delivery Comentários"
            priColuna = "K"
            priLinha = 2

            Colar(POpPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propagar as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propagar as fórmulas"
            aba = "5. Delivery Comentários"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "D"
            ultColuna = "A"
            ultColuna2 = "C"

            PropFormulas(POpPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga os dados do 6.1 Delivery Tempo e Notas - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga os dados do 6.1 Delivery Tempo e Notas - Painel Operacional PLK"
            aba = "6.1 Delivery Tempo e Notas"
            priRange = "F6:AJ22"
            With POpPlk.Worksheets(aba)
                ultLinha = .range("A" & .rows.count).end(Excel.XlDirection.xlUp).row
            End With
            segRange = "A42:AJ" & ultLinha
            With POpPlk.Worksheets(aba)
                ultLinha2 = .range("AL" & .rows.count).end(Excel.XlDirection.xlUp).row
            End With
            terRange = "AL6:BQ" & ultLinha2
            quaRange = "F29:AJ33"

            ApagarDlvTN(POpPlk, aba, priRange, segRange, terRange, quaRange)
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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "F"
            priLinha = 6

            Colar(POpPlk, aba, priLinha, priColuna)
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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "F"
            priLinha = 29

            Colar(POpPlk, aba, priLinha, priColuna)
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
            'colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados em 6.1 Delivery Tempo e Notas - Painel Operacional PLK"
            aba = "6.1 Delivery Tempo e Notas"
            priColuna = "A"
            priLinha = 42

            Colar(POpPlk, aba, priLinha, priColuna)
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

            Colar(POpPlk, aba, priLinha, priColuna)
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

            Colar(POpPlk, aba, priLinha, priColuna)
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
            form_Loading.lbl_Loading.Text = "Salvando progresso Painel Operacional PLK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpPlk

                .Application.CalculateBeforeSave = False
                .Save()

            End With

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Operacional PLK com informações do TMA MTD
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "TMA MTD -> Painel Operacional PLK"
            processo = "TMA MTD -> Painel Operacional PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo TMA MTD"
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
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional PLK com informações do TMA MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados 3. TME
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados 3. TME"
            aba = "3. TME"
            priLinha = 2
            priColuna = "D"
            ultColuna = "J"

            Apagar(POpPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da TMA MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da TMA MTD"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "G"

            Copiar(TmaMtd, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na 3. TME
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na 3. TME"
            aba = "3. TME"
            priColuna = "D"
            priLinha = 2

            Colar(POpPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "3. TME"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "D"
            ultColuna = "A"
            ultColuna2 = "B"

            PropFormulas(POpPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar TMA MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar TMA MTD"
            salvar = "N"
            lbId = "lbId_TmaMtd"

            FecharWbks(formNome, TmaMtd, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualizar Painel Operacional PLK com informações do AUV MTD
            'Seta as váriaveis para gravar Tempos
            '------------------------------------------------------------------------------------------------------------------
#Region "AUV MTD -> Painel Operacional PLK"
            processo = "AUV MTD -> Painel Operacional PLK"
            inicio = Now.ToLongTimeString

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Abrindo AUV MTD"
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Abrir AUV MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Abrir TMA MTD"
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
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Atualizando Painel Operacional PLK com informações do AUV MTD."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Apaga dados 2. AUV
            '------------------------------------------------------------------------------------------------------------------
#Region "Apaga dados 2. AUV"
            aba = "2. AUV"
            priLinha = 3
            priColuna = "A"
            ultColuna = "E"

            Apagar(POpPlk, aba, priLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Copia os dados da AUV MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Copia os dados da AUV MTD"
            aba = "1"
            priLinha = 4
            ultLinha = 0
            priColuna = "A"
            ultColuna = "E"

            Copiar(Auv, aba, priLinha, ultLinha, priColuna, ultColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'colar os dados na 2. AUV
            '------------------------------------------------------------------------------------------------------------------
#Region "colar os dados na 2. AUV"
            aba = "2. AUV"
            priColuna = "A"
            priLinha = 3

            Colar(POpPlk, aba, priLinha, priColuna)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Propaga as fórmulas
            '------------------------------------------------------------------------------------------------------------------
#Region "Propaga as fórmulas"
            aba = "2. AUV"
            ultLinha = 0
            ultLinha2 = 0
            priColuna = "A"
            ultColuna = "G"
            ultColuna2 = "J"

            PropFormulas(POpPlk, aba, ultLinha, ultLinha2, priColuna, ultColuna, ultColuna2)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Fechar TMA MTD
            '------------------------------------------------------------------------------------------------------------------
#Region "Fechar AUV MTD"
            salvar = "N"
            lbId = "lbId_Auv"

            FecharWbks(formNome, Auv, salvar, lbId)
#End Region

            fim = Now.ToLongTimeString
            total = fim.Subtract(inicio)

            GravarHoras(dgvRelatorio, processo, inicio, fim, total)
#End Region

            '------------------------------------------------------------------------------------------------------------------
            'Atualiza a barra de progresso e a label com as infos do progresso 
            '------------------------------------------------------------------------------------------------------------------
            form_Loading.lbl_Loading.Text = "Salvando progresso e Abrindo o Painel Operacional PLK."
            form_Loading.ProgressBar1.PerformStep()

            '------------------------------------------------------------------------------------------------------------------
            'Salvar progresso
            '------------------------------------------------------------------------------------------------------------------
            With POpPlk

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

            painel = "Operacional PLK"

            ExportarHoras(dgvRelatorio, painel)

        End If
#End Region

    End Sub

End Class