Imports Excel = Microsoft.Office.Interop.Excel

Module Arquivos

    Public EXCELAPP As Excel.Application = Nothing

    Public Sub SelecionarArq(ByVal ofd As OpenFileDialog, ByRef c As String)

refazer:
        '------------------------------------------------------------------------------------------------------------------
        'Abre caixa de Dialogo para selecionar arquivo
        '------------------------------------------------------------------------------------------------------------------
        Dim dr As DialogResult = ofd.ShowDialog()

        If dr = System.Windows.Forms.DialogResult.OK Then

            For Each arquivo As [String] In ofd.FileNames

                '------------------------------------------------------------------------------------------------------------------
                'Seta X com o caminho do Arquivo
                '------------------------------------------------------------------------------------------------------------------
                c = arquivo

            Next

        Else

            titulo = "Erro na Seleção"
            aviso = "Nenhum arquivo foi selecionado. Deseja selecionar agora?"
            tipo = "Alerta"
            escolha = "SimNao"

            Mensagem(aviso, tipo, titulo, escolha)

            Select Case escolha
                Case "Sim"
                    GoTo refazer

            End Select

        End If

    End Sub

    Public Sub SalvarCaminho(ByVal f As Form, ByVal c As String, ByVal n As String)

        Dim BtnNome() As Control
        Dim lb() As Control
        Dim lbNome() As Control

        BtnNome = f.Controls.Find("btn_Arq" & n, True)
        lb = f.Controls.Find("lb_" & n, True)
        lbNome = f.Controls.Find("lbNome_" & n, True)

        '------------------------------------------------------------------------------------------------------------------
        'Salva Caminho do Arquivo Selecionado na Label correspondente
        '------------------------------------------------------------------------------------------------------------------

        If c <> "" Then

            lb(0).Text = c
            lbNome(0).Text = c.Substring(c.LastIndexOf("\") + 1)
            lbNome(0).ForeColor = Color.Black

        Else

            lbNome(0).Text = "Selecione um Arquivo"
            lbNome(0).ForeColor = Color.Red

        End If

    End Sub

    Public Sub AbrirArq(ByVal f As Form, ByRef arq As Excel.Workbook, ByVal l As String, ByVal lId As String, ByVal vE As Boolean)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'f = formNome
        'arq = variável que recebe o nome do excel que vai ser aberto
        'l = lbNome
        'lId = lbId
        'vE = verExcel
        '------------------------------------------------------------------------------------------------------------------

        Dim lbNome() As Control
        Dim lbId() As Control
        Dim nomeComp As String
        Dim ext As String

        '------------------------------------------------------------------------------------------------------------------
        'Acha a label com o ocaminho do arquivo para ser aberto
        '------------------------------------------------------------------------------------------------------------------
        lbNome = f.Controls.Find(l, True)
        lbId = f.Controls.Find(lId, True)

        nomeComp = lbNome(0).Text.Substring(lbNome(0).Text.LastIndexOf("\") + 1)
        ext = nomeComp.Substring(nomeComp.LastIndexOf(".") + 1)

        If ext = "csv" Then

            AbrirCSV(f, arq, l, lId, vE)

        Else

            AbrirXML(f, arq, l, lId, vE)

        End If

    End Sub

    Public Sub AbrirXML(ByVal f As Form, ByRef arq As Excel.Workbook, ByVal l As String, ByVal lId As String, ByVal vE As Boolean)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'f = formNome
        'arq = variável que recebe o nome do excel que vai ser aberto
        'l = lbNome
        'lId = lbId
        '------------------------------------------------------------------------------------------------------------------

        Dim lbNome() As Control
        Dim lbId() As Control

        '------------------------------------------------------------------------------------------------------------------
        'Acha a label com o ocaminho do arquivo para ser aberto
        '------------------------------------------------------------------------------------------------------------------
        lbNome = f.Controls.Find(l, True)
        lbId = f.Controls.Find(lId, True)

        Try

            '------------------------------------------------------------------------------------------------------------------
            'Tenta abrir o excel do caminho selecionado 
            '------------------------------------------------------------------------------------------------------------------
            EXCELAPP = CreateObject("Excel.Application")

            '------------------------------------------------------------------------------------------------------------------
            'Salva o numero do ID do processo aberto
            '------------------------------------------------------------------------------------------------------------------
            Dim id() As Process = Process.GetProcessesByName("EXCEL")
            Dim tam As Integer
            Dim i As Integer

            tam = id.Length - 1

            With f

                i = 0

                Do While i <= tam

                    For Each control As Control In .Controls

                        If control.Name.StartsWith("lbId_") Then

                            If control.Text = CStr(id(i).Id) Then

                                GoTo proxId

                            End If

                        End If

                    Next

                    lbId(0).Text = CStr(id(i).Id)

proxId:
                    i += 1

                Loop

            End With

            arq = EXCELAPP.Workbooks.Open(lbNome(0).Text) 'ABRIR EXCEL EXISTENTE

            arq.ReadOnlyRecommended = False

            Threading.Thread.Sleep(5000)

            If vE = True Then

                arq.Application.Visible = True

            Else

                arq.Application.Visible = False

            End If

        Catch ex As Exception

            '------------------------------------------------------------------------------------------------------------------
            'Se der erro, ele sinaliza que não foi possível abrir o excel, e finaliza o processo do excel no 
            'gerenciador de tarefas
            '------------------------------------------------------------------------------------------------------------------
            Dim arqNome As String

            arqNome = lbNome(0).Text.Substring(lbNome(0).Text.LastIndexOf("\") + 1)

            MsgBox("Não foi possível abrir o arquivo " & arqNome & Chr(13) &
                   "Por favor verifique se o arquivo existe no caminho escolhido.",
                   vbExclamation, "Erro ao abrir o Arquivo selecionado")

            lbNome(0).Text = ""

            Process.GetProcessById(CInt(lbId(0).Text)).Kill()

        End Try

    End Sub

    Public Sub AbrirCSV(ByVal f As Form, ByRef arq As Excel.Workbook, ByVal l As String, ByVal lId As String, ByVal vE As Boolean)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'f = formNome
        'arq = variável que recebe o nome do excel que vai ser aberto
        'l = lbNome
        'lId = lbId
        'vE = verExcel
        '------------------------------------------------------------------------------------------------------------------

        Dim lbNome() As Control
        Dim lbId() As Control
        Dim mFormula As String
        Dim NovoExcel As New Excel.Application

        '------------------------------------------------------------------------------------------------------------------
        'Acha a label com o ocaminho do arquivo para ser aberto
        '------------------------------------------------------------------------------------------------------------------
        lbNome = f.Controls.Find(l, True)
        lbId = f.Controls.Find(lId, True)

        Try

            '------------------------------------------------------------------------------------------------------------------
            'Tenta abrir o excel do caminho selecionado 
            '------------------------------------------------------------------------------------------------------------------
            arq = NovoExcel.Workbooks.Add

            '------------------------------------------------------------------------------------------------------------------
            'Salva o numero do ID do processo aberto
            '------------------------------------------------------------------------------------------------------------------
            Dim id() As Process = Process.GetProcessesByName("EXCEL")
            Dim tam As Integer
            Dim i As Integer

            tam = id.Length - 1

            With f

                i = 0

                Do While i <= tam

                    For Each control As Control In .Controls

                        If control.Name.StartsWith("lbId_") Then

                            If control.Text = CStr(id(i).Id) Then

                                GoTo proxId

                            End If

                        End If

                    Next

                    lbId(0).Text = CStr(id(i).Id)

proxId:
                    i += 1

                Loop

            End With

            If f.Name = "form_POpPlk" And lId = "lbId_TmaMtd" Then

                mFormula = "let Source = Csv.Document(File.Contents(""" & lbNome(0).Text & """),
                    [Delimiter="",""]), #""Valor Substituído"" = Table.ReplaceValue(Source,""-"",""@"",Replacer.ReplaceText,{""Column3""}) in #""Valor Substituído"""

            Else

                mFormula = "let Source = Csv.Document(File.Contents(""" & lbNome(0).Text & """),
                    [Delimiter="",""]) in Source"

            End If

            arq.Queries.Add("query1", mFormula)

            Threading.Thread.Sleep(2500)

            With arq.ActiveSheet.listobjects.add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""query1"";Extended Properties=""""", Destination:=arq.ActiveSheet.Range("$A$1")).QueryTable

                .CommandType = Excel.XlCmdType.xlCmdSql
                .CommandText = "SELECT * FROM [query1]"
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = "query1"
                .Refresh(False)

            End With

            Threading.Thread.Sleep(1500)

            With arq.ActiveSheet.range("query1")

                .numberformat = "General"
                Threading.Thread.Sleep(2500)
                .formulaLocal = .value

            End With

            If vE = True Then

                arq.Application.Visible = True

            Else

                arq.Application.Visible = False

            End If

        Catch ex As Exception

            '------------------------------------------------------------------------------------------------------------------
            'Se der erro, ele sinaliza que não foi possível abrir o excel, e finaliza o processo do excel no 
            'gerenciador de tarefas
            '------------------------------------------------------------------------------------------------------------------
            Dim arqNome As String

            arqNome = lbNome(0).Text.Substring(lbNome(0).Text.LastIndexOf("\") + 1)

            MsgBox("Não foi possível abrir o arquivo " & arqNome & Chr(13) &
                   "Por favor verifique se o arquivo existe no caminho escolhido.",
                   vbExclamation, "Erro ao abrir o Arquivo selecionado")

            lbNome(0).Text = ""

            Process.GetProcessById(CInt(lbId(0).Text)).Kill()

        End Try

    End Sub

    Public Sub ChecarCaminhos(ByVal f As Form, ByRef v As Integer, ByRef n As ArrayList)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'f = formNome
        'v = vazios
        'n = nome
        '------------------------------------------------------------------------------------------------------------------

        With f

            '------------------------------------------------------------------------------------------------------------------
            'Verifica cada Label do formulário atual (pode variar de acordo com o formulário aberto)
            'Se a label iniciar o nome com "lb_" ele verifica se ela está vazia, se sim adiciona 1 na variável Vazios
            '------------------------------------------------------------------------------------------------------------------
            For Each control As Control In .Controls

                If control.Name.StartsWith("lb_") Then

                    If control.Text Like "lb_*" Then

                        n.Add(control.Name)
                        v += 1

                    End If

                End If

            Next

        End With

    End Sub

    Public Sub MostraVazios(ByVal f As Form, ByRef v As Integer, ByRef n As ArrayList, ByRef ns As ArrayList)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'f = formNome
        'v = vazios
        'n = nome
        'ns = nomes
        '------------------------------------------------------------------------------------------------------------------

        '------------------------------------------------------------------------------------------------------------------
        'Entra nessa função somente se o Vazios for diferente de 0
        'Ela pega o texto que está no Botão referente à base que falta selecionar e salva no array ns
        '------------------------------------------------------------------------------------------------------------------
        Dim i As Integer = 0

        Do While i < v

            Dim nomeBtn As String
            Dim lbNome As String
            Dim btn() As Control
            Dim lb() As Control

            nomeBtn = n(i).substring(n(i).lastindexof("_") + 1)
            lbNome = n(i).substring(n(i).lastindexof("_") + 1)

            btn = f.Controls.Find("btn_Arq" & nomeBtn, True)
            lb = f.Controls.Find("lbNome_" & lbNome, True)

            lb(0).Text = "Selecione o arquivo " & btn(0).Text
            lb(0).ForeColor = Color.Red

            ns.Add(btn(0).Text)

            i += 1

        Loop

    End Sub

End Module