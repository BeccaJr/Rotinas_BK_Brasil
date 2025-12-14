Module Snippets

    '-----------------------------------------------------------------------------------------
    'Modificação dos Botões
    '-----------------------------------------------------------------------------------------
    'If nome <> "Painel operacional.xlsm" Then

    '    btn_ArqPOpBkb.Text = nome
    '    btn_ArqPOpBkb.BackColor = Color.Red

    'Else

    '    btn_ArqPOpBkb.Text = nome
    '    btn_ArqPOpBkb.BackColor = Color.LightGreen

    'End If

    '-----------------------------------------------------------------------------------------
    'Variáveis do Excel
    '-----------------------------------------------------------------------------------------
    'Public WS1 As Excel.Worksheet 'Aba da planilha
    'Public WB1 As Excel.Workbook 'Planilha

    '------------------------------------------------------------------------------------------------------------------
    'cola valores
    '------------------------------------------------------------------------------------------------------------------
    '.range("A2").PasteSpecial(Excel.XlPasteType.xlPasteValues)
    '.Range(pc & pl).PasteSpecial("HTML", False)

    '------------------------------------------------------------------------------------------------------------------
    'Coloca somente a ultima parte da string do caminho na label
    '------------------------------------------------------------------------------------------------------------------
    'lb_ArqNome1.Text = arquivo.Substring(arquivo.LastIndexOf("\") + 1)

    '------------------------------------------------------------------------------------------------------------------
    'Seleciona um range com somente as linhas visiveis em uma planilha
    '------------------------------------------------------------------------------------------------------------------
    '.range("A1:G" & ultLinha).SpecialCells(Excel.XlCellType.xlCellTypeVisible).select
    'cmv.Application.Selection.copy

    '------------------------------------------------------------------------------------------------------------------
    'Deixa as planilhas visiveis
    '------------------------------------------------------------------------------------------------------------------
    'histDisp.Application.Visible = True
    'pDisp.Application.Visible = True

    '------------------------------------------------------------------------------------------------------------------
    'Atualiza a barra de progresso e a label com as infos do progresso 
    '------------------------------------------------------------------------------------------------------------------
    'form_Loading.lbl_Loading.Text = "Atualizando Painel Disponibilidade."
    'form_Loading.ProgressBar1.PerformStep()

    '------------------------------------------------------------------------------------------------------------------
    'Finaliza processos do gerenciador de tarefas
    '------------------------------------------------------------------------------------------------------------------
    'Dim datestart As Date
    'Dim dateEnd As Date

    'datestart = Date.Now
    'Processo que deseja fechar
    'dateEnd = Date.Now

    'Dim xlp() As Process = Process.GetProcessesByName("EXCEL")

    'For Each Process As Process In xlp
    '    If Process.StartTime >= datestart And Process.StartTime <= dateEnd Then
    '        Process.Kill()
    '        Exit For
    '    End If
    'Next

    '------------------------------------------------------------------------------------------------------------------
    'Atribui um valor para o ExcelApp
    '------------------------------------------------------------------------------------------------------------------
    'EXCELAPP = wb.Application

    '------------------------------------------------------------------------------------------------------------------
    'Sub abrir arquivo antigo
    '------------------------------------------------------------------------------------------------------------------
    'Public Sub AbrirArq(ByVal f As Form, ByRef arq As Excel.Workbook, ByVal l As String, ByVal lId As String)

    '        '------------------------------------------------------------------------------------------------------------------
    '        'Correspondência de variáveis:
    '        'f = formNome
    '        'arq = variável que recebe o nome do excel que vai ser aberto
    '        'l = lbNome
    '        'lId = lbId
    '        '------------------------------------------------------------------------------------------------------------------

    '        Dim lbNome() As Control
    '        Dim lbId() As Control
    '        Dim nomeComp As String
    '        Dim ext As String

    '        '------------------------------------------------------------------------------------------------------------------
    '        'Acha a label com o ocaminho do arquivo para ser aberto
    '        '------------------------------------------------------------------------------------------------------------------
    '        lbNome = f.Controls.Find(l, True)
    '        lbId = f.Controls.Find(lId, True)

    '        nomeComp = lbNome(0).Text.Substring(lbNome(0).Text.LastIndexOf("\") + 1)
    '        ext = nomeComp.Substring(nomeComp.LastIndexOf(".") + 1)

    '        Try

    '            '------------------------------------------------------------------------------------------------------------------
    '            'Tenta abrir o excel do caminho selecionado 
    '            '------------------------------------------------------------------------------------------------------------------
    '            EXCELAPP = CreateObject("Excel.Application")

    '            '------------------------------------------------------------------------------------------------------------------
    '            'Salva o numero do ID do processo aberto
    '            '------------------------------------------------------------------------------------------------------------------
    '            Dim id() As Process = Process.GetProcessesByName("EXCEL")
    '            Dim tam As Integer
    '            Dim i As Integer

    '            tam = id.Length - 1

    '            With f

    '                i = 0

    '                Do While i <= tam

    '                    For Each control As Control In .Controls

    '                        If control.Name.StartsWith("lbId_") Then

    '                            If control.Text = CStr(id(i).Id) Then

    '                                GoTo proxId

    '                            End If

    '                        End If

    '                    Next

    '                    lbId(0).Text = CStr(id(i).Id)

    'proxId:
    '                    i += 1

    '                Loop

    '            End With

    '            If ext = "csv" Then

    '                arq = EXCELAPP.Workbooks.OpenXML(lbNome(0).Text, LoadOption:=Excel.XlXmlLoadOption.xlXmlLoadOpenXml)

    '            Else

    '                arq = EXCELAPP.Workbooks.Open(lbNome(0).Text) 'ABRIR EXCEL EXISTENTE

    '            End If

    '            arq.ReadOnlyRecommended = False

    '            Threading.Thread.Sleep(5000)

    '            EXCELAPP.Visible = True

    '        Catch ex As Exception

    '            '------------------------------------------------------------------------------------------------------------------
    '            'Se der erro, ele sinaliza que não foi possível abrir o excel, e finaliza o processo do excel no 
    '            'gerenciador de tarefas
    '            '------------------------------------------------------------------------------------------------------------------
    '            Dim arqNome As String

    '            arqNome = lbNome(0).Text.Substring(lbNome(0).Text.LastIndexOf("\") + 1)

    '            MsgBox("Não foi possível abrir o arquivo " & arqNome & Chr(13) &
    '                   "Por favor verifique se o arquivo existe no caminho escolhido.",
    '                   vbExclamation, "Erro ao abrir o Arquivo selecionado")

    '            lbNome(0).Text = ""

    '            Process.GetProcessById(CInt(lbId(0).Text)).Kill()

    '        End Try

    'End Sub

    '------------------------------------------------------------------------------------------------------------------
    'Sub antiga Texto para Coluna
    '------------------------------------------------------------------------------------------------------------------
    'Public Sub TxtCol(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pc As String)

    '    wb.Application.DisplayAlerts = False

    '    With wb.Worksheets(CInt(a))

    '        '------------------------------------------------------------------------------------------------------------------
    '        'seleciona a aba 1
    '        '------------------------------------------------------------------------------------------------------------------
    '        .select

    '        '------------------------------------------------------------------------------------------------------------------
    '        'texto para coluna
    '        '------------------------------------------------------------------------------------------------------------------
    '        .Columns(CInt(pc)).TextToColumns(
    '        Excel.XlColumnDataType.xlDMYFormat,
    '        DataType:=Excel.XlTextParsingType.xlDelimited,
    '        TextQualifier:=Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
    '        ConsecutiveDelimiter:=False,
    '        TAB:=False,
    '        Semicolon:=False,
    '        Comma:=True,
    '        Space:=False,
    '        Other:=False,
    '        TrailingMinusNumbers:=False)

    '    End With

    'End Sub

    '------------------------------------------------------------------------------------------------------------------
    'Esperar Atualização em segundo plano
    '------------------------------------------------------------------------------------------------------------------
    'Dim at As Boolean

    'at = .Worksheets(4).range("BASE_TRATADA_VENDA").ListObject.QueryTable.Refreshing

    'Do Until at = "False"

    '    Threading.Thread.Sleep(15000)

    '    at = .Worksheets(4).range("BASE_TRATADA_VENDA").ListObject.QueryTable.Refreshing

    'Loop

End Module
