Imports Excel = Microsoft.Office.Interop.Excel

Module Processos

    Public EXCELAPP As Excel.Application

    Public Sub Apagar(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pl As Long, ByVal pc As String, ByVal uc As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'pl = priLinha
        'pc = priColuna
        'uc = ultColuna
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        With wb.Worksheets(a)

            Dim ul As Long

            .select

            ul = .range(pc & .rows.count).end(Excel.XlDirection.xlUp).row

            .Range(pc & pl & ":" & uc & ul).ClearContents

        End With

    End Sub

    Public Sub ApagarZeros(ByVal wb As Excel.Workbook, ByVal a As String, ByVal i As Long, ByVal pl As Long, ByVal ul As Long, ByVal pc As String, ByVal uc As String, ByVal uc2 As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'i = i
        'pl = priLinha
        'ul = ultLinha
        'pc = priColuna
        'uc = ultColuna
        'uc2 = ultColuna2
        '------------------------------------------------------------------------------------------------------------------

        With wb.Worksheets(a)

            Do While .cells(i, uc).value <> 0

                i += 1

            Loop

            pl = i
            ul = .range(uc & .rows.count).end(Excel.XlDirection.xlUp).row

            .Range(pc & pl & ":" & uc2 & ul).ClearContents

        End With

    End Sub

    Public Sub ApagarColuna(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pc As String, ByVal uc As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'pc = priColuna
        'uc = ultColuna
        '------------------------------------------------------------------------------------------------------------------

        Dim w

        If IsNumeric(a) Then

            Dim ai As Integer

            ai = CInt(a)

            w = wb.Worksheets(ai)

        Else

            w = wb.Worksheets(a)

        End If

        Threading.Thread.Sleep(2500)

        With w

            .columns(pc & ":" & uc).Delete

        End With

    End Sub

    Public Sub ApagarDlvTN(ByVal wb As Excel.Workbook, ByVal a As String, ByVal p As String, ByVal s As String, ByVal t As String, ByVal q As String)

        '------------------------------------------------------------------------------------------------------------------
        'Apagar informações do delivery tempo e notas
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'p = priRange
        's = segRange
        't = terRange
        'q = quaRange
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        With wb.Worksheets(a)

            .select

            If q = "" Then

                .Range(p).ClearContents
                .Range(s).ClearContents
                .Range(t).ClearContents

            Else

                .Range(p).ClearContents
                .Range(s).ClearContents
                .Range(t).ClearContents
                .Range(q).ClearContents

            End If

        End With

    End Sub

    Public Sub Copiar(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pl As Long, ByVal ul As Long, ByVal pc As String, ByVal uc As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'pl = priLinha
        'ul = ultLinha
        'pc = priColuna
        'uc = ultColuna
        '------------------------------------------------------------------------------------------------------------------

        Clipboard.Clear()

        Threading.Thread.Sleep(2500)

        Dim w

        If IsNumeric(a) Then

            Dim ai As Integer

            ai = CInt(a)

            w = wb.Worksheets(ai)

        Else

            w = wb.Worksheets(a)

        End If

        With w

            .Select

            If ul > 0 Then

                .Range(pc & pl & ":" & uc & ul).Copy

            ElseIf ul = 0 Then

                ul = .range(pc & .rows.count).end(Excel.XlDirection.xlUp).row

                .Range(pc & pl & ":" & uc & ul).Copy

            ElseIf ul < 0 Then

                ul = .range(pc & .rows.count).end(Excel.XlDirection.xlUp).row

                .Range(pc & pl & ":" & uc & ul - 1).Copy

            End If

        End With

    End Sub

    Public Sub Colar(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pl As Long, ByVal pc As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'pl = priLinha
        'pc = priColuna
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        With wb.Worksheets(a)

            .select

            .Range(pc & pl).PasteSpecial(Excel.XlPasteType.xlPasteValues)

        End With

    End Sub

    Public Sub ColarTexto(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pl As Long, ByVal pc As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'pl = priLinha
        'pc = priColuna
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        With wb.Worksheets(a)

            .select

            .Range(pc & pl).PasteSpecial("Texto", False)

        End With

    End Sub

    Public Sub ColarCriterioDt(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pl As Long, ByVal ul As Long, ByVal dtP As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'pl = priLinha
        'ul = ultLinha
        'pc = priColuna
        'dtP = dtPainel
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        Dim c As Long
        Dim cL As String
        Dim colL As String()

        c = 1

        wb.Application.DisplayAlerts = False

        With wb.Worksheets(a)

            .select

            Do Until CStr(.cells(pl, c).value) = dtP

                c += 1

            Loop

            cL = .Cells(, c).Address
            colL = cL.Split(New Char() {"$"c}) 'define a letra da coluna

            .Range(colL(1) & ul).PasteSpecial("HTML", False)

            .Calculate

        End With

        wb.Application.DisplayAlerts = True

    End Sub

    Public Sub PropFormatacao(ByVal wb As Excel.Workbook, ByVal a As String, ByVal pl As Long, ByVal ul As Long, ByVal pc As String, ByVal uc As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'pl = priLinha
        'ul = ultLinha
        'pc = priColuna
        'uc = ultColuna
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        With wb.Worksheets(a)

            .select

            '------------------------------------------------------------------------------------------------------------------
            'acha a ultima linha da planilha atual para ajustar as fórmulas com o tamanho dos dados
            '------------------------------------------------------------------------------------------------------------------
            ul = .range(pc & .rows.count).end(Excel.XlDirection.xlUp).row

            .range(pc & pl & " : " & uc & pl).Copy

            .Range(pc & pl & " : " & uc & ul).PasteSpecial(Excel.XlPasteType.xlPasteFormats)

        End With

    End Sub

    Public Sub PropFormulas(ByVal wb As Excel.Workbook, ByVal a As String, ByVal ul As Long, ByVal ul2 As Long, ByVal pc As String, ByVal uc As String, ByVal uc2 As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        'a = aba
        'ul = ultLinha
        'ul2 = ultLinha2
        'pc = priColuna
        'uc = ultColuna
        'uc2 = ultColuna2
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        With wb.Worksheets(a)

            .select

            '------------------------------------------------------------------------------------------------------------------
            'acha a ultima linha da planilha atual para ajustar as fórmulas com o tamanho dos dados
            '------------------------------------------------------------------------------------------------------------------
            ul = .range(pc & .rows.count).end(Excel.XlDirection.xlUp).row

            If ul2 = 2 Then

                ul2 = .range(uc & ul2).end(Excel.XlDirection.xlDown).row

            Else

                ul2 = .range(uc & .rows.count).end(Excel.XlDirection.xlUp).row

            End If

            If ul > ul2 Then

                '------------------------------------------------------------------------------------------------------------------
                'Preenche com a formula todas as linhas
                '------------------------------------------------------------------------------------------------------------------
                .range(uc & ul2 & " : " & uc2 & ul).select
                wb.Application.Selection.filldown

            ElseIf ul < ul2 Then

                '------------------------------------------------------------------------------------------------------------------
                'Deleta as linhas que estão a mais
                '------------------------------------------------------------------------------------------------------------------
                .Range(uc & ul + 1 & " : " & uc2 & ul2).ClearContents

            End If

            '------------------------------------------------------------------------------------------------------------------
            'atualiza as fórmulas na planilha atual
            '------------------------------------------------------------------------------------------------------------------
            .Calculate

        End With

    End Sub

    Public Sub FecharWbks(ByVal f As Form, ByVal wb As Excel.Workbook, ByVal s As String, ByVal lId As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'wb = Planilha que está mexendo, ex.: Painel Disponibilidade (variável pDisp)
        's = salvar
        'lId = lbId 'recebe o ID do processo foi iniciado
        '------------------------------------------------------------------------------------------------------------------

        Threading.Thread.Sleep(2500)

        Dim lbId() As Control
        Dim id As Integer

        lbId = f.Controls.Find(lId, True)

        With wb

            If s = "N" Then

                .Application.DisplayAlerts = False

            ElseIf s = "S" Then

                .Application.DisplayAlerts = False
                .Application.CalculateBeforeSave = False
                .Save()

            End If

        End With

        Threading.Thread.Sleep(5000)

        '------------------------------------------------------------------------------------------------------------------
        'Fecha também os processos que rodam em segundo plano do excel que acabou de ser fechado
        '------------------------------------------------------------------------------------------------------------------
        id = CInt(lbId(0).Text)

        Process.GetProcessById(id).Kill()

    End Sub

    Public Sub FecharWbksAbertos()

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todos os Excels que estiverem abertos antes de começar a execução do programa.
        '------------------------------------------------------------------------------------------------------------------

        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")

        If xlp.Length > 0 Then

            titulo = "Atenção!"
            aviso = "Todas planilhas de Excel serão fechadas." & Chr(13) & "Certifique-se que tudo está salvo. Deseja continuar?"
            tipo = "Alerta"
            escolha = "SimNao"

            Mensagem(aviso, tipo, titulo, escolha)

            Select Case escolha
                Case "Sim"
                    For Each Process As Process In xlp

                        Process.Kill()

                    Next

                Case "Não"
                    titulo = "Salve tudo e retorne"
                    aviso = "O programa será fechado, para que você salve todas as planilhas."
                    tipo = "Alerta"
                    escolha = "Ok"

                    Mensagem(aviso, tipo, titulo, escolha)

                    End

            End Select

        End If

    End Sub

    Public Sub GravarHoras(ByRef dgv As DataGridView, ByVal p As String, ByVal i As DateTime, ByVal f As DateTime, ByVal t As TimeSpan)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'p = processo
        'i = inicio
        'f = fim
        't = total
        '------------------------------------------------------------------------------------------------------------------

        Dim posicao As Integer = Desenvolvedor.dgv_POpBkb.Rows.Count - 1
        Dim UltimaLinha As New DataGridViewRow
        Dim linha As String() = New String() {p, i, f, t.ToString()}

        'verifica se existe lina no data grid view
        If posicao >= 0 Then
            'Pega as linhas do DataGridView
            UltimaLinha = dgv.Rows.OfType(Of DataGridViewRow).Last()
        End If

        'adiciona a linha com as informações no grid
        dgv.Rows.Add(linha)

    End Sub

    Public Sub ExportarHoras(ByRef dgv As DataGridView, ByVal p As String)

        '------------------------------------------------------------------------------------------------------------------
        'Correspondência de variáveis:
        'dgv = dgvRelatorio
        'p = painel
        '------------------------------------------------------------------------------------------------------------------

        Dim NovoExcel As New Excel.Application()
        Dim arq As Excel.Workbook

        If dgv.Rows.Count > 0 Then
            Try
                arq = NovoExcel.Application.Workbooks.Add(Type.Missing)
                arq.ActiveSheet.name = p & " " & DateTime.Now.Day & "." & DateTime.Now.Month

                For i As Integer = 1 To dgv.Columns.Count
                    NovoExcel.Cells(1, i) = dgv.Columns(i - 1).HeaderText
                Next
                '
                For i As Integer = 0 To dgv.Rows.Count - 1
                    For j As Integer = 0 To dgv.Columns.Count - 1
                        NovoExcel.Cells(i + 2, j + 1) = dgv.Rows(i).Cells(j).Value.ToString()
                    Next
                Next
                '
                NovoExcel.Columns.AutoFit()
                '
                NovoExcel.Visible = True
            Catch ex As Exception
                MessageBox.Show("Erro : " + ex.Message)
                NovoExcel.Quit()
            End Try
        End If

    End Sub

    Public Sub Mensagem(ByVal a As String, ByVal tp As String, ByVal tl As String, ByRef esc As String)

        With form_Mensagem

            .lb_Titulo.Text = tl

            .lb_Aviso.Text = a

            If escolha = "SimNao" Then

                .btn_Sim.Visible = True
                .btn_Nao.Visible = True
                .btn_Ok.Visible = False

            ElseIf escolha = "Ok" Then

                .btn_Sim.Visible = False
                .btn_Nao.Visible = False
                .btn_Ok.Visible = True

            End If

            If tp = "Alerta" Then

                .pb_Estilo.Image = My.Resources.Alerta
                .p_Titulo.BackColor = Color.FromArgb(255, 135, 50) 'Laranja

            ElseIf tp = "Erro" Then

                .pb_Estilo.Image = My.Resources.Erro
                .p_Titulo.BackColor = Color.FromArgb(215, 35, 0) 'Vermelho

            End If

            .ShowDialog()

            esc = .lb_Escolha.Text

        End With

    End Sub

End Module
