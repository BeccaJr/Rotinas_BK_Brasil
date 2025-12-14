Public Class Desenvolvedor
    Private Sub VoltarParaOMenuPrincipalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VoltarParaOMenuPrincipalToolStripMenuItem.Click

        Me.Hide()
        MainMenu.Show()

    End Sub

    Private Sub Desenvolvedor_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        MainMenu.Show()

    End Sub

    Private Sub btn_Exportar_Click(sender As Object, e As EventArgs) Handles btn_Exportar.Click

#Region "Rotinas Manhã"

        Dim dgvRelatorio As DataGridView
        Dim painel As String

        dgvRelatorio = dgv_POpBkb
        painel = "Operacional BKB"

        ExportarHoras(dgvRelatorio, painel)

        dgvRelatorio = dgv_POpFz
        painel = "Operacional FZ"

        ExportarHoras(dgvRelatorio, painel)

        dgvRelatorio = dgv_POpPlk
        painel = "Operacional PLK"

        ExportarHoras(dgvRelatorio, painel)

        dgvRelatorio = dgv_PLbPlk
        painel = "Labor PLK"

        ExportarHoras(dgvRelatorio, painel)

        dgvRelatorio = dgv_PCmvPlk
        painel = "CMV PLK"

        ExportarHoras(dgvRelatorio, painel)

#End Region

    End Sub
End Class