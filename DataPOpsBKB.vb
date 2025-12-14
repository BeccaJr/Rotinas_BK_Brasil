Public Class form_DataPOpsBKB
    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged

        '------------------------------------------------------------------------------------------------------------------
        'Coloca a data do Painel na lbl_DataPainel
        '------------------------------------------------------------------------------------------------------------------
        form_POpBkb.lbl_DataPainel.Text = MonthCalendar1.SelectionRange.Start.Date.ToShortDateString()

        form_POpBkb.lbNome_DataPainel.Text = MonthCalendar1.SelectionRange.Start.Date.ToShortDateString()

        form_POpBkb.lbNome_DataPainel.ForeColor = Color.Black

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Hide()

    End Sub
End Class