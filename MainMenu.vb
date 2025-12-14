Imports System.Drawing.Color

Public Class MainMenu

    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles Me.Load

        Pb_LogoBK1.Parent = Panel1
        Pb_LogoBK2.Parent = Panel2
        Pb_LogoPLK1.Parent = Panel1
        Pb_LogoPLK2.Parent = Panel2

        lbId_MainMenu.Text = Process.GetCurrentProcess.Id

        Label1.Select()

    End Sub

    Private Sub btn_Fechar_Click_1(sender As Object, e As EventArgs) Handles btn_Fechar.Click

        Application.Exit()

    End Sub

    Private Sub btn_Minimizar_Click_1(sender As Object, e As EventArgs) Handles btn_Minimizar.Click

        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub btn_Desenvolv_Click_1(sender As Object, e As EventArgs) Handles btn_Desenvolv.Click

        Desenvolvedor.Show()

    End Sub

    Private Sub btn_HistModif_Click_1(sender As Object, e As EventArgs) Handles btn_HistModif.Click

        form_LogModificacoes.Show()

    End Sub

    Private Sub btn_OpBkb_Click(sender As Object, e As EventArgs) Handles btn_OpBkb.Click

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todas as planilhas abertas
        '------------------------------------------------------------------------------------------------------------------
        FecharWbksAbertos()

        Me.Hide()

        form_POpBkb.Show()

    End Sub

    Private Sub btn_OpFz_Click(sender As Object, e As EventArgs) Handles btn_OpFz.Click

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todas as planilhas abertas
        '------------------------------------------------------------------------------------------------------------------
        FecharWbksAbertos()

        Me.Hide()
        form_POpFz.Show()

    End Sub

    Private Sub btn_OpPlk_Click(sender As Object, e As EventArgs) Handles btn_OpPlk.Click

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todas as planilhas abertas
        '------------------------------------------------------------------------------------------------------------------
        FecharWbksAbertos()

        Me.Hide()

        form_POpPlk.Show()

    End Sub

    Private Sub btn_PLabPlk_Click(sender As Object, e As EventArgs) Handles btn_PLabPlk.Click

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todas as planilhas abertas
        '------------------------------------------------------------------------------------------------------------------
        FecharWbksAbertos()

        Me.Hide()

        form_LaborPlk.Show()

    End Sub

    Private Sub btn_PCmvPlk_Click(sender As Object, e As EventArgs) Handles btn_PCmvPlk.Click

        '------------------------------------------------------------------------------------------------------------------
        'Fecha todas as planilhas abertas
        '------------------------------------------------------------------------------------------------------------------
        FecharWbksAbertos()

        Me.Hide()

        form_CmvPlk.Show()

    End Sub

End Class
