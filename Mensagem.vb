Public Class form_Mensagem
    Private Sub btn_Nao_Click(sender As Object, e As EventArgs) Handles btn_Nao.Click

        lb_Escolha.Text = "Não"
        Me.Close()

    End Sub

    Private Sub btn_Sim_Click(sender As Object, e As EventArgs) Handles btn_Sim.Click

        lb_Escolha.Text = "Sim"
        Me.Close()

    End Sub

    Private Sub btn_Ok_Click(sender As Object, e As EventArgs) Handles btn_Ok.Click

        Me.Close()

    End Sub

    Private Sub form_Mensagem_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Height = lb_Aviso.Height + 200

        btn_Ok.Left = 25
        btn_Ok.Top = Me.Height - 75

        btn_Nao.Left = 25
        btn_Nao.Top = Me.Height - 75

        btn_Sim.Left = 300
        btn_Sim.Top = Me.Height - 75

        pb_Estilo.Top = (lb_Aviso.Height / 2) + 50

    End Sub

End Class