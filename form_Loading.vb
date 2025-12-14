Imports System.ComponentModel

Public Class form_Loading

    Private Sub form_Loading_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        Dim id As Integer

        titulo = "Deseja fechar a Aplicação?"
        aviso = "Todos os processos serão cancelados." & Chr(13) & "E o progresso não foi salvo, será perdido." & Chr(13) & Chr(13) & "Deseja fechar mesmo assim?"
        tipo = "Alerta"
        escolha = "SimNao"

        Mensagem(aviso, tipo, titulo, escolha)

        '-----------------------------------------------------------------------------------------------------------------------------------
        'finaliza todos os Excels
        '-----------------------------------------------------------------------------------------------------------------------------------
        Select Case escolha
            Case "Sim"
                Dim xlp() As Process = Process.GetProcessesByName("EXCEL")

                If xlp.Length > 0 Then

                    For Each Process As Process In xlp

                        Process.Kill()

                    Next

                End If

                '-----------------------------------------------------------------------------------------------------------------------------------
                'finaliza o executavel no gerenciador de tarefas
                '-----------------------------------------------------------------------------------------------------------------------------------
                id = CInt(MainMenu.lbId_MainMenu.Text)

                Process.GetProcessById(id).Kill()

            Case "Não"
                e.Cancel = True
        End Select

    End Sub

End Class