Public Class USUARIO
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim ousuario As New LOGICA.USUARIO
        Dim rpta As Boolean

        rpta = ousuario.USUARIO_LOGIN(txtusuario.Text, txtcontraseña.Text)

        Static i As Int16 = 1

        If i <= 2 Then
            If rpta = True Then
                MsgBox("Inicio de sesion exitoso", vbExclamation)
                CLIENTE.Show()
                Me.Hide()
            Else
                MsgBox("Contraseña/Usuario incorrecto", vbCritical, "Aviso")
                txtusuario.Focus()
                txtusuario.Select(0, txtusuario.TextLength)
                i = i + 1

            End If
        Else
            MsgBox("Demasiados intentos procesados" +
                   "El sistema se cerrara", vbCritical, "Aviso")
            Me.Close()


        End If




    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        txtcontraseña.UseSystemPasswordChar = Not CheckBox1.Checked
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

    End Sub
End Class