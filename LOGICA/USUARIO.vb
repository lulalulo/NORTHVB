Public Class USUARIO

    Public Function USUARIO_LOGIN(ByVal Nombre As String, ByVal Contraseña As String) As Boolean

        Dim ousuario As New DATOS.USUARIO
        Return ousuario.USUARIO_LOGIN(Nombre, Contraseña)

    End Function

End Class
