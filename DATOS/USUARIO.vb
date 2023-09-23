Imports System.Data.SqlClient

Public Class USUARIO

    Public Function USUARIO_LOGIN(ByVal Nombre As String, ByVal Contraseña As String) As Boolean

        Dim rpta As Boolean = False

        Dim oconx As New SqlConnection(My.Settings.cadena)
        oconx.Open()

        Dim ocmd As New SqlCommand
        ocmd.Connection = oconx

        ocmd.CommandType = CommandType.StoredProcedure
        ocmd.CommandText = "pa_USUARIO_LOGIN"

        ocmd.Parameters.AddWithValue("@Nombre", Nombre)
        ocmd.Parameters.AddWithValue("@Contraseña", Contraseña)

        Dim i As Integer = ocmd.ExecuteScalar
        If i >= 1 Then
            rpta = True
        Else
            rpta = False
        End If

        Return rpta


    End Function

End Class
