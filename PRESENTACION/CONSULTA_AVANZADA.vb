Public Class CONSULTA_AVANZADA

    Private odt As New DataTable
    Dim indice As BindingManagerBase

    Private _Tipo As String

    Public Property Tipo As String

        Get
            Return _Tipo
        End Get
        Set(value As String)
            _Tipo = value
        End Set

    End Property

    Private Sub CONSULTA_AVANZADA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        odt = LOGICA.CLIENTE.Cliente_Select

        dgvClientes.DataSource = odt

        indice = BindingContext(odt)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click




        Try
            dgvClientes.DataSource = LOGICA.CLIENTE.SUPPLIERS_BUSQUEDAVAN(txtCompanyName.Text,
                                                             txtContactName.Text,
                                                             txtContactTitle.Text,
                                                             txtAddress.Text,
                                                             txtCity.Text,
                                                             txtRegion.Text,
                                                             txtPostalCode.Text,
                                                             txtCountry.Text,
                                                             txtPhone.Text,
                                                             txtFax.Text,
                                                             txtHomePage.Text,
                                                             txtSupplierID.Text)
        Catch ex As System.InvalidCastException
            dgvClientes.DataSource = LOGICA.CLIENTE.SUPPLIERS_BUSQUEDAVAN(txtCompanyName.Text,
                                                             txtContactName.Text,
                                                             txtContactTitle.Text,
                                                             txtAddress.Text,
                                                             txtCity.Text,
                                                             txtRegion.Text,
                                                             txtPostalCode.Text,
                                                             txtCountry.Text,
                                                             txtPhone.Text,
                                                             txtFax.Text,
                                                             txtHomePage.Text,
                                                             Nothing)
        End Try



    End Sub

    Private Sub txtSupplierID_TextChanged(sender As Object, e As EventArgs) Handles txtSupplierID.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim a As Control
        For Each a In Me.Controls
            If TypeOf a Is TextBox Then
                a.Text = Nothing
            End If
        Next
    End Sub
End Class