Public Class CLIENTES_UTILITARIOS

    Private _Tipo As String

    Public Property Tipo As String

        Get
            Return _Tipo
        End Get
        Set(value As String)
            _Tipo = value
        End Set

    End Property

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim emptyTextBoxes =
        From txt In Me.Controls.OfType(Of TextBox)()
        Where txt.Text.Length = 0
        Select txt.Name
        If emptyTextBoxes.Any Then
            MessageBox.Show(String.Format("No se ingresaron datos en las tablas: {0}",
                        String.Join(",", emptyTextBoxes)))
            txtCompanyName.Focus()
            txtCompanyName.Select(0, txtCompanyName.TextLength)

        Else
            Dim rpta As Boolean = False

            If _Tipo = "Insertar" Then

                rpta = LOGICA.CLIENTE.SUPPLIERS_INSERTAR2(txtCompanyName.Text,
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

                _Tipo = ""

                If rpta = True Then
                    'MsgBox("Datos registrados con exito la orden " + txtSupplierID.Text, vbInformation + vbOK, "Aviso")
                    MsgBox("Datos registrados con exito.", vbInformation + vbOK, "Aviso")
                Else
                    MsgBox("No se encontro el procedimiento sp_Mttoshippers", vbInformation + vbObjectError, "Aviso")
                End If

            End If

            If _Tipo = "Modificar" Then
                rpta = LOGICA.CLIENTE.SUPPLIERS_UPDATE2(txtCompanyName.Text,
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

                _Tipo = ""

                If rpta = True Then
                    MsgBox("Datos registrados con exito de la orden " + txtSupplierID.Text, vbInformation + vbOK, "Aviso")
                Else
                    MsgBox("No se encontro el procedimiento sp_Mttoshippers", vbInformation + vbObjectError, "Aviso")
                End If

            End If

        End If






    End Sub

    Private Sub CLIENTES_UTILITARIOS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim a As Control
        For Each a In Me.Controls
            If TypeOf a Is TextBox Then
                a.Text = Nothing
            End If
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub txtSupplierID_TextChanged(sender As Object, e As EventArgs) Handles txtSupplierID.TextChanged

    End Sub
End Class