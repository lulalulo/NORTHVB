Public Class CLIENTE

    Private odt As New DataTable
    Dim indice As BindingManagerBase

    Private Sub CLIENTE_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        odt = LOGICA.CLIENTE.Cliente_Select

        dgvClientes.DataSource = odt

        indice = BindingContext(odt)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim f As New CLIENTES_UTILITARIOS

        f.Tipo = "Insertar"

        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then

            odt = LOGICA.CLIENTE.Cliente_Select

            dgvClientes.DataSource = odt


        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim f As New CLIENTES_UTILITARIOS

        f.Tipo = "Modificar"

        f.txtSupplierID.Text = dgvClientes.CurrentRow.Cells(0).Value
        f.txtCompanyName.Text = dgvClientes.CurrentRow.Cells(1).Value
        'f.txtContactName.Text = dgvClientes.CurrentRow.Cells(2).Value
        'f.txtContactTitle.Text = dgvClientes.CurrentRow.Cells(3).Value
        'f.txtAddress.Text = dgvClientes.CurrentRow.Cells(4).Value
        'f.txtCity.Text = dgvClientes.CurrentRow.Cells(5).Value
        'f.txtRegion.Text = dgvClientes.CurrentRow.Cells(6).Value
        'f.txtPostalCode.Text = dgvClientes.CurrentRow.Cells(7).Value
        'f.txtCountry.Text = dgvClientes.CurrentRow.Cells(8).Value
        'f.txtPhone.Text = dgvClientes.CurrentRow.Cells(9).Value
        'f.txtFax.Text = dgvClientes.CurrentRow.Cells(10).Value
        'f.txtHomePage.Text = dgvClientes.CurrentRow.Cells(11).Value


        If IsDBNull(dgvClientes.CurrentRow.Cells(2).Value) Then
            f.txtContactName.Text = dgvClientes.CurrentRow.Cells(2).Value.ToString()
        Else
            f.txtContactName.Text = dgvClientes.CurrentRow.Cells(2).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(3).Value) Then
            f.txtContactTitle.Text = dgvClientes.CurrentRow.Cells(3).Value.ToString()
        Else
            f.txtContactTitle.Text = dgvClientes.CurrentRow.Cells(3).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(4).Value) Then
            f.txtAddress.Text = dgvClientes.CurrentRow.Cells(4).Value.ToString()
        Else
            f.txtAddress.Text = dgvClientes.CurrentRow.Cells(4).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(5).Value) Then
            f.txtCity.Text = dgvClientes.CurrentRow.Cells(5).Value.ToString()
        Else
            f.txtCity.Text = dgvClientes.CurrentRow.Cells(5).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(6).Value) Then
            f.txtRegion.Text = dgvClientes.CurrentRow.Cells(6).Value.ToString
        Else
            f.txtRegion.Text = dgvClientes.CurrentRow.Cells(6).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(7).Value) Then
            f.txtPostalCode.Text = dgvClientes.CurrentRow.Cells(7).Value.ToString()
        Else
            f.txtPostalCode.Text = dgvClientes.CurrentRow.Cells(7).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(8).Value) Then
            f.txtCountry.Text = dgvClientes.CurrentRow.Cells(8).Value.ToString()
        Else
            f.txtCountry.Text = dgvClientes.CurrentRow.Cells(8).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(9).Value) Then
            f.txtPhone.Text = dgvClientes.CurrentRow.Cells(9).Value.ToString()
        Else
            f.txtPhone.Text = dgvClientes.CurrentRow.Cells(9).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(10).Value) Then
            f.txtFax.Text = dgvClientes.CurrentRow.Cells(10).Value.ToString()
        Else
            f.txtFax.Text = dgvClientes.CurrentRow.Cells(10).Value
        End If

        If IsDBNull(dgvClientes.CurrentRow.Cells(11).Value) Then
            f.txtHomePage.Text = dgvClientes.CurrentRow.Cells(11).Value.ToString()
        Else
            f.txtHomePage.Text = dgvClientes.CurrentRow.Cells(11).Value
        End If






        Dim id As Integer = dgvClientes.CurrentRow.Cells(0).Value

        If f.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

            odt = LOGICA.CLIENTE.Cliente_Select

            dgvClientes.DataSource = odt

            indice = BindingContext(odt)

            For i = 0 To dgvClientes.Rows.Count - 1

                If id = dgvClientes.Rows(i).Cells(0).Value Then

                    dgvClientes.CurrentCell = dgvClientes.Rows(i).Cells(1)

                    Exit For

                End If

            Next
        Else
            f.Tipo = ""

        End If

        f.Dispose()



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim pos As Integer
        pos = indice.Position

        If MsgBox("¿Desea eliminar dicho registro?", vbQuestion + vbOKCancel, "AVISO") = vbOK Then

            Dim Id As Integer = dgvClientes.CurrentRow.Cells(0).Value

            If LOGICA.CLIENTE.SUPPLIERS_DELETE(Id) = True Then

                MsgBox("Registro eliminado", vbInformation, "AVISO")

                odt = LOGICA.CLIENTE.Cliente_Select

                dgvClientes.DataSource = odt

                indice = BindingContext(odt)

                indice.Position = pos - 1

            ElseIf Id = 1 Then
                MsgBox("The DELETE statement conflicted with the REFERENCE constraint FK_Products_Suppliers. The conflict occurred in database Northwind, table dbo.Products, column 'SupplierID",
                       vbInformation + vbObjectError, "Aviso")
            Else

                MsgBox("Hubo un error al ingresar los datos", vbInformation + vbObjectError, "Aviso")
            End If

        End If


    End Sub

    Private Sub txtCriterio_TextChanged(sender As Object, e As EventArgs) Handles txtCriterio.TextChanged

        dgvClientes.DataSource = LOGICA.CLIENTE.SUPPLIERS_BUSCAR(cboCampo.SelectedIndex, txtCriterio.Text)

    End Sub

    Private Sub cboCampo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCampo.SelectedIndexChanged



    End Sub


    Private Sub Label1_Click_1(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub dgvClientes_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvClientes.CellContentClick

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim f As New CONSULTA_AVANZADA

        f.Tipo = "Consulta Avanzada"

        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then

            odt = LOGICA.CLIENTE.Cliente_Select

            dgvClientes.DataSource = odt


        End If
    End Sub
End Class