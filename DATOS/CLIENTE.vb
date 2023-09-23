Imports System.Data.SqlClient
Imports System.Windows

Public Class CLIENTE

    Public Shared Function Cliente_Select() As DataTable
        Dim oconex As New SqlConnection(My.Settings.cadena)
        oconex.Open()

        Dim ocomando As New SqlCommand
        ocomando.Connection = oconex
        ocomando.CommandType = CommandType.StoredProcedure
        ocomando.CommandText = "pa_SELECT_SUPPLIERS"

        Dim oidr As IDataReader
        oidr = ocomando.ExecuteReader

        Dim odt As New DataTable
        odt.Load(oidr)

        Return odt

        oconex.Close()


    End Function

    Public Shared Function SUPPLIERS_INSERTAR(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String) As Boolean

        Dim oconx As New SqlConnection(My.Settings.cadena)
        oconx.Open()

        Dim rpta As Boolean = False
        Dim filasAfectadas As Int16


        Dim ocomando As New SqlCommand
        ocomando.Connection = oconx
        ocomando.CommandType = CommandType.StoredProcedure
        ocomando.CommandText = "pa_SUPPLIERS_INSERTAR"

        ocomando.Parameters.AddWithValue("@CompanyName", CompanyName)
        ocomando.Parameters.AddWithValue("@ContactName", ContactName)
        ocomando.Parameters.AddWithValue("@ContactTitle", ContactTitle)
        ocomando.Parameters.AddWithValue("@Address", Address)
        ocomando.Parameters.AddWithValue("@City", City)
        ocomando.Parameters.AddWithValue("@Region", Region)
        ocomando.Parameters.AddWithValue("@PostalCode", PostalCode)
        ocomando.Parameters.AddWithValue("@Country", Country)
        ocomando.Parameters.AddWithValue("@Phone", Phone)
        ocomando.Parameters.AddWithValue("@Fax", Fax)
        ocomando.Parameters.AddWithValue("@HomePage", HomePage)

        filasAfectadas = ocomando.ExecuteNonQuery
        oconx.Close()

        If filasAfectadas > 0 Then
            rpta = True

        End If

        Return rpta


    End Function

    'sp_mttoSuppliers
    Public Shared Function SUPPLIERS_INSERTAR2(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String, ByVal SupplierID As Integer) As Boolean

        Dim oconx As New SqlConnection(My.Settings.cadena)
        oconx.Open()

        Dim rpta As Boolean = False
        Dim filasAfectadas As Int16


        Dim ocomando As New SqlCommand
        ocomando.Connection = oconx
        ocomando.CommandType = CommandType.StoredProcedure
        'ocomando.CommandText = "pa_SUPPLIERS_INSERTAR"
        ocomando.CommandText = "sp_Mttoshippers"

        ocomando.Parameters.AddWithValue("@SupplierID", Integer.MaxValue)
        ocomando.Parameters.AddWithValue("@CompanyName", CompanyName)
        ocomando.Parameters.AddWithValue("@ContactName", ContactName)
        ocomando.Parameters.AddWithValue("@ContactTitle", ContactTitle)
        ocomando.Parameters.AddWithValue("@Address", Address)
        ocomando.Parameters.AddWithValue("@City", City)
        ocomando.Parameters.AddWithValue("@Region", Region)
        ocomando.Parameters.AddWithValue("@PostalCode", PostalCode)
        ocomando.Parameters.AddWithValue("@Country", Country)
        ocomando.Parameters.AddWithValue("@Phone", Phone)
        ocomando.Parameters.AddWithValue("@Fax", Fax)
        ocomando.Parameters.AddWithValue("@HomePage", HomePage)

        'filasAfectadas = ocomando.ExecuteNonQuery
        Try
            filasAfectadas = ocomando.ExecuteNonQuery
        Catch ex As System.Data.SqlClient.SqlException

            rpta = False
        End Try

        oconx.Close()

        If filasAfectadas > 0 Then
            rpta = True

        End If

        Return rpta


    End Function


    Public Shared Function SUPPLIERS_UPDATE(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String, ByVal SupplierID As Integer) As Boolean

        Dim oconex As New SqlConnection(My.Settings.cadena)
        oconex.Open()

        Dim rpta As Boolean = False
        Dim filasAfectadas As Int16

        Dim ocmd As New SqlCommand
        ocmd.Connection = oconex
        ocmd.CommandType = CommandType.StoredProcedure
        ocmd.CommandText = "pa_SUPPLIERS_UPDATE"

        ocmd.Parameters.AddWithValue("@Id", SupplierID)
        ocmd.Parameters.AddWithValue("@CompanyName", CompanyName)
        ocmd.Parameters.AddWithValue("@ContactName", ContactName)
        ocmd.Parameters.AddWithValue("@ContactTitle", ContactTitle)
        ocmd.Parameters.AddWithValue("@Address", Address)
        ocmd.Parameters.AddWithValue("@City", City)
        ocmd.Parameters.AddWithValue("@Region", Region)
        ocmd.Parameters.AddWithValue("@PostalCode", PostalCode)
        ocmd.Parameters.AddWithValue("@Country", Country)
        ocmd.Parameters.AddWithValue("@Phone", Phone)
        ocmd.Parameters.AddWithValue("@Fax", Fax)
        ocmd.Parameters.AddWithValue("@HomePage", HomePage)

        filasAfectadas = ocmd.ExecuteNonQuery
        oconex.Close()

        If filasAfectadas > 0 Then
            rpta = True
        End If

        Return rpta

    End Function

    ''nueva version de update

    Public Shared Function SUPPLIERS_UPDATE2(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String, ByVal SupplierID As Integer) As Boolean

        Dim oconex As New SqlConnection(My.Settings.cadena)
        oconex.Open()

        Dim rpta As Boolean = False
        Dim filasAfectadas As Int16

        Dim ocmd As New SqlCommand
        ocmd.Connection = oconex
        ocmd.CommandType = CommandType.StoredProcedure
        ocmd.CommandText = "sp_Mttoshippers"

        ocmd.Parameters.AddWithValue("@SupplierID", SupplierID)
        ocmd.Parameters.AddWithValue("@CompanyName", CompanyName)
        ocmd.Parameters.AddWithValue("@ContactName", ContactName)
        ocmd.Parameters.AddWithValue("@ContactTitle", ContactTitle)
        ocmd.Parameters.AddWithValue("@Address", Address)
        ocmd.Parameters.AddWithValue("@City", City)
        ocmd.Parameters.AddWithValue("@Region", Region)
        ocmd.Parameters.AddWithValue("@PostalCode", PostalCode)
        ocmd.Parameters.AddWithValue("@Country", Country)
        ocmd.Parameters.AddWithValue("@Phone", Phone)
        ocmd.Parameters.AddWithValue("@Fax", Fax)
        ocmd.Parameters.AddWithValue("@HomePage", HomePage)

        'filasAfectadas = ocmd.ExecuteNonQuery
        Try
            filasAfectadas = ocmd.ExecuteNonQuery
        Catch ex As System.Data.SqlClient.SqlException

            rpta = False
        End Try

        oconex.Close()

        If filasAfectadas > 0 Then
            rpta = True
        End If

        Return rpta

    End Function

    Public Shared Function SUPPLIERS_DELETE(ByVal SupplierID As Integer) As Boolean

        Dim oconx As New SqlConnection(My.Settings.cadena)
        oconx.Open()

        Dim rpta As Boolean = False
        Dim filasafectadas As Int16

        Dim ocmd As New SqlCommand
        ocmd.Connection = oconx
        ocmd.CommandType = CommandType.StoredProcedure
        ocmd.CommandText = "pa_SUPPLIERS_DELETE"

        ocmd.Parameters.AddWithValue("@id", SupplierID)
        Try
            filasafectadas = ocmd.ExecuteNonQuery
        Catch ex As System.Data.SqlClient.SqlException
            rpta = False
        End Try

        oconx.Close()

        If filasafectadas > 0 Then
            rpta = True
        End If

        Return rpta


    End Function

    Public Shared Function SUPPLIERS_BUSCAR(ByVal Campo As Integer, ByVal Criterio As String) As DataTable



        Dim oconx As New SqlConnection(My.Settings.cadena)
        oconx.Open()

        Dim ocmd As New SqlCommand
        ocmd.Connection = oconx
        ocmd.CommandType = CommandType.StoredProcedure
        ocmd.CommandText = "pa_SUPPLIERS_BUSCAR"

        ocmd.Parameters.AddWithValue("@Campo", Campo)
        ocmd.Parameters.AddWithValue("@Criterio", Criterio)

        Dim odr As SqlDataReader
        odr = ocmd.ExecuteReader

        Dim odt As New DataTable
        odt.Load(odr)

        Return odt

    End Function

    Public Shared Function SUPPLIERS_BUSQUEDAVAN(
       ByVal CompanyName As String, ByVal ContactName As String,
       ByVal ContactTitle As String, ByVal Address As String,
       ByVal City As String, ByVal Region As String,
       ByVal PostalCode As String, ByVal Country As String,
       ByVal Phone As String, ByVal Fax As String,
       ByVal HomePage As String, Optional ByVal SupplierID As Integer? = Nothing) As DataTable

        Dim oconx As New SqlConnection(My.Settings.cadena)
        oconx.Open()






        Dim ocomando As New SqlCommand
        ocomando.Connection = oconx
        ocomando.CommandType = CommandType.StoredProcedure
        'ocomando.CommandText = "pa_SUPPLIERS_INSERTAR"
        ocomando.CommandText = "pa_SUPPLIERS_BUSCARAVAN"

        If SupplierID.HasValue Then
            ocomando.Parameters.AddWithValue("@SupplierID", SupplierID)
        Else
            ocomando.Parameters.AddWithValue("@SupplierID", DBNull.Value)
        End If

        If String.IsNullOrEmpty(CompanyName) Or String.IsNullOrWhiteSpace(CompanyName) Then
            ocomando.Parameters.AddWithValue("@CompanyName", DBNull.Value)

        Else
            ocomando.Parameters.AddWithValue("@CompanyName", CompanyName)
        End If

        If String.IsNullOrEmpty(ContactName) Or String.IsNullOrWhiteSpace(ContactName) Then

            ocomando.Parameters.AddWithValue("@ContactName", DBNull.Value)
        Else
            ocomando.Parameters.AddWithValue("@ContactName", ContactName)
        End If

        If String.IsNullOrEmpty(ContactTitle) Or String.IsNullOrWhiteSpace(ContactTitle) Then

            ocomando.Parameters.AddWithValue("@ContactTitle", DBNull.Value)
        Else
            ocomando.Parameters.AddWithValue("@ContactTitle", ContactTitle)
        End If

        If String.IsNullOrEmpty(Address) Or String.IsNullOrWhiteSpace(Address) Then
            ocomando.Parameters.AddWithValue("@Address", DBNull.Value)

        Else
            ocomando.Parameters.AddWithValue("@Address", Address)
        End If

        If String.IsNullOrEmpty(City) Or String.IsNullOrWhiteSpace(City) Then
            ocomando.Parameters.AddWithValue("@City", DBNull.Value)
        Else

            ocomando.Parameters.AddWithValue("@City", City)
        End If

        If String.IsNullOrEmpty(Region) Or String.IsNullOrWhiteSpace(Region) Then
            ocomando.Parameters.AddWithValue("@Region", DBNull.Value)
        Else
            ocomando.Parameters.AddWithValue("@Region", Region)

        End If

        If String.IsNullOrEmpty(PostalCode) Or String.IsNullOrWhiteSpace(PostalCode) Then
            ocomando.Parameters.AddWithValue("@PostalCode", DBNull.Value)

        Else
            ocomando.Parameters.AddWithValue("@PostalCode", PostalCode)
        End If

        If String.IsNullOrEmpty(Country) Or String.IsNullOrWhiteSpace(Country) Then
            ocomando.Parameters.AddWithValue("@Country", DBNull.Value)

        Else
            ocomando.Parameters.AddWithValue("@Country", Country)
        End If

        If String.IsNullOrEmpty(Phone) Or String.IsNullOrWhiteSpace(Phone) Then
            ocomando.Parameters.AddWithValue("@Phone", DBNull.Value)

        Else
            ocomando.Parameters.AddWithValue("@Phone", Phone)
        End If

        If String.IsNullOrEmpty(Fax) Or String.IsNullOrWhiteSpace(Fax) Then
            ocomando.Parameters.AddWithValue("@Fax", DBNull.Value)

        Else
            ocomando.Parameters.AddWithValue("@Fax", Fax)
        End If

        If String.IsNullOrEmpty(HomePage) Or String.IsNullOrWhiteSpace(HomePage) Then

            ocomando.Parameters.AddWithValue("@HomePage", DBNull.Value)
        Else
            ocomando.Parameters.AddWithValue("@HomePage", HomePage)
        End If

        Dim odr As SqlDataReader
        odr = ocomando.ExecuteReader


        Dim odt As New DataTable
        odt.Load(odr)

        Return odt


    End Function

End Class
