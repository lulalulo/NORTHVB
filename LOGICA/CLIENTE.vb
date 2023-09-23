Public Class CLIENTE

    Public Shared Function Cliente_Select() As DataTable

        Return DATOS.CLIENTE.Cliente_Select

    End Function

    Public Shared Function SUPPLIERS_INSERTAR(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String) As Boolean

        Return DATOS.CLIENTE.SUPPLIERS_INSERTAR(CompanyName, ContactName,
                                                ContactTitle, Address,
                                                City, Region,
                                                PostalCode, Country,
                                                Phone, Fax, HomePage)

    End Function

    'NUEVA VERSION
    Public Shared Function SUPPLIERS_INSERTAR2(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String, ByVal SupplierID As Integer) As Boolean

        Return DATOS.CLIENTE.SUPPLIERS_INSERTAR2(CompanyName, ContactName,
                                                ContactTitle, Address,
                                                City, Region,
                                                PostalCode, Country,
                                                Phone, Fax, HomePage, SupplierID)


    End Function

    Public Shared Function SUPPLIERS_UPDATE(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String, ByVal SupplierID As Integer) As Boolean

        Return DATOS.CLIENTE.SUPPLIERS_UPDATE(CompanyName, ContactName,
                                                ContactTitle, Address,
                                                City, Region,
                                                PostalCode, Country,
                                                Phone, Fax, HomePage, SupplierID)

    End Function


    'NUEVA VERSION
    Public Shared Function SUPPLIERS_UPDATE2(
        ByVal CompanyName As String, ByVal ContactName As String,
        ByVal ContactTitle As String, ByVal Address As String,
        ByVal City As String, ByVal Region As String,
        ByVal PostalCode As String, ByVal Country As String,
        ByVal Phone As String, ByVal Fax As String,
        ByVal HomePage As String, ByVal SupplierID As Integer) As Boolean

        Return DATOS.CLIENTE.SUPPLIERS_UPDATE2(CompanyName, ContactName,
                                                ContactTitle, Address,
                                                City, Region,
                                                PostalCode, Country,
                                                Phone, Fax, HomePage, SupplierID)

    End Function

    Public Shared Function SUPPLIERS_DELETE(ByVal SupplierID As Integer) As Boolean

        Return DATOS.CLIENTE.SUPPLIERS_DELETE(SupplierID)

    End Function

    Public Shared Function SUPPLIERS_BUSCAR(ByVal Campo As Integer, ByVal Criterio As String) As DataTable

        Return DATOS.CLIENTE.SUPPLIERS_BUSCAR(Campo, Criterio)

    End Function

    Public Shared Function SUPPLIERS_BUSQUEDAVAN(
       ByVal CompanyName As String, ByVal ContactName As String,
       ByVal ContactTitle As String, ByVal Address As String,
       ByVal City As String, ByVal Region As String,
       ByVal PostalCode As String, ByVal Country As String,
       ByVal Phone As String, ByVal Fax As String,
       ByVal HomePage As String, Optional ByVal SupplierID As Integer? = Nothing) As DataTable

        Return DATOS.CLIENTE.SUPPLIERS_BUSQUEDAVAN(CompanyName, ContactName,
                                                ContactTitle, Address,
                                                City, Region,
                                                PostalCode, Country,
                                                Phone, Fax, HomePage, SupplierID)

    End Function

End Class
