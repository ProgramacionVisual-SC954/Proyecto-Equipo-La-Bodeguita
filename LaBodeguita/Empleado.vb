Public MustInherit Class Empleado
    Private _nombre As String
    Private _apellido As String
    Private _id As String
    Private _salarioBase As Decimal

    Public Sub New(nombre As String, apellido As String, id As String, salarioBase As Decimal)
        _nombre = nombre
        _apellido = apellido
        _id = id
        _salarioBase = salarioBase
    End Sub

    Public Property Nombre As String
        Get
            Return _nombre
        End Get
        Set(value As String)
            _nombre = value
        End Set
    End Property

    Public Property Apellido As String
        Get
            Return _apellido
        End Get
        Set(value As String)
            _apellido = value
        End Set
    End Property

    Public Property Id As String
        Get
            Return _id
        End Get
        Set(value As String)
            _id = value
        End Set
    End Property

    Public Property SalarioBase As Decimal
        Get
            Return _salarioBase
        End Get
        Set(value As Decimal)
            _salarioBase = value
        End Set
    End Property

    Public MustOverride Function CalcularPagoMensual() As Decimal
    Public MustOverride Function ObtenerFicha() As String
End Class