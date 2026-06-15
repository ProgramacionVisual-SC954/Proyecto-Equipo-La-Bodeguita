Public Class Gerente
    Inherits Empleado

    Private _bono As Decimal
    Private _personasACargo As Integer

    Public Sub New(nombre As String, apellido As String, id As String, salarioBase As Decimal, bono As Decimal, personasACargo As Integer)
        MyBase.New(nombre, apellido, id, salarioBase)
        _bono = bono
        _personasACargo = personasACargo
    End Sub

    Public Property Bono As Decimal
        Get
            Return _bono
        End Get
        Set(value As Decimal)
            If value >= 0 Then
                _bono = value
            Else
                Throw New ArgumentException("El bono no puede ser negativo.")
            End If
        End Set
    End Property

    Public Property PersonasACargo As Integer
        Get
            Return _personasACargo
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _personasACargo = value
            Else
                Throw New ArgumentException("Las personas a cargo no pueden ser negativas.")
            End If
        End Set
    End Property

    Public Overrides Function CalcularPagoMensual() As Decimal
        Return SalarioBase + _bono
    End Function

    Public Overrides Function ObtenerFicha() As String
        Return $"[GERENTE] {Nombre} {Apellido} | ID: {Id} | Salario Base: {SalarioBase:C} | Bono: {_bono:C} | Personas a cargo: {_personasACargo} | Pago Mensual: {CalcularPagoMensual():C}"
    End Function
End Class