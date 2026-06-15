Public Class Operario
    Inherits Empleado

    Private _turno As String
    Private _horasExtras As Integer

    Public Property Turno As String
        Get
            Return _turno
        End Get
        Set(value As String)

            If Not ModuloValidaciones.EsTurnoValido(value) Then
                Throw New Exception("El turno debe ser Matutino, Vespertino o Nocturno.")
            End If

            _turno = value
        End Set
    End Property

    Public Property HorasExtras As Integer
        Get
            Return _horasExtras
        End Get
        Set(value As Integer)

            If value < 0 OrElse value > 80 Then
                Throw New Exception("Las horas extra deben estar entre 0 y 80.")
            End If

            _horasExtras = value
        End Set
    End Property

    Public Sub New(nombre As String,
                   rfc As String,
                   salarioBase As Decimal,
                   departamento As String,
                   turno As String,
                   horasExtras As Integer)

        MyBase.New(nombre, rfc, salarioBase, departamento)

        Me.Turno = turno
        Me.HorasExtras = horasExtras

    End Sub

    Public Overrides Function CalcularPagoMensual() As Decimal

        Const pagoHoraExtra As Decimal = 150D

        Return SalarioBase + (HorasExtras * pagoHoraExtra)

    End Function

    Public Overrides Function ObtenerFicha() As String

        Return MyBase.ObtenerFicha() &
               vbCrLf & "Tipo: Operario" &
               vbCrLf & "Turno: " & Turno &
               vbCrLf & "Horas Extras: " & HorasExtras &
               vbCrLf & "Pago Mensual: " &
               ModuloValidaciones.FormatearMoneda(CalcularPagoMensual())

    End Function

End Class