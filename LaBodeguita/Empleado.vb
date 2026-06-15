Public MustInherit Class Empleado

    Private _nombre As String
    Private _rfc As String
    Private _salarioBase As Decimal
    Private _departamento As String

    Private Shared _totalEmpleados As Integer = 0

    Public Property Nombre As String
        Get
            Return _nombre
        End Get
        Set(value As String)

            If Not ModuloValidaciones.EsNombreValido(value) Then
                Throw New Exception("El nombre es inválido.")
            End If

            _nombre = value
        End Set
    End Property

    Public Property RFC As String
        Get
            Return _rfc
        End Get
        Set(value As String)

            If Not ModuloValidaciones.EsRFCValido(value) Then
                Throw New Exception("El RFC es inválido.")
            End If

            _rfc = value
        End Set
    End Property

    Public Property SalarioBase As Decimal
        Get
            Return _salarioBase
        End Get
        Set(value As Decimal)

            If Not ModuloValidaciones.EsSalarioValido(value) Then
                Throw New Exception("El salario es inválido.")
            End If

            _salarioBase = value
        End Set
    End Property

    Public Property Departamento As String
        Get
            Return _departamento
        End Get
        Set(value As String)
            _departamento = value
        End Set
    End Property

    Public ReadOnly Property NombreCompleto As String
        Get
            Return Nombre & " (" & RFC & ")"
        End Get
    End Property

    Public Sub New(nombre As String,
                   rfc As String,
                   salarioBase As Decimal,
                   departamento As String)

        Me.Nombre = nombre
        Me.RFC = rfc
        Me.SalarioBase = salarioBase
        Me.Departamento = departamento

        _totalEmpleados += 1

    End Sub

    Public MustOverride Function CalcularPagoMensual() As Decimal

    Public Overridable Function ObtenerFicha() As String

        Return "Nombre: " & Nombre &
               vbCrLf & "RFC: " & RFC &
               vbCrLf & "Departamento: " & Departamento &
               vbCrLf & "Salario Base: " &
               ModuloValidaciones.FormatearMoneda(SalarioBase)

    End Function

    Public Shared Function GetTotal() As Integer
        Return _totalEmpleados
    End Function

End Class