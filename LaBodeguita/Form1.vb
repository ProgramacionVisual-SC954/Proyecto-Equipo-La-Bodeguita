Public Class Form1

    Private empleados As New List(Of Empleado)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dgvEmpleados.Columns.Clear()

        dgvEmpleados.Columns.Add("Nombre", "Nombre")
        dgvEmpleados.Columns.Add("RFC", "RFC")
        dgvEmpleados.Columns.Add("Tipo", "Tipo")
        dgvEmpleados.Columns.Add("Pago", "Pago Mensual")

        lblTotalEmpleados.Text = "Total Empleados: 0"
        lblTotalNomina.Text = "Nómina Total: $0.00"

    End Sub

    Private Sub ActualizarTotales()

        lblTotalEmpleados.Text = "Total Empleados: " & empleados.Count

        Dim totalNomina As Decimal = 0

        For Each emp As Empleado In empleados
            totalNomina += emp.CalcularPagoMensual()
        Next

        lblTotalNomina.Text = "Nómina Total: " &
            ModuloValidaciones.FormatearMoneda(totalNomina)

    End Sub

    Private Sub btnAgregarGerente_Click(sender As Object, e As EventArgs) Handles btnAgregarGerente.Click

        Try

            Dim g As New Gerente(
                "Laura Carrillo",
                "ABCD123456EFG",
                20000D,
                "Administración",
                20,
                5)

            empleados.Add(g)

            dgvEmpleados.Rows.Add(
                g.Nombre,
                g.RFC,
                "Gerente",
                g.CalcularPagoMensual())

            ActualizarTotales()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub btnAgregarOperario_Click(sender As Object, e As EventArgs) Handles btnAgregarOperario.Click

        Try

            Dim o As New Operario(
                "Juan Pérez",
                "EFGH123456IJK",
                12000D,
                "Almacén",
                "Matutino",
                10)

            empleados.Add(o)

            dgvEmpleados.Rows.Add(
                o.Nombre,
                o.RFC,
                "Operario",
                o.CalcularPagoMensual())

            ActualizarTotales()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

        If dgvEmpleados.SelectedRows.Count = 0 Then
            MessageBox.Show("Seleccione un empleado.")
            Return
        End If

        Dim indice As Integer =
            dgvEmpleados.SelectedRows(0).Index

        empleados.RemoveAt(indice)

        dgvEmpleados.Rows.RemoveAt(indice)

        ActualizarTotales()

    End Sub

    Private Sub btnVerDetalle_Click(sender As Object, e As EventArgs) Handles btnVerDetalle.Click

        If dgvEmpleados.SelectedRows.Count = 0 Then
            MessageBox.Show("Seleccione un empleado.")
            Return
        End If

        Dim indice As Integer =
            dgvEmpleados.SelectedRows(0).Index

        MessageBox.Show(
            empleados(indice).ObtenerFicha(),
            "Detalle del Empleado")

    End Sub

    Private Sub dgvEmpleados_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvEmpleados.CellContentClick

    End Sub

    Private Sub lblTotalEmpleados_Click(sender As Object, e As EventArgs) Handles lblTotalEmpleados.Click

    End Sub

    Private Sub lblTotalNomina_Click(sender As Object, e As EventArgs) Handles lblTotalNomina.Click

    End Sub

End Class