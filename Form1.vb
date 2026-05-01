Imports System.Data.SqlClient

Public Class Form1
    Dim conn As New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""C:\Users\saman\Desktop\ENSC 26 GROUP 4 FINAL PROJ TRY\Characteristic Properties of Pure Species.mdf"";Integrated Security=True")
    Dim tr(), tc(), pc(), y(), aij(,), q(), a(), b(), qbar(), phi(), abar() As Double

    Private Sub btninput_Click(sender As Object, e As EventArgs) Handles btninput.Click

    End Sub

    Private Sub Initialize_Click(sender As Object, e As EventArgs) Handles btninitialize.Click
        Dim n As Integer
        Try
            n = Integer.Parse(txtn.Text)
            If n <= 0 Then
                MessageBox.Show("Number of components must be greater than zero.")
                Exit Sub
            ElseIf n > 5 Then
                MessageBox.Show("Number components must be less than five.")
                Exit Sub
            End If

            dgvinitialdata.Rows.Clear()
            For i As Integer = 1 To n
                dgvinitialdata.Rows.Add(i, "", "", "", "", "", "")
            Next

            If rdbliquid.Checked Then
                dgvinitialdata.Columns(2).HeaderText = "x₁"
            ElseIf rdbvapor.Checked Then
                dgvinitialdata.Columns(2).HeaderText = "y₁"
            End If
        Catch ex As Exception
            MessageBox.Show("Enter a valid number of components.")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btncalculate.Click
        dgvinitialdata.EndEdit()

        Dim n As Integer = Integer.Parse(txtn.Text)
        Dim dblpressure As Double = txtp.Text
        Dim dbltemp As Double = txtT.Text
        Dim amix As Double = 0, bmix As Double = 0
        Dim q, beta, yjaij, dblI, Z, Zold As Double

        txtn.Enabled = False
        txtp.Enabled = False
        txtT.Enabled = False

        ReDim y(n - 1), a(n - 1), b(n - 1), aij(n - 1, n - 1), tr(n - 1), tc(n - 1), pc(n - 1), phi(n - 1), abar(n - 1), qbar(n - 1)

        For i As Integer = 0 To n - 1
            y(i) = Val(dgvinitialdata.Rows(i).Cells(2).Value)
            tc(i) = Val(dgvinitialdata.Rows(i).Cells(3).Value)
            pc(i) = Val(dgvinitialdata.Rows(i).Cells(4).Value)

            tr(i) = dbltemp / tc(i)
            a(i) = (0.42748 * tr(i) ^ -0.5 * (83.14) ^ 2 * tc(i) ^ 2) / pc(i)
            b(i) = (0.08664 * 83.14 * tc(i)) / pc(i)

            dgvinitialdata.Rows(i).Cells(5).Value = tr(i)
            dgvinitialdata.Rows(i).Cells(6).Value = a(i)
            dgvinitialdata.Rows(i).Cells(7).Value = b(i)
        Next

        For i As Integer = 0 To n - 1
            bmix += y(i) * b(i)
        Next

        ' Ensure bmix isn't zero to avoid crash
        If bmix = 0 Then
            MessageBox.Show("Calculation error: Check your inputs.")
            Exit Sub
        End If

        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                aij(i, j) = Math.Sqrt(a(i) * a(j))
            Next
        Next

        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                amix += y(i) * y(j) * aij(i, j)
            Next
        Next

        q = amix / (bmix * 83.14 * dbltemp)
        beta = (bmix * dblpressure) / (83.14 * dbltemp)

        For i As Integer = 0 To n - 1
            yjaij = 0
            For j As Integer = 0 To n - 1
                yjaij += y(j) * aij(i, j)
            Next
            abar(i) = 2 * yjaij - amix
            qbar(i) = q * (1 + (abar(i) / amix) - (b(i) / bmix))
            dgvinitialdata.Rows(i).Cells(8).Value = abar(i)
            dgvinitialdata.Rows(i).Cells(9).Value = qbar(i)
        Next

        Dim iter As Integer = 0
        If rdbliquid.Checked Then
            Z = beta
            Do
                Zold = Z
                Z = beta + Zold * (Zold + beta) * ((1 + beta - Zold) / (q * beta))
                iter += 1
                If iter > 1000 Then Exit Do
            Loop Until Math.Abs(Z - Zold) < 0.000001
        ElseIf rdbvapor.Checked Then
            Z = 1
            Do
                Zold = Z
                Z = 1 + beta - (q * beta) * ((Zold - beta) / (Zold * (Zold + beta)))
                iter += 1
                If iter > 1000 Then Exit Do
            Loop Until Math.Abs(Z - Zold) < 0.000001
        End If

        If iter >= 1000 Then
            MessageBox.Show("Z failed to converge. The phase might be unstable at this Pressure or Temperature.")
            Exit Sub
        End If

        dblI = Math.Log((Z + beta) / Z)

        For i As Integer = 0 To n - 1
            phi(i) = Math.Exp(((b(i) / bmix) * (Z - 1) - Math.Log(Z - beta) - qbar(i) * dblI))
            dgvinitialdata.Rows(i).Cells(10).Value = phi(i)
        Next
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim cmbcomponents As DataGridViewComboBoxColumn = CType(dgvinitialdata.Columns(1), DataGridViewComboBoxColumn)
        Try
            If conn.State = ConnectionState.Closed Then conn.Open()
            Dim cmd As New SqlCommand("SELECT [Component] FROM [dbo].[Table]", conn)
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            cmbcomponents.Items.Clear()
            While reader.Read()
                cmbcomponents.Items.Add(reader("Component").ToString())
            End While
            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub cmbcomponents_selectedindexchanged(sender As Object, e As EventArgs)
        Dim combo As ComboBox = TryCast(sender, ComboBox)
        If combo Is Nothing OrElse combo.SelectedIndex = -1 Then Exit Sub

        Dim selecteditem As String = combo.Text
        Try
            If conn.State = ConnectionState.Closed Then conn.Open()
            Dim cmd As New SqlCommand("SELECT [Tc (K)], [Pc (bar)] FROM [dbo].[Table] WHERE [Component]=@n1", conn)
            cmd.Parameters.AddWithValue("@n1", selecteditem)
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            If reader.Read() Then
                Dim rowindex As Integer = dgvinitialdata.CurrentCell.RowIndex
                dgvinitialdata.Rows(rowindex).Cells(3).Value = reader("Tc (K)")
                dgvinitialdata.Rows(rowindex).Cells(4).Value = reader("Pc (bar)")
            End If
            reader.Close()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub dgvinitialdata_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgvinitialdata.EditingControlShowing
        If dgvinitialdata.CurrentCell.ColumnIndex = 1 AndAlso TypeOf e.Control Is ComboBox Then
            Dim combo As ComboBox = CType(e.Control, ComboBox)
            combo.DropDownStyle = ComboBoxStyle.DropDown
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            combo.AutoCompleteSource = AutoCompleteSource.ListItems
            RemoveHandler combo.SelectedIndexChanged, AddressOf cmbcomponents_selectedindexchanged
            AddHandler combo.SelectedIndexChanged, AddressOf cmbcomponents_selectedindexchanged
        End If
    End Sub
End Class