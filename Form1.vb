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

    Dim dblpressure, dbltemp As Double
    Dim n As Integer

    If Not Integer.TryParse(txtn.Text, n) OrElse n <= 0 Then
        MessageBox.Show("Number of components must be greater than zero.")
        txtn.Focus()
        txtp.BackColor = Color.MistyRose
        Exit Sub
    ElseIf n > 5 Then
        MessageBox.Show("Number components must be less than five.")
        txtn.Focus()
        txtp.BackColor = Color.MistyRose
        Exit Sub
    End If

    If Not Double.TryParse(txtp.Text, dblpressure) OrElse dblpressure <= 0 Then
        MessageBox.Show("Please enter a valid positive numeric value for Pressure.")
        txtp.Focus()
        txtp.BackColor = Color.MistyRose
        Exit Sub
    End If

    If Not Double.TryParse(txtT.Text, dbltemp) OrElse dbltemp <= 0 Then
        MessageBox.Show("Please enter a valid positive numeric value for Temperature.")
        txtT.Focus()
        txtp.BackColor = Color.MistyRose
        Exit Sub
    End If

    If Not (rdbBar.Checked Or rdbAtm.Checked Or rdbPsi.Checked Or
        rdbMmhg.Checked Or rdbPa.Checked Or rdbKpa.Checked) Then
        MessageBox.Show("Please select a Pressure unit (e.g., bar).")
        Exit Sub
    End If

    If Not (rdbKelvin.Checked Or rdbCelsius.Checked Or
        rdbFahrenheit.Checked Or rdbRankine.Checked) Then
        MessageBox.Show("Please select a Temperature unit (e.g., Kelvin).")
        Exit Sub
    End If

    If Not (rdbvapor.Checked Or rdbliquid.Checked) Then
        MessageBox.Show("Please select the Phase of Mixture (Vapor or Liquid).")
        Exit Sub
    End If

    If rdbAtm.Checked Then
        dblpressure *= 1.01325
    ElseIf rdbPsi.Checked Then
        dblpressure *= 0.0689476
    ElseIf rdbMmhg.Checked Then
        dblpressure /= 750.062
    ElseIf rdbPa.Checked Then
        dblpressure /= 100000.0
    ElseIf rdbKpa.Checked Then
        dblpressure /= 100.0
    End If

    If rdbCelsius.Checked Then
        dbltemp += 273.15
    ElseIf rdbFahrenheit.Checked Then
        dbltemp = (dbltemp - 32) * 5 / 9 + 273.15
    ElseIf rdbRankine.Checked Then
        dbltemp = dbltemp * 5 / 9
    End If


    Dim sumY As Double = 0
    For i As Integer = 0 To n - 1
        Dim cellComp = dgvinitialdata.Rows(i).Cells(1)
        Dim cellVal = dgvinitialdata.Rows(i).Cells(2)


        If cellComp.Value Is Nothing OrElse cellComp.Value.ToString() = "" Then
            MessageBox.Show($"Please select a component for row {i + 1}.")
            cellComp.Style.BackColor = Color.LightYellow
            Exit Sub
        Else
            cellComp.Style.BackColor = Color.White
        End If


        Dim valY As Double
        If Not Double.TryParse(cellVal.Value?.ToString(), valY) OrElse valY < 0 OrElse valY > 1 Then
            MessageBox.Show($"Invalid mole fraction at row {i + 1}. Must be between 0 and 1.")
            cellVal.Style.BackColor = Color.MistyRose
            Exit Sub
        Else
            cellVal.Style.BackColor = Color.White
            sumY += valY
        End If
    Next

    If Math.Abs(sumY - 1.0) > 0.0001 Then
        MessageBox.Show($"Mole fractions must sum to 1.0. Current sum: {sumY}")
        Exit Sub
    End If

    Dim amix As Double = 0, bmix As Double = 0
    Dim q_val, beta, yjaij, dblI, Z, Zold As Double
    Dim iter As Integer = 0



    txtn.Enabled = False
    txtp.Enabled = False
    txtT.Enabled = False
    btncalculate.Enabled = False

    ReDim y(n - 1), a(n - 1), b(n - 1), aij(n - 1, n - 1), tr(n - 1), tc(n - 1), pc(n - 1), phi(n - 1), abar(n - 1), qbar(n - 1)
    Const R_CONST As Double = 83.14

    Try
        For i As Integer = 0 To n - 1
            y(i) = Convert.ToDouble(dgvinitialdata.Rows(i).Cells(2).Value)
            tc(i) = Convert.ToDouble(dgvinitialdata.Rows(i).Cells(3).Value)
            pc(i) = Convert.ToDouble(dgvinitialdata.Rows(i).Cells(4).Value)

            tr(i) = dbltemp / tc(i)
            a(i) = (0.42748 * tr(i) ^ -0.5 * (R_CONST) ^ 2 * tc(i) ^ 2) / pc(i)
            b(i) = (0.08664 * R_CONST * tc(i)) / pc(i)

            dgvinitialdata.Rows(i).Cells(5).Value = Math.Round(tr(i), 4)
            dgvinitialdata.Rows(i).Cells(6).Value = Math.Round(a(i), 2)
            dgvinitialdata.Rows(i).Cells(7).Value = Math.Round(b(i), 4)
        Next

        For i As Integer = 0 To n - 1
            bmix += y(i) * b(i)
        Next

        If bmix = 0 Then Throw New Exception("Bmix cannot be zero.")


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

        q_val = amix / (bmix * R_CONST * dbltemp)
        beta = (bmix * dblpressure) / (R_CONST * dbltemp)

        For i As Integer = 0 To n - 1
            yjaij = 0
            For j As Integer = 0 To n - 1
                yjaij += y(j) * aij(i, j)
            Next
            abar(i) = 2 * yjaij - amix
            qbar(i) = q_val * (1 + (abar(i) / amix) - (b(i) / bmix))
            dgvinitialdata.Rows(i).Cells(8).Value = Math.Round(abar(i), 2)
            dgvinitialdata.Rows(i).Cells(9).Value = Math.Round(qbar(i), 4)
        Next

        If rdbliquid.Checked Then
            Z = beta
            Do
                Zold = Z
                Z = beta + Zold * (Zold + beta) * ((1 + beta - Zold) / (q_val * beta))
                iter += 1
                If iter > 1000 Then Exit Do
            Loop Until Math.Abs(Z - Zold) < 0.000001
        Else
            Z = 1
            Do
                Zold = Z
                Z = 1 + beta - (q_val * beta) * ((Zold - beta) / (Zold * (Zold + beta)))
                iter += 1
                If iter > 1000 Then Exit Do
            Loop Until Math.Abs(Z - Zold) < 0.000001
        End If

        If iter >= 1000 Then
            MessageBox.Show("Z failed to converge. The phase might be unstable.")
            Exit Sub
        End If

        dblI = Math.Log((Z + beta) / Z)

        For i As Integer = 0 To n - 1
            phi(i) = Math.Exp(((b(i) / bmix) * (Z - 1) - Math.Log(Z - beta) - qbar(i) * dblI))
            dgvinitialdata.Rows(i).Cells(10).Value = Math.Round(phi(i), 6)
        Next

        MessageBox.Show("Calculated")

    Catch ex As Exception
        MessageBox.Show("An error occurred during calculation: " & ex.Message)
    Finally
        btncalculate.Enabled = True
    End Try
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
