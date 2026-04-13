Public Class Form1
    Dim tc() As Doubleeeee
    Dim tr() As Double
    Dim pc() As Double
    Dim y() As Double
    Dim aij(,) As Double
    Dim q() As Double
    Dim a() As Double
    Dim b() As Double
    Dim qbar() As Double
    Dim phi() As Double
    Dim abar() As Double

    Private Sub Initialize_Click(sender As Object, e As EventArgs) Handles btninitialize.Click
        Dim n As Integer

        Try
            n = Integer.Parse(txtn.Text)

            Select Case True
                Case n <= 0
                    MessageBox.Show("Number of components must be greater than zero.", "Positive non-zero values.")
                    Exit Sub
                Case n > 5
                    MessageBox.Show("Number components must be less than five.", "Maximum number of components.")
                    Exit Sub
            End Select


            For i As Integer = 1 To n
                dgvinitialdata.Rows.Add(i, "", "", "", "", "", "")
            Next

            If rdbliquid.Checked = True Then
                dgvinitialdata.Columns(2).HeaderText = "x₁"
            ElseIf rdbvapor.Checked = True Then
                dgvinitialdata.Columns(2).HeaderText = "y₁"
            End If
        Catch ex As Exception
            MessageBox.Show("Enter a valid number of components.")
        End Try
    End Sub

    Private Sub btninput_Click(sender As Object, e As EventArgs) Handles btninput.Click


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btncalculate.Click
        Dim n As Integer
        Dim dbltemp, dblpressure As Double
        Dim amix, bmix As Double
        Dim q, beta As Double
        Dim yjaij As Double
        Dim dblI As Double
        Dim Z As Double
        Dim Zold As Double
        Dim count As Integer

        n = Integer.Parse(txtn.Text)
        dblpressure = txtp.Text
        dbltemp = txtT.Text
        txtn.Enabled = False
        txtp.Enabled = False
        txtT.Enabled = False
        ReDim y(n - 1)
        ReDim a(n - 1)
        ReDim b(n - 1)
        ReDim aij(n - 1, n - 1)
        ReDim tr(n - 1)
        ReDim tc(n - 1)
        ReDim pc(n - 1)
        ReDim phi(n - 1)
        ReDim abar(n - 1)
        ReDim qbar(n - 1)
        For i As Integer = 0 To n - 1
            y(i) = dgvinitialdata.Rows(i).Cells(2).Value
            tc(i) = dgvinitialdata.Rows(i).Cells(3).Value
            pc(i) = dgvinitialdata.Rows(i).Cells(4).Value
            tr(i) = dbltemp / tc(i)
            ' computes for ai and bi of each component
            a(i) = (0.42748 * tr(i) ^ -0.5 * (83.14) ^ 2 * tc(i) ^ 2) / pc(i)
            b(i) = (0.08664 * 83.14 * tc(i)) / pc(i)
        Next

        'Display Tri, ai, bi computed into datagridview
        For i As Integer = 0 To n - 1
            dgvinitialdata.Rows(i).Cells(5).Value = tr(i)
            dgvinitialdata.Rows(i).Cells(6).Value = a(i)
            dgvinitialdata.Rows(i).Cells(7).Value = b(i)
        Next
        'Computes for bmix
        For i As Integer = 0 To n - 1
            bmix += y(i) * b(i)
        Next
        'Computes for aij
        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                aij(i, j) = Math.Sqrt(a(i) * a(j))
            Next
        Next
        'Computes for amix
        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                amix += y(i) * y(j) * aij(i, j)
            Next
        Next

        TextBox2.Text = amix
        'Computes for q
        q = amix / (bmix * 83.14 * dbltemp)
        'Computes for beta
        beta = (bmix * dblpressure) / (83.14 * dbltemp)
        TextBox3.Text = q

        'Computes for abari for each component
        For i As Integer = 0 To n - 1
            yjaij = 0

            For j As Integer = 0 To n - 1
                yjaij += y(j) * aij(i, j)
            Next

            abar(i) = 2 * yjaij - amix
        Next

        'Computes for qbari for each component
        For i As Integer = 0 To n - 1
            qbar(i) = q * (1 + (abar(i) / amix) - (b(i) / bmix))
        Next
        For i As Integer = 0 To n - 1
            dgvinitialdata.Rows(i).Cells(8).Value = qbar(i)
        Next

        'Computes for Z

        Select Case True
            Case rdbliquid.Checked = True
                Z = 1
                Do
                    Zold = Z
                    Z = beta + Zold * (Zold + beta) * ((1 + beta - Zold) / (q * beta))
                    count += 1
                Loop Until Math.Abs(Z - Zold) < 0.000001

            Case rdbvapor.Checked = True
                Z = 1

                Do
                    Zold = Z
                    Z = 1 + beta - (q * beta) * ((Zold - beta) / (Zold * (Zold + beta)))
                    count += 1
                Loop Until Math.Abs(Z - Zold) < 0.000001
        End Select
        TextBox4.Text = Z
        'Computes for I
        dblI = Math.Log((Z + beta) / Z)

        'Computes for phi(fugacity coefficients)
        For i As Integer = 0 To n - 1
            phi(i) = Math.Exp(((b(i) / bmix) * (Z - 1) - Math.Log(Z - beta) - qbar(i) * dblI))
            'Displays results in the datagrid view
            dgvinitialdata.Rows(i).Cells(9).Value = phi(i)
        Next

    End Sub


End Class
