Public Class Frm1
    Dim palabra As String = "PALABRA"
    Dim vErrors As Integer = 0
    Dim vCount As Integer = 0
    Dim aWord(palabra.Length) As String
    Public Function CambiaColor(vLetra As String)
        Select Case vLetra
            Case "A"
                LblA.ForeColor = Color.Red
            Case "B"
                LblB.ForeColor = Color.Red
            Case "C"
                LblC.ForeColor = Color.Red
            Case "D"
                LblD.ForeColor = Color.Red
            Case "E"
                LblE.ForeColor = Color.Red
            Case "F"
                LblF.ForeColor = Color.Red
            Case "G"
                LblG.ForeColor = Color.Red
            Case "H"
                LblH.ForeColor = Color.Red
            Case "I"
                LblI.ForeColor = Color.Red
            Case "J"
                LblJ.ForeColor = Color.Red
            Case "K"
                LblK.ForeColor = Color.Red
            Case "L"
                LblL.ForeColor = Color.Red
            Case "M"
                LblM.ForeColor = Color.Red
            Case "N"
                LblN.ForeColor = Color.Red
            Case "O"
                LblO.ForeColor = Color.Red
            Case "P"
                LblP.ForeColor = Color.Red
            Case "Q"
                LblQ.ForeColor = Color.Red
            Case "R"
                LblR.ForeColor = Color.Red
            Case "S"
                LblS.ForeColor = Color.Red
            Case "T"
                LblT.ForeColor = Color.Red
            Case "U"
                LblU.ForeColor = Color.Red
            Case "V"
                LblV.ForeColor = Color.Red
            Case "W"
                LblW.ForeColor = Color.Red
            Case "X"
                LblX.ForeColor = Color.Red
            Case "Y"
                LblY.ForeColor = Color.Red
            Case "Z"
                LblZ.ForeColor = Color.Red
        End Select
    End Function
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles LblA.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles LblC.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles LblH.Click

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles LblL.Click

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles LblM.Click

    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles LblW.Click

    End Sub

    Private Sub Frm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For x As Integer = 0 To (palabra.Length - 1)
            aWord(x) = palabra.Substring(x, 1)
        Next
    End Sub

    Private Sub BtnComprobar_Click(sender As Object, e As EventArgs) Handles BtnComprobar.Click
        If txtComprobar.Text IsNot "" Then
            Dim letra As String = UCase(txtComprobar.Text)
            CambiaColor(letra)

            If palabra.IndexOf(letra) > -1 Then
                For x As Integer = 0 To (palabra.Length - 1)
                    If letra.Equals(aWord(x)) Then
                        Select Case x
                            Case 0
                                Lbl1.Text = letra
                            Case 1
                                Lbl2.Text = letra
                            Case 2
                                Lbl3.Text = letra
                            Case 3
                                Lbl4.Text = letra
                            Case 4
                                Lbl5.Text = letra
                            Case 5
                                Lbl6.Text = letra
                            Case 6
                                Lbl7.Text = letra
                        End Select
                    End If
                Next
            Else
                vErrors += 1
                txtErrores.Text = CStr(vErrors)
                PBAhorcado.ImageLocation = Application.StartupPath & "/img/" & vErrors & ".jpg"
                If vErrors > 4 Then
                    BtnComprobar.Enabled = False
                    txtComprobar.Enabled = False
                End If
            End If
            txtComprobar.Text = ""
        End If
    End Sub

    Private Sub txtComprobar_TextChanged(sender As Object, e As EventArgs) Handles txtComprobar.TextChanged
        If vCount < 1 Then
            If txtComprobar.Text = "" Then
                BtnComprobar.Enabled = False
            Else
                BtnComprobar.Enabled = True
            End If
            vCount += 1
        End If
    End Sub
End Class
