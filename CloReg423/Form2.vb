Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim result As New System.Text.StringBuilder
        For Each tb In Panel1.Controls.OfType(Of TextBox).OrderBy(Function(tbx)
                                                                      Return tbx.Name
                                                                  End Function)
            If tb.TextLength > 0 Then
                If result.Length > 0 Then result.Append(", ")
                result.Append(tb.Text)
                result.Append("% ")
                result.Append(Panel1.Controls.Item("Label" & tb.Name.Chars(tb.Name.Length - 1)).Text)
            End If
        Next
        Textboxfinalresult.Text = result.ToString



        Dim material As String
        material = Textboxfinalresult.Text

        If Form1.TabControl1.SelectedIndex = 0 Then
            Form1.NoteringTextBox.Text = material
        End If

        If Form1.TabControl1.SelectedIndex = 1 Then
            Form1.MaterialTextBoxByxor.Text = material
        End If

        If Form1.TabControl1.SelectedIndex = 2 Then
            Form1.MaterialTextBoxJacka.Text = material
        End If

        If Form1.TabControl1.SelectedIndex = 3 Then
            Form1.MaterialTextBoxKjol.Text = material
        End If

        If Form1.TabControl1.SelectedIndex = 4 Then
            Form1.MaterialTextBoxKlanning.Text = material
        End If

        If Form1.TabControl1.SelectedIndex = 5 Then
            Form1.MaterialTextBoxToppar.Text = material
        End If

        If Form1.TabControl1.SelectedIndex = 6 Then
            Form1.MaterialTextBoxJumpsuit.Text = material
        End If

        ''not used yet (skor)
        'If Form1.TabControl1.SelectedIndex = 6 Then
        '    Form1.placeholder.Text = material
        'End If

        Form1.Focus()
        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If Form1.TabControl1.SelectedIndex = 0 Then
            Form1.NoteringTextBox.Text = Button2.Text
        End If

        If Form1.TabControl1.SelectedIndex = 1 Then
            Form1.MaterialTextBoxByxor.Text = Button2.Text
        End If

        If Form1.TabControl1.SelectedIndex = 2 Then
            Form1.MaterialTextBoxJacka.Text = Button2.Text
        End If

        If Form1.TabControl1.SelectedIndex = 3 Then
            Form1.MaterialTextBoxKjol.Text = Button2.Text
        End If

        If Form1.TabControl1.SelectedIndex = 4 Then
            Form1.MaterialTextBoxKlanning.Text = Button2.Text
        End If

        If Form1.TabControl1.SelectedIndex = 5 Then
            Form1.MaterialTextBoxToppar.Text = Button2.Text
        End If

        If Form1.TabControl1.SelectedIndex = 6 Then
            Form1.MaterialTextBoxJumpsuit.Text = Button2.Text
        End If

        Form1.Focus()
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If Form1.TabControl1.SelectedIndex = 0 Then
            Form1.NoteringTextBox.Text = Button3.Text
        End If

        If Form1.TabControl1.SelectedIndex = 1 Then
            Form1.MaterialTextBoxByxor.Text = Button3.Text
        End If

        If Form1.TabControl1.SelectedIndex = 2 Then
            Form1.MaterialTextBoxJacka.Text = Button3.Text
        End If

        If Form1.TabControl1.SelectedIndex = 3 Then
            Form1.MaterialTextBoxKjol.Text = Button3.Text
        End If

        If Form1.TabControl1.SelectedIndex = 4 Then
            Form1.MaterialTextBoxKlanning.Text = Button3.Text
        End If

        If Form1.TabControl1.SelectedIndex = 5 Then
            Form1.MaterialTextBoxToppar.Text = Button3.Text
        End If

        If Form1.TabControl1.SelectedIndex = 6 Then
            Form1.MaterialTextBoxJumpsuit.Text = Button3.Text
        End If

        Form1.Focus()
        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If Form1.TabControl1.SelectedIndex = 0 Then
            Form1.NoteringTextBox.Text = Button4.Text
        End If

        If Form1.TabControl1.SelectedIndex = 1 Then
            Form1.MaterialTextBoxByxor.Text = Button4.Text
        End If

        If Form1.TabControl1.SelectedIndex = 2 Then
            Form1.MaterialTextBoxJacka.Text = Button4.Text
        End If

        If Form1.TabControl1.SelectedIndex = 3 Then
            Form1.MaterialTextBoxKjol.Text = Button4.Text
        End If

        If Form1.TabControl1.SelectedIndex = 4 Then
            Form1.MaterialTextBoxKlanning.Text = Button4.Text
        End If

        If Form1.TabControl1.SelectedIndex = 5 Then
            Form1.MaterialTextBoxToppar.Text = Button4.Text
        End If

        If Form1.TabControl1.SelectedIndex = 6 Then
            Form1.MaterialTextBoxJumpsuit.Text = Button4.Text
        End If

        Form1.Focus()
        Me.Close()

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBoxA.Text = 100
        Else
            TextBoxA.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBoxB.Text = 100
        Else
            TextBoxB.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBoxC.Text = 100
        Else
            TextBoxC.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TextBoxD.Text = 100
        Else
            TextBoxD.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            TextBoxE.Text = 100
        Else
            TextBoxE.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            TextBoxF.Text = 100
        Else
            TextBoxF.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            TextBoxG.Text = 100
        Else
            TextBoxG.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            TextBoxH.Text = 100
        Else
            TextBoxH.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            TextBoxI.Text = 100
        Else
            TextBoxI.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            TextBoxJ.Text = 100
        Else
            TextBoxJ.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            TextBoxK.Text = 100
        Else
            TextBoxK.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            TextBoxL.Text = 100
        Else
            TextBoxL.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            TextBoxM.Text = 100
        Else
            TextBoxM.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        If CheckBox14.Checked = True Then
            TextBoxN.Text = 100
        Else
            TextBoxN.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = True Then
            TextBoxO.Text = 100
        Else
            TextBoxO.Text = Nothing
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            TextBoxP.Text = 100
        Else
            TextBoxP.Text = Nothing
        End If
    End Sub
End Class