Imports System
Imports System.IO

Public Class Form1
    Private Sub Dam_ToppBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        Me.Validate()
        Me.Dam_ToppBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.CLOREGDBDataSet)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Dam_Jumpsuit' table. You can move, or remove it, as needed.
        Me.Dam_JumpsuitTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Jumpsuit)
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Dam_Byxor' table. You can move, or remove it, as needed.
        Me.Dam_ByxorTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Byxor)
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Dam_Jacka' table. You can move, or remove it, as needed.
        Me.Dam_JackaTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Jacka)
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Dam_Jacka' table. You can move, or remove it, as needed.
        Me.Dam_JackaTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Jacka)
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Accessoarer' table. You can move, or remove it, as needed.
        Me.AccessoarerTableAdapter.Fill(Me.CLOREGDBDataSet.Accessoarer)
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Dam_Klänning' table. You can move, or remove it, as needed.
        Me.Dam_KlänningTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Klänning)
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Dam_Kjol' table. You can move, or remove it, as needed.
        Me.Dam_KjolTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Kjol)
        'TODO: This line of code loads data into the 'CLOREGDBDataSet.Dam_Topp' table. You can move, or remove it, as needed.
        Me.Dam_ToppTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Topp)

        Panel1.Enabled = False


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ComboBox1.Text = Nothing Then
            MsgBox("Sorry, du måste ange typ av topp (linne, kortärmat, etc)")
            Exit Sub
        End If

        Panel6.Enabled = True

        Button2.Enabled = True
        Button100.Enabled = True
        Button1.Enabled = False

        PlaggTextBoxToppar.Text = ComboBox1.Text

        Me.Dam_ToppBindingSource.AddNew()

        Dim newitem As Int16
        Dam_ToppBindingSource.MoveLast()
        newitem = ArtikelnrTextBoxToppar.Text.Remove(0, 3) + 1

        If ForvaldhyllaTextBoxTopp.Text IsNot Nothing Then
            HyllaTextBoxToppar.Text = ForvaldhyllaTextBoxTopp.Text
        End If

        If ForvaldlevComboBoxTopp.Text IsNot Nothing Then
            InlagdTextBox4.Text = ForvaldlevComboBoxTopp.Text
        End If

        ArtikelnrTextBoxToppar.Text = "Dtp" & newitem

        ArtikelnrTextBoxToppar.Focus()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If PlaggTextBoxToppar.Text = Nothing Then
            MsgBox("Typ av topp måste anges!")
            Exit Sub
        End If

        SetSizeBtnTopp.PerformClick()
        CalcPriceBtnTopp.PerformClick()

        'Sätter storleksangivelse i slutet av produktnamn/rubrik
        'Dim tempstring As String
        'Dim old As String = " - " & Vår_storlekTextBoxTopp.Text
        'Dim removeold As String = SV_BeskrivningTextBox.Text.Replace(old, Nothing)
        'tempstring = removeold
        'SV_BeskrivningTextBox.Text = removeold
        'tempstring = SV_BeskrivningTextBox.Text & " - " & Vår_storlekTextBoxTopp.Text
        'SV_BeskrivningTextBox.Text = tempstring

        Me.Validate()
        Me.Dam_ToppBindingSource.EndEdit()
        Me.Dam_ToppTableAdapter.Update(Me.CLOREGDBDataSet)
        'Me.TableAdapterManager.UpdateAll(Me.CLOREGDBDataSet)

        Panel6.Enabled = False

        Button2.Enabled = False
        Button100.Enabled = False
        Button1.Enabled = True
        Button107.Enabled = True

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles SetSizeBtnTopp.Click

        ' Nya storleksuppskattning
        Try

            If BystTextBoxTopparar.Text <> Nothing Then

                Dim toppus As Integer = BystTextBoxTopparar.Text
                Dim toppgenomsnitt As Integer

                toppgenomsnitt = toppus * 2

                'MsgBox(toppgenomsnitt)

                '30/32 - X-small
                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxTopp.Text = "30/32 - X-small"
                End If

                '34/34 - Small
                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxTopp.Text = "34/36 - Small"
                End If

                '38/40 - Medium
                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
                    Vår_storlekTextBoxTopp.Text = "38/40 - Medium"
                End If

                '42-44 - Large
                If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
                    Vår_storlekTextBoxTopp.Text = "42/44 - Large"
                End If

                '46/48 - X-large
                If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
                    Vår_storlekTextBoxTopp.Text = "46/48 - X-large"
                End If

                '50/52 - 2X-large
                If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxTopp.Text = "50/52 - 2X-large"
                End If

            End If


            If BystTextBoxTopparar.Text = Nothing Then

                Dim toppus As Integer
                Dim toppms As Integer
                Dim toppgenomsnitt As Integer

                toppus = Convert.ToInt32(Byst_usTextBoxToppar.Text)
                toppms = Convert.ToInt32(Byst_msTextBoxToppar.Text)

                toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


                'MsgBox(toppgenomsnitt)

                '30/32 - X-small
                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxTopp.Text = "30/32 - X-small"
                End If

                '34/34 - Small
                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxTopp.Text = "34/36 - Small"
                End If

                '38/40 - Medium
                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
                    Vår_storlekTextBoxTopp.Text = "38/40 - Medium"
                End If

                '42-44 - Large
                If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
                    Vår_storlekTextBoxTopp.Text = "42/44 - Large"
                End If

                '46/48 - X-large
                If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
                    Vår_storlekTextBoxTopp.Text = "46/48 - X-large"
                End If

                '50/52 - 2X-large
                If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxTopp.Text = "50/52 - 2X-large"
                End If

            End If

        Catch ex As Exception
            MsgBox("Fälten för byst verkar tomma? Fälten bör innehålla antingen mått på byst med och utan stretch alternativt vanligt mått för icke-stretchplagg")
        End Try




        ' Gamla storleksuppskattning
        'Try

        '    If BystTextBoxTopparar.Text <> Nothing Then

        '        Dim toppus As Integer = BystTextBoxTopparar.Text
        '        Dim toppgenomsnitt As Integer

        '        toppgenomsnitt = toppus * 2

        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 72) Then
        '            Vår_storlekTextBoxTopp.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 73) And (toppgenomsnitt <= 77) Then
        '            Vår_storlekTextBoxTopp.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 78) And (toppgenomsnitt <= 86) Then
        '            Vår_storlekTextBoxTopp.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 88) Then
        '            Vår_storlekTextBoxTopp.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 89) And (toppgenomsnitt <= 97) Then
        '            Vår_storlekTextBoxTopp.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 100) Then
        '            Vår_storlekTextBoxTopp.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 101) And (toppgenomsnitt <= 109) Then
        '            Vår_storlekTextBoxTopp.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 110) And (toppgenomsnitt <= 111) Then
        '            Vår_storlekTextBoxTopp.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 112) And (toppgenomsnitt <= 200) Then
        '            Vår_storlekTextBoxTopp.Text = "Large - XL"
        '        End If

        '    End If


        '    If BystTextBoxTopparar.Text = Nothing Then

        '        Dim toppus As Integer
        '        Dim toppms As Integer
        '        Dim toppgenomsnitt As Integer

        '        toppus = Convert.ToInt32(Byst_usTextBoxToppar.Text)
        '        toppms = Convert.ToInt32(Byst_msTextBoxToppar.Text)

        '        toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 72) Then
        '            Vår_storlekTextBoxTopp.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 73) And (toppgenomsnitt <= 77) Then
        '            Vår_storlekTextBoxTopp.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 78) And (toppgenomsnitt <= 86) Then
        '            Vår_storlekTextBoxTopp.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 88) Then
        '            Vår_storlekTextBoxTopp.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 89) And (toppgenomsnitt <= 97) Then
        '            Vår_storlekTextBoxTopp.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 100) Then
        '            Vår_storlekTextBoxTopp.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 101) And (toppgenomsnitt <= 109) Then
        '            Vår_storlekTextBoxTopp.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 110) And (toppgenomsnitt <= 111) Then
        '            Vår_storlekTextBoxTopp.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 112) And (toppgenomsnitt <= 200) Then
        '            Vår_storlekTextBoxTopp.Text = "Large - XL"
        '        End If

        '    End If

        'Catch ex As Exception
        '    MsgBox("Fälten för byst verkar tomma? Fälten bör innehålla antingen mått på byst med och utan stretch alternativt vanligt mått för icke-stretchplagg")
        'End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        BilderListBoxToppar.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                BilderListBoxToppar.Items.Add(foundFile.Substring(foundFile.Length - 12))

            Next
        Catch ex As Exception

        End Try

        If BilderListBoxToppar.Items.Count.ToString = 0 Then
            BilderTextBoxToppar.Text = Nothing
        End If

        If BilderListBoxToppar.Items.Count.ToString = 1 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 2 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 3 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 4 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 5 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 6 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 7 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 8 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString & ";" & BilderListBoxToppar.Items(7).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 9 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString & ";" & BilderListBoxToppar.Items(7).ToString & ";" & BilderListBoxToppar.Items(8).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 10 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString & ";" & BilderListBoxToppar.Items(7).ToString & ";" & BilderListBoxToppar.Items(8).ToString & ";" & BilderListBoxToppar.Items(89).ToString
        End If

        BilderTextBoxToppar.Text.Replace("G:\Produktbilder\", Nothing)

        BilderListBoxToppar.Items.Clear()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles WriteImagesTopparBtn.Click
        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        BilderListBoxToppar.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                BilderListBoxToppar.Items.Add(foundFile.Substring(foundFile.Length - 12))

            Next
        Catch ex As Exception

        End Try

        If BilderListBoxToppar.Items.Count.ToString = 0 Then
            BilderTextBoxToppar.Text = Nothing
        End If

        If BilderListBoxToppar.Items.Count.ToString = 1 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 2 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 3 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 4 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 5 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 6 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 7 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 8 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString & ";" & BilderListBoxToppar.Items(7).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 9 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString & ";" & BilderListBoxToppar.Items(7).ToString & ";" & BilderListBoxToppar.Items(8).ToString
        End If

        If BilderListBoxToppar.Items.Count.ToString = 10 Then
            BilderTextBoxToppar.Text = BilderListBoxToppar.Items(0).ToString & ";" & BilderListBoxToppar.Items(1).ToString & ";" & BilderListBoxToppar.Items(2).ToString & ";" & BilderListBoxToppar.Items(3).ToString & ";" & BilderListBoxToppar.Items(4).ToString & ";" & BilderListBoxToppar.Items(5).ToString & ";" & BilderListBoxToppar.Items(6).ToString & ";" & BilderListBoxToppar.Items(7).ToString & ";" & BilderListBoxToppar.Items(8).ToString & ";" & BilderListBoxToppar.Items(89).ToString
        End If

        BilderTextBoxToppar.Text.Replace("G:\Produktbilder\", Nothing)

        BilderListBoxToppar.Items.Clear()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dam_ToppBindingSource.MoveFirst()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dam_ToppBindingSource.MovePrevious()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dam_ToppBindingSource.MoveNext()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dam_ToppBindingSource.MoveLast()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        For Each dataRow As DataRowView In Dam_ToppListBoxToppar.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        MsgBox("Klart! Mappar skappade.")

        'Dim path As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text

        'Try
        '    ' Determine whether the directory exists.
        '    If Directory.Exists(path) Then
        '        MsgBox("Mapp finns redan")
        '        Exit Sub
        '    End If

        '    ' Try to create the directory.
        '    Dim di As DirectoryInfo = Directory.CreateDirectory(path)
        '    MsgBox("Mapp skapad")


        'Catch a As Exception
        '    Console.WriteLine("The process failed: {0}.", a.ToString())
        'End Try

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles SetSizeBtnKjol.Click

        Try

            If MidjaTextBoxKjol.Text <> Nothing Then

                Dim toppus As Integer = MidjaTextBoxKjol.Text
                Dim toppgenomsnitt As Integer

                toppgenomsnitt = toppus * 2

                'MsgBox(toppgenomsnitt)


                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 60) Then
                    Vår_storkelTextBoxKjol.Text = "30/32 - X-small"
                End If


                If (toppgenomsnitt >= 61) And (toppgenomsnitt <= 70) Then
                    Vår_storkelTextBoxKjol.Text = "34/36 - Small"
                End If


                If (toppgenomsnitt >= 71) And (toppgenomsnitt <= 78) Then
                    Vår_storkelTextBoxKjol.Text = "38/40 - Medium"
                End If


                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storkelTextBoxKjol.Text = "42/44 - Large"
                End If


                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 97) Then
                    Vår_storkelTextBoxKjol.Text = "46/48 - X-large"
                End If


                If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 200) Then
                    Vår_storkelTextBoxKjol.Text = "50/52 - 2X-large"
                End If

            End If


            If MidjaTextBoxKjol.Text = Nothing Then

                Dim toppus As Integer
                Dim toppms As Integer
                Dim toppgenomsnitt As Integer

                toppus = Convert.ToInt32(Midja_usTextBoxKjol.Text)
                toppms = Convert.ToInt32(Midja_msTextBoxKjol.Text)

                toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


                'MsgBox(toppgenomsnitt)


                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 60) Then
                    Vår_storkelTextBoxKjol.Text = "30/32 - X-small"
                End If


                If (toppgenomsnitt >= 61) And (toppgenomsnitt <= 70) Then
                    Vår_storkelTextBoxKjol.Text = "34/36 - Small"
                End If


                If (toppgenomsnitt >= 71) And (toppgenomsnitt <= 78) Then
                    Vår_storkelTextBoxKjol.Text = "38/40 - Medium"
                End If


                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storkelTextBoxKjol.Text = "42/44 - Large"
                End If


                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 97) Then
                    Vår_storkelTextBoxKjol.Text = "46/48 - X-large"
                End If


                If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 200) Then
                    Vår_storkelTextBoxKjol.Text = "50/52 - 2X-large"
                End If

            End If

        Catch ex As Exception
            MsgBox("Fälten för midja verkar tomma? Fälten bör innehålla antingen mått på midja med och utan stretch alternativt vanligt midjemått för icke-stretchplagg")
        End Try




        'Gammal storleksuppskattning
        'Try

        '    If MidjaTextBoxKjol.Text <> Nothing Then

        '        Dim toppus As Integer = MidjaTextBoxKjol.Text
        '        Dim toppgenomsnitt As Integer

        '        toppgenomsnitt = toppus * 2

        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 58) Then
        '            Vår_storkelTextBoxKjol.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 59) And (toppgenomsnitt <= 60) Then
        '            Vår_storkelTextBoxKjol.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 61) And (toppgenomsnitt <= 69) Then
        '            Vår_storkelTextBoxKjol.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 70) And (toppgenomsnitt <= 70) Then
        '            Vår_storkelTextBoxKjol.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 71) And (toppgenomsnitt <= 79) Then
        '            Vår_storkelTextBoxKjol.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 80) And (toppgenomsnitt <= 81) Then
        '            Vår_storkelTextBoxKjol.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 82) And (toppgenomsnitt <= 90) Then
        '            Vår_storkelTextBoxKjol.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 91) And (toppgenomsnitt <= 92) Then
        '            Vår_storkelTextBoxKjol.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 93) And (toppgenomsnitt <= 200) Then
        '            Vår_storkelTextBoxKjol.Text = "Large - XL"
        '        End If

        '    End If


        '    If MidjaTextBoxKjol.Text = Nothing Then

        '        Dim toppus As Integer
        '        Dim toppms As Integer
        '        Dim toppgenomsnitt As Integer

        '        toppus = Convert.ToInt32(Midja_usTextBoxKjol.Text)
        '        toppms = Convert.ToInt32(Midja_msTextBoxKjol.Text)

        '        toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 58) Then
        '            Vår_storkelTextBoxKjol.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 59) And (toppgenomsnitt <= 60) Then
        '            Vår_storkelTextBoxKjol.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 61) And (toppgenomsnitt <= 69) Then
        '            Vår_storkelTextBoxKjol.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 70) And (toppgenomsnitt <= 70) Then
        '            Vår_storkelTextBoxKjol.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 71) And (toppgenomsnitt <= 79) Then
        '            Vår_storkelTextBoxKjol.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 81) And (toppgenomsnitt <= 81) Then
        '            Vår_storkelTextBoxKjol.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 82) And (toppgenomsnitt <= 90) Then
        '            Vår_storkelTextBoxKjol.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 91) And (toppgenomsnitt <= 92) Then
        '            Vår_storkelTextBoxKjol.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 93) And (toppgenomsnitt <= 200) Then
        '            Vår_storkelTextBoxKjol.Text = "Large - XL"
        '        End If

        '    End If

        'Catch ex As Exception
        '    MsgBox("Fälten för midja verkar tomma? Fälten bör innehålla antingen mått på midja med och utan stretch alternativt vanligt midjemått för icke-stretchplagg")
        'End Try
    End Sub

    Private Sub WriteImagesKjolBtn_Click(sender As Object, e As EventArgs) Handles WriteImagesKjolBtn.Click
        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxKjol.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        BilderListBoxKjol.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                BilderListBoxKjol.Items.Add(foundFile.Substring(foundFile.Length - 12))

            Next
        Catch ex As Exception

        End Try

        If BilderListBoxKjol.Items.Count.ToString = 0 Then
            BilderTextBoxKjol.Text = Nothing
        End If

        If BilderListBoxKjol.Items.Count.ToString = 1 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 2 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 3 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 4 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString & ";" & BilderListBoxKjol.Items(3).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 5 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString & ";" & BilderListBoxKjol.Items(3).ToString & ";" & BilderListBoxKjol.Items(4).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 6 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString & ";" & BilderListBoxKjol.Items(3).ToString & ";" & BilderListBoxKjol.Items(4).ToString & ";" & BilderListBoxKjol.Items(5).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 7 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString & ";" & BilderListBoxKjol.Items(3).ToString & ";" & BilderListBoxKjol.Items(4).ToString & ";" & BilderListBoxKjol.Items(5).ToString & ";" & BilderListBoxKjol.Items(6).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 8 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString & ";" & BilderListBoxKjol.Items(3).ToString & ";" & BilderListBoxKjol.Items(4).ToString & ";" & BilderListBoxKjol.Items(5).ToString & ";" & BilderListBoxKjol.Items(6).ToString & ";" & BilderListBoxKjol.Items(7).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 9 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString & ";" & BilderListBoxKjol.Items(3).ToString & ";" & BilderListBoxKjol.Items(4).ToString & ";" & BilderListBoxKjol.Items(5).ToString & ";" & BilderListBoxKjol.Items(6).ToString & ";" & BilderListBoxKjol.Items(7).ToString & ";" & BilderListBoxKjol.Items(8).ToString
        End If

        If BilderListBoxKjol.Items.Count.ToString = 10 Then
            BilderTextBoxKjol.Text = BilderListBoxKjol.Items(0).ToString & ";" & BilderListBoxKjol.Items(1).ToString & ";" & BilderListBoxKjol.Items(2).ToString & ";" & BilderListBoxKjol.Items(3).ToString & ";" & BilderListBoxKjol.Items(4).ToString & ";" & BilderListBoxKjol.Items(5).ToString & ";" & BilderListBoxKjol.Items(6).ToString & ";" & BilderListBoxKjol.Items(7).ToString & ";" & BilderListBoxKjol.Items(8).ToString & ";" & BilderListBoxKjol.Items(89).ToString
        End If

        BilderTextBoxKjol.Text.Replace("G:\Produktbilder\", Nothing)

        BilderListBoxKjol.Items.Clear()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles SetSizeBtnKlanning.Click

        Try

            If BystTextBoxKlanning.Text <> Nothing Then

                Dim toppus As Integer = BystTextBoxKlanning.Text
                Dim toppgenomsnitt As Integer

                toppgenomsnitt = toppus * 2

                'MsgBox(toppgenomsnitt)

                '30/32 - X-small
                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxKlanning.Text = "30/32 - X-small"
                End If

                '34/34 - Small
                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxKlanning.Text = "34/36 - Small"
                End If

                '38/40 - Medium
                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
                    Vår_storlekTextBoxKlanning.Text = "38/40 - Medium"
                End If

                '42-44 - Large
                If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
                    Vår_storlekTextBoxKlanning.Text = "42/44 - Large"
                End If

                '46/48 - X-large
                If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
                    Vår_storlekTextBoxKlanning.Text = "46/48 - X-large"
                End If

                '50/52 - 2X-large
                If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxKlanning.Text = "50/52 - 2X-large"
                End If

            End If


            If BystTextBoxKlanning.Text = Nothing Then

                Dim toppus As Integer
                Dim toppms As Integer
                Dim toppgenomsnitt As Integer

                toppus = Convert.ToInt32(Byst_usTextBoxKlanning.Text)
                toppms = Convert.ToInt32(Byst_msTextBoxKlanning.Text)

                toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


                'MsgBox(toppgenomsnitt)

                '30/32 - X-small
                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxKlanning.Text = "30/32 - X-small"
                End If

                '34/34 - Small
                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxKlanning.Text = "34/36 - Small"
                End If

                '38/40 - Medium
                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
                    Vår_storlekTextBoxKlanning.Text = "38/40 - Medium"
                End If

                '42-44 - Large
                If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
                    Vår_storlekTextBoxKlanning.Text = "42/44 - Large"
                End If

                '46/48 - X-large
                If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
                    Vår_storlekTextBoxKlanning.Text = "46/48 - X-large"
                End If

                '50/52 - 2X-large
                If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxKlanning.Text = "50/52 - 2X-large"
                End If

            End If

        Catch ex As Exception
            MsgBox("Fälten för byst verkar tomma? Fälten bör innehålla antingen mått på byst med och utan stretch alternativt vanligt mått för icke-stretchplagg")
        End Try



        'Gammal storleksuppskattning
        'Try

        '    If BystTextBoxKlanning.Text <> Nothing Then

        '        Dim toppus As Integer = BystTextBoxKlanning.Text
        '        Dim toppgenomsnitt As Integer

        '        toppgenomsnitt = toppus * 2

        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 72) Then
        '            Vår_storlekTextBoxKlanning.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 73) And (toppgenomsnitt <= 77) Then
        '            Vår_storlekTextBoxKlanning.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 78) And (toppgenomsnitt <= 86) Then
        '            Vår_storlekTextBoxKlanning.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 88) Then
        '            Vår_storlekTextBoxKlanning.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 89) And (toppgenomsnitt <= 97) Then
        '            Vår_storlekTextBoxKlanning.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 100) Then
        '            Vår_storlekTextBoxKlanning.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 101) And (toppgenomsnitt <= 109) Then
        '            Vår_storlekTextBoxKlanning.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 110) And (toppgenomsnitt <= 111) Then
        '            Vår_storlekTextBoxKlanning.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 112) And (toppgenomsnitt <= 200) Then
        '            Vår_storlekTextBoxKlanning.Text = "Large - XL"
        '        End If

        '    End If


        '    If BystTextBoxKlanning.Text = Nothing Then

        '        Dim toppus As Integer
        '        Dim toppms As Integer
        '        Dim toppgenomsnitt As Integer

        '        toppus = Convert.ToInt32(Byst_usTextBoxKlanning.Text)
        '        toppms = Convert.ToInt32(Byst_msTextBoxKlanning.Text)

        '        toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 72) Then
        '            Vår_storlekTextBoxKlanning.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 73) And (toppgenomsnitt <= 77) Then
        '            Vår_storlekTextBoxKlanning.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 78) And (toppgenomsnitt <= 86) Then
        '            Vår_storlekTextBoxKlanning.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 88) Then
        '            Vår_storlekTextBoxKlanning.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 89) And (toppgenomsnitt <= 97) Then
        '            Vår_storlekTextBoxKlanning.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 100) Then
        '            Vår_storlekTextBoxKlanning.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 101) And (toppgenomsnitt <= 109) Then
        '            Vår_storlekTextBoxKlanning.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 110) And (toppgenomsnitt <= 111) Then
        '            Vår_storlekTextBoxKlanning.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 112) And (toppgenomsnitt <= 200) Then
        '            Vår_storlekTextBoxKlanning.Text = "Large - XL"
        '        End If

        '    End If

        'Catch ex As Exception
        '    MsgBox("Fälten för byst verkar tomma? Fälten bör innehålla antingen mått på byst med och utan stretch alternativt vanligt mått för icke-stretchplagg")
        'End Try

    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxKlanning.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        BilderListBoxKlanning.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                BilderListBoxKlanning.Items.Add(foundFile.Substring(foundFile.Length - 12))

            Next
        Catch ex As Exception

        End Try

        If BilderListBoxKlanning.Items.Count.ToString = 0 Then
            BilderTextBoxKlanning.Text = Nothing
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 1 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 2 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 3 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 4 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString & ";" & BilderListBoxKlanning.Items(3).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 5 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString & ";" & BilderListBoxKlanning.Items(3).ToString & ";" & BilderListBoxKlanning.Items(4).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 6 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString & ";" & BilderListBoxKlanning.Items(3).ToString & ";" & BilderListBoxKlanning.Items(4).ToString & ";" & BilderListBoxKlanning.Items(5).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 7 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString & ";" & BilderListBoxKlanning.Items(3).ToString & ";" & BilderListBoxKlanning.Items(4).ToString & ";" & BilderListBoxKlanning.Items(5).ToString & ";" & BilderListBoxKlanning.Items(6).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 8 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString & ";" & BilderListBoxKlanning.Items(3).ToString & ";" & BilderListBoxKlanning.Items(4).ToString & ";" & BilderListBoxKlanning.Items(5).ToString & ";" & BilderListBoxKlanning.Items(6).ToString & ";" & BilderListBoxKlanning.Items(7).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 9 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString & ";" & BilderListBoxKlanning.Items(3).ToString & ";" & BilderListBoxKlanning.Items(4).ToString & ";" & BilderListBoxKlanning.Items(5).ToString & ";" & BilderListBoxKlanning.Items(6).ToString & ";" & BilderListBoxKlanning.Items(7).ToString & ";" & BilderListBoxKlanning.Items(8).ToString
        End If

        If BilderListBoxKlanning.Items.Count.ToString = 10 Then
            BilderTextBoxKlanning.Text = BilderListBoxKlanning.Items(0).ToString & ";" & BilderListBoxKlanning.Items(1).ToString & ";" & BilderListBoxKlanning.Items(2).ToString & ";" & BilderListBoxKlanning.Items(3).ToString & ";" & BilderListBoxKlanning.Items(4).ToString & ";" & BilderListBoxKlanning.Items(5).ToString & ";" & BilderListBoxKlanning.Items(6).ToString & ";" & BilderListBoxKlanning.Items(7).ToString & ";" & BilderListBoxKlanning.Items(8).ToString & ";" & BilderListBoxKlanning.Items(89).ToString
        End If

        BilderTextBoxKlanning.Text.Replace("G:\Produktbilder\", Nothing)

        BilderListBoxKlanning.Items.Clear()
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click

        SetSizeBtnKlanning.PerformClick()
        CalcPriceBtnKlanning.PerformClick()

        'Sätter storleksangivelse i slutet av produktnamn/rubrik
        'Dim tempstring As String
        'Dim old As String = " - " & Vår_storlekTextBoxKlanning.Text
        'Dim removeold As String = SV_BeskrivningTextBox2.Text.Replace(old, Nothing)
        'tempstring = removeold
        'SV_BeskrivningTextBox2.Text = removeold
        'tempstring = SV_BeskrivningTextBox2.Text & " - " & Vår_storlekTextBoxKlanning.Text
        'SV_BeskrivningTextBox2.Text = tempstring

        Me.Validate()
        Me.Dam_KlänningBindingSource.EndEdit()
        Me.Dam_KlänningTableAdapter.Update(Me.CLOREGDBDataSet)

        Panel5.Enabled = False

        Button25.Enabled = True
        Button99.Enabled = False
        Button24.Enabled = False
        Button106.Enabled = True


    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click

        Panel5.Enabled = True

        Button25.Enabled = False
        Button24.Enabled = True
        Button29.Enabled = True

        Me.Dam_KlänningBindingSource.AddNew()

        Dim newitem As Int16
        Dam_KlänningBindingSource.MoveLast()
        newitem = ArtikelnrTextBoxKlanning.Text.Remove(0, 3) + 1

        If ForvaldhyllaTextBoxKlanning.Text IsNot Nothing Then
            HyllaTextBoxKlanning.Text = ForvaldhyllaTextBoxKlanning.Text
        End If

        If ForvaldlevComboBoxKlanning.Text IsNot Nothing Then
            InlagdTextBox3.Text = ForvaldlevComboBoxKlanning.Text
        End If

        ArtikelnrTextBoxKlanning.Text = "Dkl" & newitem

        ArtikelnrTextBoxKlanning.Focus()


    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dam_KlänningBindingSource.MoveFirst()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dam_KlänningBindingSource.MovePrevious()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dam_KlänningBindingSource.MoveNext()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dam_KlänningBindingSource.MoveLast()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click

        For Each dataRow As DataRowView In Dam_KlänningListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        MsgBox("Klart! Mappar skappade.")

        'Dim path As String = "G:\Produktbilder\" & ArtikelnrTextBoxKlanning.Text

        'Try
        '    ' Determine whether the directory exists.
        '    If Directory.Exists(path) Then
        '        MsgBox("Mapp finns redan")
        '        Exit Sub
        '    End If

        '    ' Try to create the directory.
        '    Dim di As DirectoryInfo = Directory.CreateDirectory(path)
        '    MsgBox("Mapp skapad")


        'Catch a As Exception
        '    Console.WriteLine("The process failed: {0}.", a.ToString())
        'End Try

    End Sub

    Private Sub CalcPriceBtnTopp_Click(sender As Object, e As EventArgs) Handles CalcPriceBtnTopp.Click

        Dim storlekskat As String

        'Sätt kategorier för kortärmade toppar
        If PlaggTextBoxToppar.Text = "Topp kortärmad" Then

            If Vår_storlekTextBoxTopp.Text = "30/32 - X-small" Then
                storlekskat = "12|18|43|52||12|18|43||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "34/36 - Small" Then
                storlekskat = "12|18|43|53||12|18|43||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "38/40 - Medium" Then
                storlekskat = "12|18|43|54||12|18|43||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "42/44 - Large" Then
                storlekskat = "12|18|43|55||12|18|43||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "46/48 - X-large" Then
                storlekskat = "12|18|43|56||12|18|43||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "50/52 - 2X-large" Then
                storlekskat = "12|18|43|57||12|18|43||12|18||12"
            End If

        End If

        'Sätt kategorier för långärmade toppar
        If PlaggTextBoxToppar.Text = "Topp långärmad" Then
            If Vår_storlekTextBoxTopp.Text = "30/32 - X-small" Then
                storlekskat = "12|18|42|58||12|18|42||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "34/36 - Small" Then
                storlekskat = "12|18|42|59||12|18|42||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "38/40 - Medium" Then
                storlekskat = "12|18|42|60||12|18|42||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "42/44 - Large" Then
                storlekskat = "12|18|42|61||12|18|42||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "46/48 - X-large" Then
                storlekskat = "12|18|42|62||12|18|42||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "50/52 - 2X-large" Then
                storlekskat = "12|18|42|63||12|18|42||12|18||12"
            End If

        End If

        'Sätt storlek för linnen
        If PlaggTextBoxToppar.Text = "Linne" Then
            If Vår_storlekTextBoxTopp.Text = "30/32 - X-small" Then
                storlekskat = "12|18|44|64||12|18|44||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "34/36 - Small" Then
                storlekskat = "12|18|44|65||12|18|44||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "38/40 - Medium" Then
                storlekskat = "12|18|44|66||12|18|44||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "42/44 - Large" Then
                storlekskat = "12|18|44|67||12|18|44||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "46/48 - X-large" Then
                storlekskat = "12|18|44|68||12|18|44||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "50/52 - 2X-large" Then
                storlekskat = "12|18|44|69||12|18|44||12|18||12"
            End If
        End If


        'Sätt storlek för tröjor
        If PlaggTextBoxToppar.Text = "Tröja" Then
            If Vår_storlekTextBoxTopp.Text = "30/32 - X-small" Then
                storlekskat = "12|18|47|70||12|18|47||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "34/36 - Small" Then
                storlekskat = "12|18|47|71||12|18|47||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "38/40 - Medium" Then
                storlekskat = "12|18|47|72||12|18|47||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "42/44 - Large" Then
                storlekskat = "12|18|47|73||12|18|47||12|18||12"
            End If

            If Vår_storlekTextBoxTopp.Text = "46/48 - X-large" Then
                storlekskat = "12|18|47|74||12|18|47||12|18||12"

            End If

            If Vår_storlekTextBoxTopp.Text = "50/52 - 2X-large" Then
                storlekskat = "12|18|47|75||12|18|47||12|18||12"
            End If

        End If


        Try
            Pris_umTextBoxToppar.Text = Pris_mmTextBoxToppar.Text * 0.8
        Catch ex As Exception
            MsgBox("FYI: inget pris angivet")
        End Try

        ActiveradTextBoxToppar.Text = "0"
        ConditionComboBoxToppar.Text = "used"
        CategoryTextBoxToppar.Text = storlekskat
        LagersaldoTextBoxToppar.Text = "1"
        SKATTTextBoxToppar.Text = "1"
        MåttbeskrivningTextBoxToppar.Text = "Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag."
    End Sub

    Private Sub CalcPriceBtnKlanning_Click(sender As Object, e As EventArgs) Handles CalcPriceBtnKlanning.Click

        Dim storlekskat As String



        If Vår_storlekTextBoxKlanning.Text = "30/32 - X-small" Then
            storlekskat = "12|17|76||12|17||12"
        End If

        If Vår_storlekTextBoxKlanning.Text = "34/36 - Small" Then
            storlekskat = "12|17|77||12|17||12"

        End If

        If Vår_storlekTextBoxKlanning.Text = "38/40 - Medium" Then
            storlekskat = "12|17|78||12|17||12"
        End If

        If Vår_storlekTextBoxKlanning.Text = "42/44 - Large" Then
            storlekskat = "12|17|79||12|17||12"
        End If

        If Vår_storlekTextBoxKlanning.Text = "46/48 - X-large" Then
            storlekskat = "12|17|80||12|17||12"
        End If

        If Vår_storlekTextBoxKlanning.Text = "50/52 - 2X-large" Then
            storlekskat = "12|17|81||12|17||12"
        End If



        Try
            Pris_umTextBoxKlanning.Text = Pris_mmTextBoxKlanning.Text * 0.8
        Catch ex As Exception
            MsgBox("FYI: Inget pris angivet!")
        End Try

        ActiveradTextBoxKlanning.Text = "0"
        ConditionComboBoxKlanning.Text = "used"
        CategoryTextBoxKlanning.Text = storlekskat
        LagersaldoTextBoxKlanning.Text = "1"
        SKATTTextBoxKlanning.Text = "1"
        MåttbeskrivningTextBoxKlanning.Text = "Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag."
    End Sub

    Private Sub CalcPriceBtnAcc_Click(sender As Object, e As EventArgs) Handles CalcPriceBtnAcc.Click

        Try
            Pris_umTextBoxAcc.Text = Pris_mmTextBoxAcc.Text * 0.8
        Catch ex As Exception
            MsgBox("FYI: Inget pris angivet!")
        End Try

        ActiveradTextBoxAcc.Text = "0"
        ConditionComboBoxAcc.Text = "used"
        CategoryTextBoxAcc.Text = "14"
        LagersaldoTextBoxAcc.Text = "1"
        SKATTTextBoxAcc.Text = "1"
        MåttbeskrivningTextBoxAcc.Text = "Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag."
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click

        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxAcc.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        ListBoxAccessoarer.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                ListBoxAccessoarer.Items.Add(foundFile)
            Next
        Catch ex As Exception

        End Try

        'System.IO.Directory.CreateDirectory("G:\produktbilder\" + ArtikelnrTextBox6.Text)




        If ListBoxAccessoarer.Items.Count.ToString = 0 Then
            BilderTextBoxAcc.Text = Nothing
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 1 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 2 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 3 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 4 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString & ";" & ListBoxAccessoarer.Items(3).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 5 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString & ";" & ListBoxAccessoarer.Items(3).ToString & ";" & ListBoxAccessoarer.Items(4).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 6 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString & ";" & ListBoxAccessoarer.Items(3).ToString & ";" & ListBoxAccessoarer.Items(4).ToString & ";" & ListBoxAccessoarer.Items(5).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 7 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString & ";" & ListBoxAccessoarer.Items(3).ToString & ";" & ListBoxAccessoarer.Items(4).ToString & ";" & ListBoxAccessoarer.Items(5).ToString & ";" & ListBoxAccessoarer.Items(6).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 8 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString & ";" & ListBoxAccessoarer.Items(3).ToString & ";" & ListBoxAccessoarer.Items(4).ToString & ";" & ListBoxAccessoarer.Items(5).ToString & ";" & ListBoxAccessoarer.Items(6).ToString & ";" & ListBoxAccessoarer.Items(7).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 9 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString & ";" & ListBoxAccessoarer.Items(3).ToString & ";" & ListBoxAccessoarer.Items(4).ToString & ";" & ListBoxAccessoarer.Items(5).ToString & ";" & ListBoxAccessoarer.Items(6).ToString & ";" & ListBoxAccessoarer.Items(7).ToString & ";" & ListBoxAccessoarer.Items(8).ToString
        End If

        If ListBoxAccessoarer.Items.Count.ToString = 10 Then
            BilderTextBoxAcc.Text = ListBoxAccessoarer.Items(0).ToString & ";" & ListBoxAccessoarer.Items(1).ToString & ";" & ListBoxAccessoarer.Items(2).ToString & ";" & ListBoxAccessoarer.Items(3).ToString & ";" & ListBoxAccessoarer.Items(4).ToString & ";" & ListBoxAccessoarer.Items(5).ToString & ";" & ListBoxAccessoarer.Items(6).ToString & ";" & ListBoxAccessoarer.Items(7).ToString & ";" & ListBoxAccessoarer.Items(8).ToString & ";" & ListBoxAccessoarer.Items(89).ToString
        End If

        ListBoxAccessoarer.Items.Clear()
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        AccessoarerBindingSource.MoveFirst()
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        AccessoarerBindingSource.MovePrevious()
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        AccessoarerBindingSource.MoveNext()
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        AccessoarerBindingSource.MoveLast()
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click

        If ComboBox3.Text = Nothing Then
            MsgBox("Sorry, du måste ange typ (väska, handskar, etc)")
            Exit Sub
        End If

        'fundera på att implementera automatiskt inlägg av artikelnummer
        'Dim newitem As Int16
        'AccessoarerBindingSource.MoveLast()
        'newitem = ArtikelnrTextBoxAcc.Text.Remove(0, 3) + 1

        Panel1.Enabled = True
        AvbrytAccButton.Enabled = True
        Button33.Enabled = True


        Me.AccessoarerBindingSource.AddNew()

        If ForvaldhyllaAccTextBox.Text IsNot Nothing Then
            HyllaTextBoxAcc.Text = ForvaldhyllaAccTextBox.Text
        End If

        If ForvaldlevAccComboBox.Text IsNot Nothing Then
            InlagdTextBox6.Text = ForvaldlevAccComboBox.Text
        End If

        'ArtikelnrTextBoxByxor.Text = "Acc" & newitem

        ArtikelnrTextBoxAcc.Focus()



    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click


        CalcPriceBtnAcc.PerformClick()

        Me.Validate()
        Me.AccessoarerBindingSource.EndEdit()
        Me.AccessoarerTableAdapter.Update(Me.CLOREGDBDataSet)

        Panel1.Enabled = False
        AvbrytAccButton.Enabled = False
        Button102.Enabled = True


    End Sub

    Private Sub Button26_Click_1(sender As Object, e As EventArgs) Handles Button26.Click

        For Each dataRow As DataRowView In AccessoarerListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        MsgBox("Klart! Mappar skappade.")

        'Dim path As String = "G:\Produktbilder\" & ArtikelnrTextBoxAcc.Text

        'Try
        '    ' Determine whether the directory exists.
        '    If Directory.Exists(path) Then
        '        MsgBox("Mapp finns redan")
        '        Exit Sub
        '    End If

        '    ' Try to create the directory.
        '    Dim di As DirectoryInfo = Directory.CreateDirectory(path)
        '    MsgBox("Mapp skapad")


        'Catch a As Exception
        '    Console.WriteLine("The process failed: {0}.", a.ToString())
        'End Try
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        Dam_JackaBindingSource.MoveFirst()
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        Dam_JackaBindingSource.MovePrevious()
    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        Dam_JackaBindingSource.MoveNext()
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Dam_JackaBindingSource.MoveLast()
    End Sub

    Private Sub CalcPriceBtnJacka_Click(sender As Object, e As EventArgs) Handles CalcPriceBtnJacka.Click

        Dim storlekskat As String



        If Vår_storlekTextBoxJacka.Text = "30/32 - X-small" Then
            storlekskat = "12|28|112||12|28||12"
        End If

        If Vår_storlekTextBoxJacka.Text = "34/36 - Small" Then
            storlekskat = "12|28|113||12|28||12"
        End If

        If Vår_storlekTextBoxJacka.Text = "38/40 - Medium" Then
            storlekskat = "12|28|114||12|28||12"
        End If

        If Vår_storlekTextBoxJacka.Text = "42/44 - Large" Then
            storlekskat = "12|28|115||12|28||12"
        End If

        If Vår_storlekTextBoxJacka.Text = "46/48 - X-large" Then
            storlekskat = "12|28|116||12|28||12"
        End If

        If Vår_storlekTextBoxJacka.Text = "50/52 - 2X-large" Then
            storlekskat = "12|28|117||12|28||12"
        End If

        Try
            Pris_umTextBoxJacka.Text = Pris_mmTextBoxJacka.Text * 0.8
        Catch ex As Exception
            MsgBox("FYI: Inget pris angivet!")
        End Try

        ActiveradTextBoxJacka.Text = "0"
        ConditionComboBoxJacka.Text = "used"
        CategoryTextBoxJacka.Text = storlekskat
        LagersaldoTextBoxJacka.Text = 1
        SKATTTextBoxJacka.Text = "1"
        MåttbeskrivningTextBoxJacka.Text = "Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag."
    End Sub

    Private Sub CalcPriceBtnKjol_Click(sender As Object, e As EventArgs) Handles CalcPriceBtnKjol.Click

        Dim storlekskat As String

        If Vår_storkelTextBoxKjol.Text = "30/32 - X-small" Then
                storlekskat = "12|16|88||12|16||12"
            End If

        If Vår_storkelTextBoxKjol.Text = "34/36 - Small" Then
            storlekskat = "12|16|89||12|16||12"

        End If

        If Vår_storkelTextBoxKjol.Text = "38/40 - Medium" Then
            storlekskat = "12|16|90||12|16||12"
        End If

        If Vår_storkelTextBoxKjol.Text = "42/44 - Large" Then
            storlekskat = "12|16|91||12|16||12"
        End If

        If Vår_storkelTextBoxKjol.Text = "46/48 - X-large" Then
            storlekskat = "12|16|92||12|16||12"
        End If

        If Vår_storkelTextBoxKjol.Text = "50/52 - 2X-large" Then
            storlekskat = "12|16|93||12|16||12"
        End If

        Try
            Pris_um1TextBoxKjol.Text = Pris_mmTextBoxKjol.Text * 0.8
        Catch ex As Exception
            MsgBox("FYI: Inget pris angivet!")
        End Try

        ActiveradTextBoxKjol.Text = "0"
        ConditionComboBoxKjol.Text = "used"
        CategoryTextBoxKjol.Text = storlekskat
        LagersaldoTextBoxKjol.Text = 1
        SKATTTextBoxKjol.Text = "1"
        MåttbeskrivningTextBoxKjol.Text = "Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag."
    End Sub

    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles SetSizeBtnJacka.Click

        Try

            If BystTextBoxJacka.Text <> Nothing Then

                Dim toppus As Integer = BystTextBoxJacka.Text
                Dim toppgenomsnitt As Integer

                toppgenomsnitt = toppus * 2

                'MsgBox(toppgenomsnitt)

                '30/32 - X-small
                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxJacka.Text = "30/32 - X-small"
                End If

                '34/34 - Small
                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxJacka.Text = "34/36 - Small"
                End If

                '38/40 - Medium
                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
                    Vår_storlekTextBoxJacka.Text = "38/40 - Medium"
                End If

                '42-44 - Large
                If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
                    Vår_storlekTextBoxJacka.Text = "42/44 - Large"
                End If

                '46/48 - X-large
                If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
                    Vår_storlekTextBoxJacka.Text = "46/48 - X-large"
                End If

                '50/52 - 2X-large
                If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxJacka.Text = "50/52 - 2X-large"
                End If

            End If

            ' BYST MED/UTAN STRETCH BEHÖVS NOG INTE FÖR JACKOR/KAPPOR
            'If BystTextBoxJacka.Text = Nothing Then

            '    Dim toppus As Integer
            '    Dim toppms As Integer
            '    Dim toppgenomsnitt As Integer

            '    toppus = Convert.ToInt32(Byst_usTextBoxKlanning.Text)
            '    toppms = Convert.ToInt32(Byst_msTextBoxKlanning.Text)

            '    toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


            '    'MsgBox(toppgenomsnitt)

            '    '30/32 - X-small
            '    If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
            '        Vår_storlekTextBoxKlanning.Text = "30/32 - X-small"
            '    End If

            '    '34/34 - Small
            '    If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
            '        Vår_storlekTextBoxKlanning.Text = "34/36 - Small"
            '    End If

            '    '38/40 - Medium
            '    If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
            '        Vår_storlekTextBoxKlanning.Text = "38/40 - Medium"
            '    End If

            '    '42-44 - Large
            '    If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
            '        Vår_storlekTextBoxKlanning.Text = "42-44 - Large"
            '    End If

            '    '46/48 - X-large
            '    If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
            '        Vår_storlekTextBoxKlanning.Text = "46/48 - X-large"
            '    End If

            '    '50/52 - 2X-large
            '    If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
            '        Vår_storlekTextBoxKlanning.Text = "50/52 - 2X-large"
            '    End If

            'End If

        Catch ex As Exception
            MsgBox("Fälten för byst verkar tomma?")
        End Try





        'GAMMAL
        'Try

        '    If BystTextBoxJacka.Text <> Nothing Then

        '        Dim toppus As Integer = BystTextBoxJacka.Text
        '        Dim toppgenomsnitt As Integer

        '        toppgenomsnitt = toppus * 2

        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 72) Then
        '            Vår_storlekTextBoxJacka.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 73) And (toppgenomsnitt <= 77) Then
        '            Vår_storlekTextBoxJacka.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 78) And (toppgenomsnitt <= 86) Then
        '            Vår_storlekTextBoxJacka.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 88) Then
        '            Vår_storlekTextBoxJacka.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 89) And (toppgenomsnitt <= 97) Then
        '            Vår_storlekTextBoxJacka.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 100) Then
        '            Vår_storlekTextBoxJacka.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 101) And (toppgenomsnitt <= 109) Then
        '            Vår_storlekTextBoxJacka.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 110) And (toppgenomsnitt <= 111) Then
        '            Vår_storlekTextBoxJacka.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 112) And (toppgenomsnitt <= 200) Then
        '            Vår_storlekTextBoxJacka.Text = "Large - XL"
        '        End If

        '    End If

        'Catch ex As Exception
        '    MsgBox("Fälten för byst verkar tomma? Fälten bör innehålla antingen mått på byst med och utan stretch alternativt vanligt mått för icke-stretchplagg")
        'End Try

    End Sub

    Private Sub Button45_Click(sender As Object, e As EventArgs) Handles Button45.Click
        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxJacka.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        BilderListBoxJacka.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                BilderListBoxJacka.Items.Add(foundFile.Substring(foundFile.Length - 12))

            Next
        Catch ex As Exception

        End Try

        If BilderListBoxJacka.Items.Count.ToString = 0 Then
            BilderTextBoxJacka.Text = Nothing
        End If

        If BilderListBoxJacka.Items.Count.ToString = 1 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 2 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 3 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 4 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString & ";" & BilderListBoxJacka.Items(3).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 5 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString & ";" & BilderListBoxJacka.Items(3).ToString & ";" & BilderListBoxJacka.Items(4).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 6 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString & ";" & BilderListBoxJacka.Items(3).ToString & ";" & BilderListBoxJacka.Items(4).ToString & ";" & BilderListBoxJacka.Items(5).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 7 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString & ";" & BilderListBoxJacka.Items(3).ToString & ";" & BilderListBoxJacka.Items(4).ToString & ";" & BilderListBoxJacka.Items(5).ToString & ";" & BilderListBoxJacka.Items(6).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 8 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString & ";" & BilderListBoxJacka.Items(3).ToString & ";" & BilderListBoxJacka.Items(4).ToString & ";" & BilderListBoxJacka.Items(5).ToString & ";" & BilderListBoxJacka.Items(6).ToString & ";" & BilderListBoxJacka.Items(7).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 9 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString & ";" & BilderListBoxJacka.Items(3).ToString & ";" & BilderListBoxJacka.Items(4).ToString & ";" & BilderListBoxJacka.Items(5).ToString & ";" & BilderListBoxJacka.Items(6).ToString & ";" & BilderListBoxJacka.Items(7).ToString & ";" & BilderListBoxJacka.Items(8).ToString
        End If

        If BilderListBoxJacka.Items.Count.ToString = 10 Then
            BilderTextBoxJacka.Text = BilderListBoxJacka.Items(0).ToString & ";" & BilderListBoxJacka.Items(1).ToString & ";" & BilderListBoxJacka.Items(2).ToString & ";" & BilderListBoxJacka.Items(3).ToString & ";" & BilderListBoxJacka.Items(4).ToString & ";" & BilderListBoxJacka.Items(5).ToString & ";" & BilderListBoxJacka.Items(6).ToString & ";" & BilderListBoxJacka.Items(7).ToString & ";" & BilderListBoxJacka.Items(8).ToString & ";" & BilderListBoxJacka.Items(89).ToString
        End If

        BilderTextBoxJacka.Text.Replace("G:\Produktbilder\", Nothing)

        BilderListBoxJacka.Items.Clear()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click

        Panel4.Enabled = True

        Button17.Enabled = False
        Button16.Enabled = True
        Button98.Enabled = True

        Me.Dam_KjolBindingSource.AddNew()

        Dim newitem As Int16
        Dam_KjolBindingSource.MoveLast()
        newitem = ArtikelnrTextBoxKjol.Text.Remove(0, 3) + 1

        If ForvaldhyllaTextBoxKjol.Text IsNot Nothing Then
            HyllaTextBoxKjol.Text = ForvaldhyllaTextBoxKjol.Text
        End If

        If ForvaldlevComboBoxKjol.Text IsNot Nothing Then
            InlagdTextBox2.Text = ForvaldlevComboBoxKjol.Text
        End If

        ArtikelnrTextBoxKjol.Text = "Dkj" & newitem

        ArtikelnrTextBoxKjol.Focus()


    End Sub

    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click

        Panel3.Enabled = True
        Button43.Enabled = False
        Button42.Enabled = True
        Button97.Enabled = True

        Me.Dam_JackaBindingSource.AddNew()

        Dim newitem As Int16
        Dam_JackaBindingSource.MoveLast()
        newitem = ArtikelnrTextBoxJacka.Text.Remove(0, 3) + 1




        If ForvaldhyllaTextBoxJacka.Text IsNot Nothing Then
            HyllaTextBoxJacka.Text = ForvaldhyllaTextBoxJacka.Text
        End If

        If ForvaldlevComboBoxJacka.Text IsNot Nothing Then
            InlagdTextBox1.Text = ForvaldlevComboBoxJacka.Text
        End If

        ArtikelnrTextBoxJacka.Text = "Dja" & newitem

        ArtikelnrTextBoxJacka.Focus()

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        SetSizeBtnKjol.PerformClick()
        CalcPriceBtnKjol.PerformClick()

        'Sätter storleksangivelse i slutet av produktnamn/rubrik
        'Dim tempstring As String
        'Dim old As String = " - " & Vår_storkelTextBoxKjol.Text
        'Dim removeold As String = SV_BeskrivningTextBox3.Text.Replace(old, Nothing)
        'tempstring = removeold
        'SV_BeskrivningTextBox3.Text = removeold
        'tempstring = SV_BeskrivningTextBox3.Text & " - " & Vår_storkelTextBoxKjol.Text
        'SV_BeskrivningTextBox3.Text = tempstring

        Me.Validate()
        Me.Dam_KjolBindingSource.EndEdit()
        Me.Dam_KjolTableAdapter.Update(Me.CLOREGDBDataSet)

        Panel4.Enabled = False
        Button17.Enabled = True
        Button98.Enabled = False
        Button16.Enabled = False
        Button105.Enabled = True

    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        SetSizeBtnJacka.PerformClick()
        CalcPriceBtnJacka.PerformClick()

        'Sätter storleksangivelse i slutet av produktnamn/rubrik
        'Dim tempstring As String
        'Dim old As String = " - " & Vår_storlekTextBoxJacka.Text
        'Dim removeold As String = SV_BeskrivningTextBox5.Text.Replace(old, Nothing)
        'tempstring = removeold
        'SV_BeskrivningTextBox5.Text = removeold
        'tempstring = SV_BeskrivningTextBox5.Text & " - " & Vår_storlekTextBoxJacka.Text
        'SV_BeskrivningTextBox5.Text = tempstring

        Me.Validate()
        Me.Dam_JackaBindingSource.EndEdit()
        Me.Dam_JackaTableAdapter.Update(Me.CLOREGDBDataSet)

        Panel3.Enabled = False
        Button43.Enabled = True
        Button97.Enabled = False
        Button42.Enabled = False
        Button104.Enabled = True

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        For Each dataRow As DataRowView In Dam_KjolListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        MsgBox("Klart! Mappar skappade.")

        'Dim path As String = "G:\Produktbilder\" & ArtikelnrTextBoxKjol.Text

        'Try
        '    ' Determine whether the directory exists.
        '    If Directory.Exists(path) Then
        '        MsgBox("Mapp finns redan")
        '        Exit Sub
        '    End If

        '    ' Try to create the directory.
        '    Dim di As DirectoryInfo = Directory.CreateDirectory(path)
        '    MsgBox("Mapp skapad")


        'Catch a As Exception
        '    Console.WriteLine("The process failed: {0}.", a.ToString())
        'End Try

    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click

        For Each dataRow As DataRowView In Dam_JackaListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        MsgBox("Klart! Mappar skappade.")

        'Dim path As String = "G:\Produktbilder\" & ArtikelnrTextBoxJacka.Text

        'Try
        '    ' Determine whether the directory exists.
        '    If Directory.Exists(path) Then
        '        MsgBox("Mapp finns redan")
        '        Exit Sub
        '    End If

        '    ' Try to create the directory.
        '    Dim di As DirectoryInfo = Directory.CreateDirectory(path)
        '    MsgBox("Mapp skapad")


        'Catch a As Exception
        '    Console.WriteLine("The process failed: {0}.", a.ToString())
        'End Try

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dam_KjolBindingSource.MoveFirst()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dam_KjolBindingSource.MovePrevious()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dam_KjolBindingSource.MoveNext()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dam_KjolBindingSource.MoveLast()
    End Sub

    Private Sub Button44_Click_1(sender As Object, e As EventArgs) Handles Button44.Click

        For Each dataRow As DataRowView In Dam_ByxorListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        MsgBox("Klart! Mappar skappade.")

        'Dim path As String = "G:\Produktbilder\" & ArtikelnrTextBoxByxor.Text

        'Try
        '    ' Determine whether the directory exists.
        '    If Directory.Exists(path) Then
        '        MsgBox("Mapp finns redan")
        '        Exit Sub
        '    End If

        '    ' Try to create the directory.
        '    Dim di As DirectoryInfo = Directory.CreateDirectory(path)
        '    MsgBox("Mapp skapad")


        'Catch a As Exception
        '    Console.WriteLine("The process failed: {0}.", a.ToString())
        'End Try

    End Sub

    Private Sub Button53_Click(sender As Object, e As EventArgs) Handles Button53.Click
        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxByxor.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        BilderListboxByxor.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                BilderListboxByxor.Items.Add(foundFile.Substring(foundFile.Length - 12))
                'File.Copy(foundFile, "G:\produktbilder")

            Next
        Catch ex As Exception

        End Try

        If BilderListboxByxor.Items.Count.ToString = 0 Then
            BilderTextBoxByxor.Text = Nothing
        End If

        If BilderListboxByxor.Items.Count.ToString = 1 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 2 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 3 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 4 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString & ";" & BilderListboxByxor.Items(3).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 5 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString & ";" & BilderListboxByxor.Items(3).ToString & ";" & BilderListboxByxor.Items(4).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 6 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString & ";" & BilderListboxByxor.Items(3).ToString & ";" & BilderListboxByxor.Items(4).ToString & ";" & BilderListboxByxor.Items(5).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 7 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString & ";" & BilderListboxByxor.Items(3).ToString & ";" & BilderListboxByxor.Items(4).ToString & ";" & BilderListboxByxor.Items(5).ToString & ";" & BilderListboxByxor.Items(6).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 8 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString & ";" & BilderListboxByxor.Items(3).ToString & ";" & BilderListboxByxor.Items(4).ToString & ";" & BilderListboxByxor.Items(5).ToString & ";" & BilderListboxByxor.Items(6).ToString & ";" & BilderListboxByxor.Items(7).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 9 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString & ";" & BilderListboxByxor.Items(3).ToString & ";" & BilderListboxByxor.Items(4).ToString & ";" & BilderListboxByxor.Items(5).ToString & ";" & BilderListboxByxor.Items(6).ToString & ";" & BilderListboxByxor.Items(7).ToString & ";" & BilderListboxByxor.Items(8).ToString
        End If

        If BilderListboxByxor.Items.Count.ToString = 10 Then
            BilderTextBoxByxor.Text = BilderListboxByxor.Items(0).ToString & ";" & BilderListboxByxor.Items(1).ToString & ";" & BilderListboxByxor.Items(2).ToString & ";" & BilderListboxByxor.Items(3).ToString & ";" & BilderListboxByxor.Items(4).ToString & ";" & BilderListboxByxor.Items(5).ToString & ";" & BilderListboxByxor.Items(6).ToString & ";" & BilderListboxByxor.Items(7).ToString & ";" & BilderListboxByxor.Items(8).ToString & ";" & BilderListboxByxor.Items(89).ToString
        End If

        BilderTextBoxByxor.Text.Replace("G:\Produktbilder\", Nothing)

        BilderListboxByxor.Items.Clear()
    End Sub

    Private Sub Button54_Click(sender As Object, e As EventArgs) Handles Button54.Click

        For Each dataRow As DataRowView In AccessoarerListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        For Each dataRow As DataRowView In Dam_ByxorListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        For Each dataRow As DataRowView In Dam_JackaListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        For Each dataRow As DataRowView In Dam_KjolListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        For Each dataRow As DataRowView In Dam_KlänningListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        For Each dataRow As DataRowView In Dam_ToppListBoxToppar.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next


    End Sub

    Private Sub Button55_Click(sender As Object, e As EventArgs) Handles AvbrytAccButton.Click

        Me.AccessoarerTableAdapter.Fill(Me.CLOREGDBDataSet.Accessoarer)
        Panel1.Enabled = False
        Button33.Enabled = False
        Button102.Enabled = True
        AvbrytAccButton.Enabled = False

    End Sub

    Private Sub Button62_Click(sender As Object, e As EventArgs) Handles Button62.Click

        Panel7.Enabled = True

        Button63.Enabled = True
        Button101.Enabled = True
        Button62.Enabled = False

        Me.Dam_JumpsuitBindingSource.AddNew()

        Dim newitem As Int16
        Dam_JumpsuitBindingSource.MoveLast()
        newitem = ArtikelnrTextBoxJumpsuit.Text.Remove(0, 3) + 1

        If ForvaldhyllaTextBoxJumpsuit.Text IsNot Nothing Then
            HyllaTextBoxJumpsuit.Text = ForvaldhyllaTextBoxJumpsuit.Text
        End If

        If ForvaldlevComboBoxJumpsuit.Text IsNot Nothing Then
            InlagdTextBox5.Text = ForvaldlevComboBoxJumpsuit.Text
        End If

        ArtikelnrTextBoxJumpsuit.Text = "Djs" & newitem

        ArtikelnrTextBoxJumpsuit.Focus()

    End Sub

    Private Sub Button61_Click(sender As Object, e As EventArgs) Handles Button63.Click
        SetSizeBtnJumpsuit.PerformClick()
        CalcPriceBtnJumpsuit.PerformClick()

        'Sätter storleksangivelse i slutet av produktnamn/rubrik
        'Dim tempstring As String
        'Dim old As String = " - " & Vår_storlekTextBoxJumpsuit.Text
        'Dim removeold As String = SV_BeskrivningTextBoxJumpsuit.Text.Replace(old, Nothing)
        'tempstring = removeold
        'SV_BeskrivningTextBoxJumpsuit.Text = removeold
        'tempstring = SV_BeskrivningTextBoxJumpsuit.Text & " - " & Vår_storlekTextBoxJumpsuit.Text
        'SV_BeskrivningTextBoxJumpsuit.Text = tempstring

        Me.Validate()
        Me.Dam_JumpsuitBindingSource.EndEdit()
        Me.Dam_JumpsuitTableAdapter.Update(Me.CLOREGDBDataSet)

        Panel7.Enabled = False

        Button63.Enabled = False
        Button101.Enabled = False
        Button62.Enabled = True
        Button108.Enabled = True

    End Sub

    Private Sub Button57_Click(sender As Object, e As EventArgs) Handles Button57.Click
        Dam_JumpsuitBindingSource.MoveNext()
    End Sub

    Private Sub Button56_Click(sender As Object, e As EventArgs) Handles Button56.Click
        Dam_JumpsuitBindingSource.MoveLast()
    End Sub

    Private Sub Button58_Click(sender As Object, e As EventArgs) Handles Button58.Click
        Dam_JumpsuitBindingSource.MovePrevious()
    End Sub

    Private Sub Button59_Click(sender As Object, e As EventArgs) Handles Button59.Click
        Dam_JumpsuitBindingSource.MoveFirst()
    End Sub

    Private Sub Button63_Click(sender As Object, e As EventArgs) Handles SetSizeBtnJumpsuit.Click

        Try

            If BystTextBoxJumpsuit.Text <> Nothing Then

                Dim toppus As Integer = BystTextBoxJumpsuit.Text
                Dim toppgenomsnitt As Integer

                toppgenomsnitt = toppus * 2

                'MsgBox(toppgenomsnitt)

                '30/32 - X-small
                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxJumpsuit.Text = "30/32 - X-small"
                End If

                '34/34 - Small
                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxJumpsuit.Text = "34/36 - Small"
                End If

                '38/40 - Medium
                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
                    Vår_storlekTextBoxJumpsuit.Text = "38/40 - Medium"
                End If

                '42-44 - Large
                If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
                    Vår_storlekTextBoxJumpsuit.Text = "42/44 - Large"
                End If

                '46/48 - X-large
                If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
                    Vår_storlekTextBoxJumpsuit.Text = "46/48 - X-large"
                End If

                '50/52 - 2X-large
                If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxJumpsuit.Text = "50/52 - 2X-large"
                End If

            End If


            If BystTextBoxJumpsuit.Text = Nothing Then

                Dim toppus As Integer
                Dim toppms As Integer
                Dim toppgenomsnitt As Integer

                toppus = Convert.ToInt32(Byst_usTextBoxJumpsuit.Text)
                toppms = Convert.ToInt32(Byst_msTextBoxJumpsuit.Text)

                toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


                'MsgBox(toppgenomsnitt)

                '30/32 - X-small
                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxJumpsuit.Text = "30/32 - X-small"
                End If

                '34/34 - Small
                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxJumpsuit.Text = "34/36 - Small"
                End If

                '38/40 - Medium
                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 94) Then
                    Vår_storlekTextBoxJumpsuit.Text = "38/40 - Medium"
                End If

                '42-44 - Large
                If (toppgenomsnitt >= 95) And (toppgenomsnitt <= 102) Then
                    Vår_storlekTextBoxJumpsuit.Text = "42/44 - Large"
                End If

                '46/48 - X-large
                If (toppgenomsnitt >= 103) And (toppgenomsnitt <= 113) Then
                    Vår_storlekTextBoxJumpsuit.Text = "46/48 - X-large"
                End If

                '50/52 - 2X-large
                If (toppgenomsnitt >= 114) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxJumpsuit.Text = "50/52 - 2X-large"
                End If

            End If

        Catch ex As Exception
            MsgBox("Fälten för byst verkar tomma? Fälten bör innehålla antingen mått på byst med och utan stretch alternativt vanligt mått för icke-stretchplagg")
        End Try


        'Gammal storleksuppskattning
        'Try

        '    If BystTextBoxJumpsuit.Text <> Nothing Then

        '        Dim toppus As Integer = BystTextBoxJumpsuit.Text
        '        Dim toppgenomsnitt As Integer

        '        toppgenomsnitt = toppus * 2

        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 72) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 73) And (toppgenomsnitt <= 77) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 78) And (toppgenomsnitt <= 86) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 88) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 89) And (toppgenomsnitt <= 97) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 100) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 101) And (toppgenomsnitt <= 109) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 110) And (toppgenomsnitt <= 111) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 112) And (toppgenomsnitt <= 200) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Large - XL"
        '        End If

        '    End If


        '    If BystTextBoxJumpsuit.Text = Nothing Then

        '        Dim toppus As Integer
        '        Dim toppms As Integer
        '        Dim toppgenomsnitt As Integer

        '        toppus = Convert.ToInt32(Byst_usTextBoxJumpsuit.Text)
        '        toppms = Convert.ToInt32(Byst_msTextBoxJumpsuit.Text)

        '        toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


        '        'MsgBox(toppgenomsnitt)

        '        'XS
        '        If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 72) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Small - XS"
        '        End If

        '        'XS-S
        '        If (toppgenomsnitt >= 73) And (toppgenomsnitt <= 77) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Small - XS/S"
        '        End If

        '        'S
        '        If (toppgenomsnitt >= 78) And (toppgenomsnitt <= 86) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Small - S"
        '        End If

        '        'S/M
        '        If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 88) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Medium - S/M"
        '        End If

        '        'M
        '        If (toppgenomsnitt >= 89) And (toppgenomsnitt <= 97) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Medium - M"
        '        End If

        '        'M/L
        '        If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 100) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Medium - M/L"
        '        End If

        '        'L
        '        If (toppgenomsnitt >= 101) And (toppgenomsnitt <= 109) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Large - L"
        '        End If

        '        'L/XL
        '        If (toppgenomsnitt >= 110) And (toppgenomsnitt <= 111) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Large - L/XL"
        '        End If

        '        'XL
        '        If (toppgenomsnitt >= 112) And (toppgenomsnitt <= 200) Then
        '            Vår_storlekTextBoxJumpsuit.Text = "Large - XL"
        '        End If

        '    End If

        'Catch ex As Exception
        '    MsgBox("Fälten för byst verkar tomma? Fälten bör innehålla antingen mått på byst med och utan stretch alternativt vanligt mått för icke-stretchplagg")
        'End Try



    End Sub

    Private Sub Button61_Click_1(sender As Object, e As EventArgs) Handles CalcPriceBtnJumpsuit.Click
        'Private Sub CalcPriceBtnJumpsuit_Click_1(sender As Object, e As EventArgs) Handles CalcPriceBtnJumpsuit.Click

        Dim storlekskat As String



        If Vår_storlekTextBoxJumpsuit.Text = "30/32 - X-small" Then
            storlekskat = "12|20|82||12|20||12"
        End If

        If Vår_storlekTextBoxJumpsuit.Text = "34/36 - Small" Then
            storlekskat = "12|20|83||12|20||12"

        End If

        If Vår_storlekTextBoxJumpsuit.Text = "38/40 - Medium" Then
            storlekskat = "12|20|84||12|20||12"
        End If

        If Vår_storlekTextBoxJumpsuit.Text = "42/44 - Large" Then
            storlekskat = "12|20|85||12|20||12"
        End If

        If Vår_storlekTextBoxJumpsuit.Text = "46/48 - X-large" Then
            storlekskat = "12|20|86||12|20||12"
        End If

        If Vår_storlekTextBoxJumpsuit.Text = "50/52 - 2X-large" Then
            storlekskat = "12|20|87||12|20||12"
        End If


        Try
            Pris_umTextBoxJumpsuit.Text = Pris_mmTextBoxJumpsuit.Text * 0.8
        Catch ex As Exception
            MsgBox("FYI: Inget pris angivet!")
        End Try

        ActiveradTextBoxJumpsuit.Text = "0"
        ConditionComboBoxJumpsuit.Text = "used"
        CategoryTextBoxJumpsuit.Text = storlekskat
        LagersaldoTextBoxJumpsuit.Text = "1"
        SKATTTextBoxJumpsuit.Text = "1"
        MåttbeskrivningTextBoxJumpsuit.Text = "Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag."
    End Sub

    Private Sub Button18_Click_1(sender As Object, e As EventArgs) Handles Button18.Click

        For Each dataRow As DataRowView In Dam_JumpsuitListBox.Items

            Dim value As String = dataRow.Item("Artikelnr")

            Dim path As String = "G:\Produktbilder\" & value

            Try

                Dim di As DirectoryInfo = Directory.CreateDirectory(path)

            Catch a As Exception
                Console.WriteLine("The process failed: {0}.", a.ToString())
            End Try

        Next

        MsgBox("Klart! Mappar skappade.")

        'Dim path As String = "G:\Produktbilder\" & ArtikelnrTextBoxJumpsuit.Text

        'Try
        '    ' Determine whether the directory exists.
        '    If Directory.Exists(path) Then
        '        MsgBox("Mapp finns redan")
        '        Exit Sub
        '    End If

        '    ' Try to create the directory.
        '    Dim di As DirectoryInfo = Directory.CreateDirectory(path)
        '    MsgBox("Mapp skapad")


        'Catch a As Exception
        '    Console.WriteLine("The process failed: {0}.", a.ToString())
        'End Try

    End Sub

    Private Sub Button61_Click_2(sender As Object, e As EventArgs) Handles Button61.Click

        Dim myString As String = "G:\Produktbilder\" & ArtikelnrTextBoxJumpsuit.Text

        'Me.WebBrowser2.Navigate(New Uri(myString))

        BilderListBoxJumpsuit.Items.Clear()

        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
          myString)

                BilderListBoxJumpsuit.Items.Add(foundFile.Substring(foundFile.Length - 12))

            Next
        Catch ex As Exception

        End Try

        If BilderListBoxJumpsuit.Items.Count.ToString = 0 Then
            BilderTextBoxJumpsuit.Text = Nothing
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 1 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 2 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 3 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 4 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString & ";" & BilderListBoxJumpsuit.Items(3).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 5 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString & ";" & BilderListBoxJumpsuit.Items(3).ToString & ";" & BilderListBoxJumpsuit.Items(4).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 6 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString & ";" & BilderListBoxJumpsuit.Items(3).ToString & ";" & BilderListBoxJumpsuit.Items(4).ToString & ";" & BilderListBoxJumpsuit.Items(5).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 7 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString & ";" & BilderListBoxJumpsuit.Items(3).ToString & ";" & BilderListBoxJumpsuit.Items(4).ToString & ";" & BilderListBoxJumpsuit.Items(5).ToString & ";" & BilderListBoxJumpsuit.Items(6).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 8 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString & ";" & BilderListBoxJumpsuit.Items(3).ToString & ";" & BilderListBoxJumpsuit.Items(4).ToString & ";" & BilderListBoxJumpsuit.Items(5).ToString & ";" & BilderListBoxJumpsuit.Items(6).ToString & ";" & BilderListBoxJumpsuit.Items(7).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 9 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString & ";" & BilderListBoxJumpsuit.Items(3).ToString & ";" & BilderListBoxJumpsuit.Items(4).ToString & ";" & BilderListBoxJumpsuit.Items(5).ToString & ";" & BilderListBoxJumpsuit.Items(6).ToString & ";" & BilderListBoxJumpsuit.Items(7).ToString & ";" & BilderListBoxJumpsuit.Items(8).ToString
        End If

        If BilderListBoxJumpsuit.Items.Count.ToString = 10 Then
            BilderTextBoxJumpsuit.Text = BilderListBoxJumpsuit.Items(0).ToString & ";" & BilderListBoxJumpsuit.Items(1).ToString & ";" & BilderListBoxJumpsuit.Items(2).ToString & ";" & BilderListBoxJumpsuit.Items(3).ToString & ";" & BilderListBoxJumpsuit.Items(4).ToString & ";" & BilderListBoxJumpsuit.Items(5).ToString & ";" & BilderListBoxJumpsuit.Items(6).ToString & ";" & BilderListBoxJumpsuit.Items(7).ToString & ";" & BilderListBoxJumpsuit.Items(8).ToString & ";" & BilderListBoxJumpsuit.Items(89).ToString
        End If

        BilderTextBoxJumpsuit.Text.Replace("G:\Produktbilder\", Nothing)

        BilderListBoxJumpsuit.Items.Clear()

    End Sub

    Private Sub Button64_Click(sender As Object, e As EventArgs) Handles Button64.Click

        For Each dataRow As DataRowView In Dam_JackaListBox.Items
            SetSizeBtnJacka.PerformClick()
        Next

    End Sub

    Private Sub Button65_Click(sender As Object, e As EventArgs) Handles Button65.Click
        CalcPriceBtnKjol.PerformClick()
        'SetSizeBtnKjol.PerformClick()
        Button12.PerformClick()

    End Sub

    Private Sub Button66_Click(sender As Object, e As EventArgs) Handles Button66.Click


        CalcPriceBtnKlanning.PerformClick()
        Button20.PerformClick()


    End Sub

    Private Sub Button67_Click(sender As Object, e As EventArgs) Handles Button67.Click

        SetSizeBtnTopp.PerformClick()
        'CalcPriceBtnTopp.PerformClick()
        Button7.PerformClick()

        'For Each dataRow As DataRowView In Dam_ToppListBoxToppar.Items

        '    SetSizeBtnTopp.PerformClick()
        '    'CalcPriceBtnTopp.PerformClick()

        'Next

    End Sub

    Private Sub Button68_Click(sender As Object, e As EventArgs) Handles Button68.Click


        SetSizeBtnJumpsuit.PerformClick()
        Button57.PerformClick()


    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub CalcSizeBtnByxor_Click(sender As Object, e As EventArgs) Handles CalcSizeBtnByxor.Click

        Try

            If MidjemåttTextBoxByxor.Text <> Nothing Then

                Dim toppus As Integer = MidjemåttTextBoxByxor.Text
                Dim toppgenomsnitt As Integer

                toppgenomsnitt = toppus * 2

                'MsgBox(toppgenomsnitt)


                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 60) Then
                    Vår_storlekTextBoxByxor.Text = "30/32 - X-small"
                End If


                If (toppgenomsnitt >= 61) And (toppgenomsnitt <= 70) Then
                    Vår_storlekTextBoxByxor.Text = "34/36 - Small"
                End If


                If (toppgenomsnitt >= 71) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxByxor.Text = "38/40 - Medium"
                End If


                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxByxor.Text = "42/44 - Large"
                End If


                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 97) Then
                    Vår_storlekTextBoxByxor.Text = "46/48 - X-large"
                End If


                If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxByxor.Text = "50/52 - 2X-large"
                End If

            End If


            If MidjemåttTextBoxByxor.Text = Nothing Then

                Dim toppus As Integer
                Dim toppms As Integer
                Dim toppgenomsnitt As Integer

                toppus = Convert.ToInt32(MidjausTextboxByxor.Text)
                toppms = Convert.ToInt32(MidjamsTextboxByxor.Text)

                toppgenomsnitt = ((toppms - toppus) / 2 + toppus) * 2


                'MsgBox(toppgenomsnitt)


                If (toppgenomsnitt >= 0) And (toppgenomsnitt <= 60) Then
                    Vår_storlekTextBoxByxor.Text = "30/32 - X-small"
                End If


                If (toppgenomsnitt >= 61) And (toppgenomsnitt <= 70) Then
                    Vår_storlekTextBoxByxor.Text = "34/36 - Small"
                End If


                If (toppgenomsnitt >= 71) And (toppgenomsnitt <= 78) Then
                    Vår_storlekTextBoxByxor.Text = "38/40 - Medium"
                End If


                If (toppgenomsnitt >= 79) And (toppgenomsnitt <= 86) Then
                    Vår_storlekTextBoxByxor.Text = "42/44 - Large"
                End If


                If (toppgenomsnitt >= 87) And (toppgenomsnitt <= 97) Then
                    Vår_storlekTextBoxByxor.Text = "46/48 - X-large"
                End If


                If (toppgenomsnitt >= 98) And (toppgenomsnitt <= 200) Then
                    Vår_storlekTextBoxByxor.Text = "50/52 - 2X-large"
                End If

            End If

        Catch ex As Exception
            MsgBox("Fälten för midja verkar tomma? Fälten bör innehålla antingen mått på midja med och utan stretch alternativt vanligt midjemått för icke-stretchplagg")
        End Try

    End Sub

    Private Sub Button69_Click(sender As Object, e As EventArgs) Handles Button69.Click
        CalcPriceBtnByxor.PerformClick()
        'CalcSizeBtnByxor.PerformClick()
        Button47.PerformClick()

    End Sub

    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Button47.Click
        Dam_ByxorBindingSource.MoveNext()
    End Sub

    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click

        If PlaggTextBoxByxor.Text = Nothing Then
            MsgBox("Typ av byxa måste anges!")
            Exit Sub
        End If

        CalcSizeBtnByxor.PerformClick()
        CalcPriceBtnByxor.PerformClick()

        'Sätter storleksangivelse i slutet av produktnamn/rubrik
        'Dim tempstring As String
        'Dim old As String = " - " & Vår_storlekTextBoxByxor.Text
        'Dim removeold As String = SV_BeskrivningTextBox4.Text.Replace(old, Nothing)
        'tempstring = removeold
        'SV_BeskrivningTextBox4.Text = removeold
        'tempstring = SV_BeskrivningTextBox4.Text & " - " & Vår_storlekTextBoxByxor.Text
        'SV_BeskrivningTextBox4.Text = tempstring

        Me.Validate()
        Me.Dam_ByxorBindingSource.EndEdit()
        Me.Dam_ByxorTableAdapter.Update(Me.CLOREGDBDataSet)

        Panel2.Enabled = False
        Button55.Enabled = False
        Button52.Enabled = True
        Button103.Enabled = True


    End Sub

    Private Sub CalcPriceBtnByxor_Click(sender As Object, e As EventArgs) Handles CalcPriceBtnByxor.Click

        Dim storlekskat As String

        If PlaggTextBoxByxor.Text = "Byxor" Then

            If Vår_storlekTextBoxByxor.Text = "30/32 - X-small" Then
                storlekskat = "12|19|45|100||12|19|45||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "34/36 - Small" Then
                storlekskat = "12|19|45|101||12|19|45||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "38/40 - Medium" Then
                storlekskat = "12|19|45|102||12|19|45||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "42/44 - Large" Then
                storlekskat = "12|19|45|103||12|19|45||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "46/48 - X-large" Then
                storlekskat = "12|19|45|104||12|19|45||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "50/52 - 2X-large" Then
                storlekskat = "12|19|45|105||12|19|45||12|19||12"
            End If

        End If



        If PlaggTextBoxByxor.Text = "Leggings" Then

            If Vår_storlekTextBoxByxor.Text = "30/32 - X-small" Then
                storlekskat = "12|19|46|106||12|19|46||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "34/36 - Small" Then
                storlekskat = "12|19|46|107||12|19|46||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "38/40 - Medium" Then
                storlekskat = "12|19|46|108||12|19|46||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "42/44 - Large" Then
                storlekskat = "12|19|46|109||12|19|46||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "46/48 - X-large" Then
                storlekskat = "12|19|46|110||12|19|46||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "50/52 - 2X-large" Then
                storlekskat = "12|19|46|111||12|19|46||12|19||12"
            End If

        End If



        If PlaggTextBoxByxor.Text = "Shorts" Then

            If Vår_storlekTextBoxByxor.Text = "30/32 - X-small" Then
                storlekskat = "12|19|121|122||12|19|121||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "34/36 - Small" Then
                storlekskat = "12|19|121|123||12|19|121||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "38/40 - Medium" Then
                storlekskat = "12|19|121|124||12|19|121||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "42/44 - Large" Then
                storlekskat = "12|19|121|125||12|19|121||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "46/48 - X-large" Then
                storlekskat = "12|19|121|126||12|19|121||12|19||12"
            End If

            If Vår_storlekTextBoxByxor.Text = "50/52 - 2X-large" Then
                storlekskat = "12|19|121|127||12|19|121||12|19||12"
            End If

        End If



        Try
            Pris_umTextBoxByxor.Text = Pris_mmTextBoxByxor.Text * 0.8
        Catch ex As Exception
            MsgBox("FYI: Inget pris angivet!")
        End Try

        ActiveradTextBoxByxor.Text = "0"
        ConditionComboBoxByxor.Text = "used"
        CategoryTextBoxByxor.Text = storlekskat
        LagersaldoTextBoxByxor.Text = 1
        SKATTTextBoxByxor.Text = "1"
        MåttbeskrivningTextBoxByxor.Text = "Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag."

    End Sub

    Private Sub Button52_Click(sender As Object, e As EventArgs) Handles Button52.Click

        If ComboBox2.Text = Nothing Then
            MsgBox("Sorry, du måste ange typ (shorts, byxor, leggings)")
            Exit Sub
        End If

        Panel2.Enabled = True
        Button55.Enabled = True
        Button51.Enabled = True
        Button52.Enabled = False

        Dim newitem As Int16
        Dam_ByxorBindingSource.MoveLast()
        newitem = ArtikelnrTextBoxByxor.Text.Remove(0, 3) + 1

        Me.Dam_ByxorBindingSource.AddNew()


        PlaggTextBoxByxor.Text = ComboBox2.Text

        If ForvaldhyllaTextBoxByxor.Text IsNot Nothing Then
            HyllaTextBoxByxor.Text = ForvaldhyllaTextBoxByxor.Text
        End If

        If ForvaldlevComboBox.Text IsNot Nothing Then
            InlagdTextBox.Text = ForvaldlevComboBox.Text
        End If

        ArtikelnrTextBoxByxor.Text = "Dby" & newitem

        ArtikelnrTextBoxByxor.Focus()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub AccessoarerListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AccessoarerListBox.SelectedIndexChanged

        Try

            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxAcc.Text.ToString & "\"

            If Directory.Exists(dir) Then
                Dim fileEntries As String() = Directory.GetFiles(dir)
                If fileEntries.Length = 0 Then
                    PictureBox1.ImageLocation = Nothing
                Else
                    PictureBox1.ImageLocation = fileEntries.First
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Dam_ByxorListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Dam_ByxorListBox.SelectedIndexChanged

        Try

            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxByxor.Text.ToString & "\"

            If Directory.Exists(dir) Then
                Dim fileEntries As String() = Directory.GetFiles(dir)
                If fileEntries.Length = 0 Then
                    PictureBox2.ImageLocation = Nothing
                Else
                    PictureBox2.ImageLocation = fileEntries.First
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Dam_JackaListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Dam_JackaListBox.SelectedIndexChanged

        Try

            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxJacka.Text.ToString & "\"

            If Directory.Exists(dir) Then
                Dim fileEntries As String() = Directory.GetFiles(dir)
                If fileEntries.Length = 0 Then
                    PictureBox3.ImageLocation = Nothing
                Else
                    PictureBox3.ImageLocation = fileEntries.First
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Dam_KjolListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Dam_KjolListBox.SelectedIndexChanged

        Try

            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxKjol.Text.ToString & "\"

            If Directory.Exists(dir) Then
                Dim fileEntries As String() = Directory.GetFiles(dir)
                If fileEntries.Length = 0 Then
                    PictureBox4.ImageLocation = Nothing
                Else
                    PictureBox4.ImageLocation = fileEntries.First
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Dam_KlänningListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Dam_KlänningListBox.SelectedIndexChanged

        Try

            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxKlanning.Text.ToString & "\"

            If Directory.Exists(dir) Then
                Dim fileEntries As String() = Directory.GetFiles(dir)
                If fileEntries.Length = 0 Then
                    PictureBox5.ImageLocation = Nothing
                Else
                    PictureBox5.ImageLocation = fileEntries.First
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Dam_ToppListBoxToppar_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Dam_ToppListBoxToppar.SelectedIndexChanged
        Try

            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text.ToString & "\"

            If Directory.Exists(dir) Then
                Dim fileEntries As String() = Directory.GetFiles(dir)
                If fileEntries.Length = 0 Then
                    PictureBox6.ImageLocation = Nothing
                Else
                    PictureBox6.ImageLocation = fileEntries.First
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Dam_JumpsuitListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Dam_JumpsuitListBox.SelectedIndexChanged

        Try

            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxJumpsuit.Text.ToString & "\"

            If Directory.Exists(dir) Then
                Dim fileEntries As String() = Directory.GetFiles(dir)
                If fileEntries.Length = 0 Then
                    PictureBox7.ImageLocation = Nothing
                Else
                    PictureBox7.ImageLocation = fileEntries.First
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub SetImagesAcc_Click(sender As Object, e As EventArgs) Handles SetImagesAcc.Click

        Dim mystream As Stream = Nothing
        Dim openfiledialog1 As New OpenFileDialog()
        openfiledialog1.Multiselect = True
        openfiledialog1.InitialDirectory = "C:\temp"
        'openfiledialog1.Filter = "JPG images (*.jpg)"
        If openfiledialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                mystream = openfiledialog1.OpenFile()
                If (mystream IsNot Nothing) Then
                    ' fundera ut hur programmet ska skapa katalog och sen kopiera vald fil till katalogen

                    Try
                        Cursor.Current = Cursors.WaitCursor
                        Dim Path As String = "G:\Produktbilder\" & ArtikelnrTextBoxAcc.Text.ToString & "\"

                        If Directory.Exists(Path) Then

                            'radera befintliga filer i mappen
                            For Each deleteFile In Directory.GetFiles(Path, "*.*", SearchOption.TopDirectoryOnly)
                                File.Delete(deleteFile)
                            Next

                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)

                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxAcc.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox1.ImageLocation = Nothing
                                Else
                                    PictureBox1.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button35.PerformClick()
                            Cursor.Current = Cursors.Default

                        Else
                            Cursor.Current = Cursors.WaitCursor
                            Directory.CreateDirectory(Path & "ArtikelnrTextBoxAcc.text")
                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)
                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxAcc.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox1.ImageLocation = Nothing
                                Else
                                    PictureBox1.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button35.PerformClick()
                            Cursor.Current = Cursors.Default

                        End If



                    Catch ex As Exception

                    End Try

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (mystream IsNot Nothing) Then
                    mystream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub SetImagesByxor_Click(sender As Object, e As EventArgs) Handles SetImagesByxor.Click

        Dim mystream As Stream = Nothing
        Dim openfiledialog1 As New OpenFileDialog()
        openfiledialog1.Multiselect = True
        openfiledialog1.InitialDirectory = "C:\temp"
        'openfiledialog1.Filter = "JPG images (*.jpg)"
        If openfiledialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                mystream = openfiledialog1.OpenFile()
                If (mystream IsNot Nothing) Then
                    ' fundera ut hur programmet ska skapa katalog och sen kopiera vald fil till katalogen

                    Try
                        Cursor.Current = Cursors.WaitCursor
                        Dim Path As String = "G:\Produktbilder\" & ArtikelnrTextBoxByxor.Text.ToString & "\"

                        If Directory.Exists(Path) Then

                            'radera befintliga filer i mappen
                            For Each deleteFile In Directory.GetFiles(Path, "*.*", SearchOption.TopDirectoryOnly)
                                File.Delete(deleteFile)
                            Next

                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)

                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxByxor.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox2.ImageLocation = Nothing
                                Else
                                    PictureBox2.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button53.PerformClick()
                            Cursor.Current = Cursors.Default

                        Else
                            Cursor.Current = Cursors.WaitCursor
                            Directory.CreateDirectory(Path & ArtikelnrTextBoxByxor.Text)
                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)
                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxByxor.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox2.ImageLocation = Nothing
                                Else
                                    PictureBox2.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button53.PerformClick()
                            Cursor.Current = Cursors.Default

                        End If



                    Catch ex As Exception

                    End Try

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (mystream IsNot Nothing) Then
                    mystream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub SetImagesJackor_Click(sender As Object, e As EventArgs) Handles SetImagesJackor.Click

        Dim mystream As Stream = Nothing
        Dim openfiledialog1 As New OpenFileDialog()
        openfiledialog1.Multiselect = True
        openfiledialog1.InitialDirectory = "C:\temp"
        'openfiledialog1.Filter = "JPG images (*.jpg)"
        If openfiledialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                mystream = openfiledialog1.OpenFile()
                If (mystream IsNot Nothing) Then
                    ' fundera ut hur programmet ska skapa katalog och sen kopiera vald fil till katalogen

                    Try
                        Cursor.Current = Cursors.WaitCursor
                        Dim Path As String = "G:\Produktbilder\" & ArtikelnrTextBoxJacka.Text.ToString & "\"

                        If Directory.Exists(Path) Then

                            'radera befintliga filer i mappen
                            For Each deleteFile In Directory.GetFiles(Path, "*.*", SearchOption.TopDirectoryOnly)
                                File.Delete(deleteFile)
                            Next

                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)

                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxJacka.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox3.ImageLocation = Nothing
                                Else
                                    PictureBox3.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button45.PerformClick()
                            Cursor.Current = Cursors.Default

                        Else
                            Cursor.Current = Cursors.WaitCursor
                            Directory.CreateDirectory(Path & ArtikelnrTextBoxJacka.Text)
                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)
                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxJacka.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox3.ImageLocation = Nothing
                                Else
                                    PictureBox3.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button45.PerformClick()
                            Cursor.Current = Cursors.Default

                        End If



                    Catch ex As Exception

                    End Try

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (mystream IsNot Nothing) Then
                    mystream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub SetImagesKjol_Click(sender As Object, e As EventArgs) Handles SetImagesKjol.Click

        Dim mystream As Stream = Nothing
        Dim openfiledialog1 As New OpenFileDialog()
        openfiledialog1.Multiselect = True
        openfiledialog1.InitialDirectory = "C:\temp"
        'openfiledialog1.Filter = "JPG images (*.jpg)"
        If openfiledialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                mystream = openfiledialog1.OpenFile()
                If (mystream IsNot Nothing) Then
                    ' fundera ut hur programmet ska skapa katalog och sen kopiera vald fil till katalogen

                    Try
                        Cursor.Current = Cursors.WaitCursor
                        Dim Path As String = "G:\Produktbilder\" & ArtikelnrTextBoxKjol.Text.ToString & "\"

                        If Directory.Exists(Path) Then

                            'radera befintliga filer i mappen
                            For Each deleteFile In Directory.GetFiles(Path, "*.*", SearchOption.TopDirectoryOnly)
                                File.Delete(deleteFile)
                            Next

                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)

                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxKjol.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox4.ImageLocation = Nothing
                                Else
                                    PictureBox4.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            WriteImagesKjolBtn.PerformClick()
                            Cursor.Current = Cursors.Default

                        Else
                            Cursor.Current = Cursors.WaitCursor
                            Directory.CreateDirectory(Path & ArtikelnrTextBoxKjol.Text)
                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)
                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxKjol.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox4.ImageLocation = Nothing
                                Else
                                    PictureBox4.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            WriteImagesKjolBtn.PerformClick()
                            Cursor.Current = Cursors.Default

                        End If



                    Catch ex As Exception

                    End Try

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (mystream IsNot Nothing) Then
                    mystream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub SetImagesKlanning_Click(sender As Object, e As EventArgs) Handles SetImagesKlanning.Click

        Dim mystream As Stream = Nothing
        Dim openfiledialog1 As New OpenFileDialog()
        openfiledialog1.Multiselect = True
        openfiledialog1.InitialDirectory = "C:\temp"
        'openfiledialog1.Filter = "JPG images (*.jpg)"
        If openfiledialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                mystream = openfiledialog1.OpenFile()
                If (mystream IsNot Nothing) Then
                    ' fundera ut hur programmet ska skapa katalog och sen kopiera vald fil till katalogen

                    Try
                        Cursor.Current = Cursors.WaitCursor
                        Dim Path As String = "G:\Produktbilder\" & ArtikelnrTextBoxKlanning.Text.ToString & "\"

                        If Directory.Exists(Path) Then

                            'radera befintliga filer i mappen
                            For Each deleteFile In Directory.GetFiles(Path, "*.*", SearchOption.TopDirectoryOnly)
                                File.Delete(deleteFile)
                            Next

                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)

                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxKlanning.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox5.ImageLocation = Nothing
                                Else
                                    PictureBox5.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button27.PerformClick()
                            Cursor.Current = Cursors.Default

                        Else
                            Cursor.Current = Cursors.WaitCursor
                            Directory.CreateDirectory(Path & ArtikelnrTextBoxKlanning.Text)
                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)
                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxKlanning.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox5.ImageLocation = Nothing
                                Else
                                    PictureBox5.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button27.PerformClick()
                            Cursor.Current = Cursors.Default

                        End If



                    Catch ex As Exception

                    End Try

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (mystream IsNot Nothing) Then
                    mystream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub SetImagesTopp_Click(sender As Object, e As EventArgs) Handles SetImagesTopp.Click

        Dim mystream As Stream = Nothing
        Dim openfiledialog1 As New OpenFileDialog()
        openfiledialog1.Multiselect = True
        openfiledialog1.InitialDirectory = "C:\temp"
        'openfiledialog1.Filter = "JPG images (*.jpg)"
        If openfiledialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                mystream = openfiledialog1.OpenFile()
                If (mystream IsNot Nothing) Then
                    ' fundera ut hur programmet ska skapa katalog och sen kopiera vald fil till katalogen

                    Try
                        Cursor.Current = Cursors.WaitCursor
                        Dim Path As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text.ToString & "\"

                        If Directory.Exists(Path) Then

                            'radera befintliga filer i mappen
                            For Each deleteFile In Directory.GetFiles(Path, "*.*", SearchOption.TopDirectoryOnly)
                                File.Delete(deleteFile)
                            Next

                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)

                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox6.ImageLocation = Nothing
                                Else
                                    PictureBox6.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            WriteImagesTopparBtn.PerformClick()
                            Cursor.Current = Cursors.Default

                        Else
                            Cursor.Current = Cursors.WaitCursor
                            Directory.CreateDirectory(Path & ArtikelnrTextBoxToppar.Text)
                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)
                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox6.ImageLocation = Nothing
                                Else
                                    PictureBox6.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            WriteImagesTopparBtn.PerformClick()
                            Cursor.Current = Cursors.Default

                        End If



                    Catch ex As Exception

                    End Try

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (mystream IsNot Nothing) Then
                    mystream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub SetImagesJumpsuit_Click(sender As Object, e As EventArgs) Handles SetImagesJumpsuit.Click

        Dim mystream As Stream = Nothing
        Dim openfiledialog1 As New OpenFileDialog()
        openfiledialog1.Multiselect = True
        openfiledialog1.InitialDirectory = "C:\temp"
        'openfiledialog1.Filter = "JPG images (*.jpg)"
        If openfiledialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                mystream = openfiledialog1.OpenFile()
                If (mystream IsNot Nothing) Then
                    ' fundera ut hur programmet ska skapa katalog och sen kopiera vald fil till katalogen

                    Try
                        Cursor.Current = Cursors.WaitCursor
                        Dim Path As String = "G:\Produktbilder\" & ArtikelnrTextBoxJumpsuit.Text.ToString & "\"

                        If Directory.Exists(Path) Then

                            'radera befintliga filer i mappen
                            For Each deleteFile In Directory.GetFiles(Path, "*.*", SearchOption.TopDirectoryOnly)
                                File.Delete(deleteFile)
                            Next

                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)

                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxJumpsuit.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox7.ImageLocation = Nothing
                                Else
                                    PictureBox7.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button61.PerformClick()
                            Cursor.Current = Cursors.Default

                        Else
                            Cursor.Current = Cursors.WaitCursor
                            Directory.CreateDirectory(Path & ArtikelnrTextBoxJumpsuit.Text)
                            For Each file In openfiledialog1.FileNames

                                Dim filnamn As String = file.Substring(openfiledialog1.FileName.Length - 12)

                                FileCopy(file, Path & filnamn)
                                FileCopy(file, "G:\Produktbilder\" & filnamn)
                            Next
                            'test
                            Dim dir As String = "G:\Produktbilder\" & ArtikelnrTextBoxToppar.Text.ToString & "\"
                            If Directory.Exists(dir) Then
                                Dim fileEntries As String() = Directory.GetFiles(dir)
                                If fileEntries.Length = 0 Then
                                    PictureBox7.ImageLocation = Nothing
                                Else
                                    PictureBox7.ImageLocation = fileEntries.First
                                End If

                            End If
                            'test slutar

                            Threading.Thread.Sleep(1000)

                            Button61.PerformClick()
                            Cursor.Current = Cursors.Default

                        End If



                    Catch ex As Exception

                    End Try

                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (mystream IsNot Nothing) Then
                    mystream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub Button70_Click(sender As Object, e As EventArgs) Handles Button70.Click
        Form2.Show()
    End Sub

    Private Sub Button71_Click(sender As Object, e As EventArgs) Handles Button71.Click
        Form2.Show()
    End Sub

    Private Sub Button72_Click(sender As Object, e As EventArgs) Handles Button72.Click
        Form2.Show()
    End Sub

    Private Sub Button73_Click(sender As Object, e As EventArgs) Handles Button73.Click
        Form2.Show()
    End Sub

    Private Sub Button74_Click(sender As Object, e As EventArgs) Handles Button74.Click
        Form2.Show()
    End Sub

    Private Sub Button75_Click(sender As Object, e As EventArgs) Handles Button75.Click
        Form2.Show()
    End Sub

    Private Sub Button76_Click(sender As Object, e As EventArgs) Handles Button76.Click
        Form2.Show()
    End Sub

    Private Sub Button77_Click(sender As Object, e As EventArgs) Handles Button77.Click

        If Ang_storlekTextBoxByxor.Text = Nothing Then
            Ang_storlekTextBoxByxor.Text = "Saknas"
        Else
            Ang_storlekTextBoxByxor.Text = Nothing
        End If

    End Sub

    Private Sub Button78_Click(sender As Object, e As EventArgs) Handles Button78.Click

        If MärkeTextBoxByxor.Text = Nothing Then
            MärkeTextBoxByxor.Text = "Ej angivet"
        Else
            MärkeTextBoxByxor.Text = Nothing
        End If

    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button79_Click(sender As Object, e As EventArgs) Handles Button79.Click

        If Ang_storlekTextBoxJacka.Text = Nothing Then
            Ang_storlekTextBoxJacka.Text = "Saknas"
        Else
            Ang_storlekTextBoxJacka.Text = Nothing
        End If

    End Sub

    Private Sub Button80_Click(sender As Object, e As EventArgs) Handles Button80.Click

        If MärkeTextBoxJacka.Text = Nothing Then
            MärkeTextBoxJacka.Text = "Ej angivet"
        Else
            MärkeTextBoxJacka.Text = Nothing
        End If

    End Sub

    Private Sub Button81_Click(sender As Object, e As EventArgs) Handles Button81.Click

        If Ang_storlekTextBoxKjol.Text = Nothing Then
            Ang_storlekTextBoxKjol.Text = "Saknas"
        Else
            Ang_storlekTextBoxKjol.Text = Nothing
        End If

    End Sub

    Private Sub Button82_Click(sender As Object, e As EventArgs) Handles Button82.Click

        If MärkeTextBoxKjol.Text = Nothing Then
            MärkeTextBoxKjol.Text = "Ej angivet"
        Else
            MärkeTextBoxKjol.Text = Nothing
        End If

    End Sub

    Private Sub Button83_Click(sender As Object, e As EventArgs) Handles Button83.Click

        If Ang_storlekTextBoxKlanning.Text = Nothing Then
            Ang_storlekTextBoxKlanning.Text = "Saknas"
        Else
            Ang_storlekTextBoxKlanning.Text = Nothing
        End If

    End Sub

    Private Sub Button84_Click(sender As Object, e As EventArgs) Handles Button84.Click

        If MärkeTextBoxKlanning.Text = Nothing Then
            MärkeTextBoxKlanning.Text = "Ej angivet"
        Else
            MärkeTextBoxKlanning.Text = Nothing
        End If

    End Sub

    Private Sub Button85_Click(sender As Object, e As EventArgs) Handles Button85.Click

        If Ang_storlekTextBoxToppar.Text = Nothing Then
            Ang_storlekTextBoxToppar.Text = "Saknas"
        Else
            Ang_storlekTextBoxToppar.Text = Nothing
        End If

    End Sub

    Private Sub Button86_Click(sender As Object, e As EventArgs) Handles Button86.Click

        If MärkeTextBoxToppar.Text = Nothing Then
            MärkeTextBoxToppar.Text = "Ej angivet"
        Else
            MärkeTextBoxToppar.Text = Nothing
        End If

    End Sub

    Private Sub Button87_Click(sender As Object, e As EventArgs) Handles Button87.Click

        If Ang_storlekTextBoxJumpsuit.Text = Nothing Then
            Ang_storlekTextBoxJumpsuit.Text = "Saknas"
        Else
            Ang_storlekTextBoxJumpsuit.Text = Nothing
        End If

    End Sub

    Private Sub Button88_Click(sender As Object, e As EventArgs) Handles Button88.Click

        If MärkeTextBoxJumpsuit.Text = Nothing Then
            MärkeTextBoxJumpsuit.Text = "Ej angivet"
        Else
            MärkeTextBoxJumpsuit.Text = Nothing
        End If

    End Sub

    Private Sub Button89_Click(sender As Object, e As EventArgs) Handles Button89.Click

        If Angiven_storlekTextBoxAcc.Text = Nothing Then
            Angiven_storlekTextBoxAcc.Text = "Saknas"
        Else
            Angiven_storlekTextBoxAcc.Text = Nothing

        End If

    End Sub

    Private Sub Button90_Click(sender As Object, e As EventArgs) Handles Button90.Click

        If MärkeTextBoxAcc.Text = Nothing Then
            MärkeTextBoxAcc.Text = "Ej angivet"
        Else
            MärkeTextBoxAcc.Text = Nothing
        End If

    End Sub

    Private Sub Button91_Click(sender As Object, e As EventArgs) Handles Button91.Click

        Dim text As String
        text = SV_Detaljerad_beskrivningTextBox4.Text & vbNewLine & vbNewLine & "• Ungefärlig storlek: " & Vår_storlekTextBoxByxor.Text & vbNewLine & "• Måttbeskrivning : Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag." & vbNewLine & "• Märke: " & MärkeTextBoxByxor.Text & vbNewLine & "• Material: " & MaterialTextBoxByxor.Text & vbNewLine & "• Midjemått: " & MidjemåttTextBoxByxor.Text & vbNewLine & "• Midjemått utan stretch: " & MidjausTextboxByxor.Text & vbNewLine & "• Midjemått med 'bekväm stretch' :" & MidjamsTextboxByxor.Text & vbNewLine & "• Lårvidd: " & LårmåttTextboxByxor.Text & vbNewLine & "• Benlängd utsida (från midja och nedåt till benslut): " & Yttre_längdTextBoxByxor.Text & vbNewLine & "• Benlängd insida: " & BeninnerlängdTextBoxByxor.Text & vbNewLine & vbNewLine & "Se gärna våra andra byxor! Du hittar dem här: https://www.tradera.com/profile/items/4351205/ra-rely?categoryId=1631&expanded=1"

        Clipboard.SetText(text)

    End Sub

    Private Sub Button92_Click(sender As Object, e As EventArgs) Handles Button92.Click

        Dim text As String
        text = SV_Detaljerad_beskrivningTextBox5.Text & vbNewLine & vbNewLine & "• Ungefärlig storlek: " & Vår_storlekTextBoxJacka.Text & vbNewLine & "• Måttbeskrivning : Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag." & vbNewLine & "• Märke: " & MärkeTextBoxJacka.Text & vbNewLine & "• Material: " & MaterialTextBoxJacka.Text & vbNewLine & "• Byst: " & BystTextBoxJacka.Text & "• Midjemått: " & MidjaTextBoxJacka.Text & vbNewLine & "• Rygglängd (mätt från plaggets nacke och ned): " & RygglängdTextBoxJacka.Text & vbNewLine & "• Sidlängd (mätt från ärmsömm och ned): " & SidlängdTextBoxJacka.Text & vbNewLine & "• Ärmlängd utsida (mätt från axelslut): " & Yttre_ärmTextBoxJacka.Text & vbNewLine & "• Ärmlängd insida (mätt från ärmhålssöm utåt): " & Inre_ärmTextBoxJacka.Text & vbNewLine & "• Slitslängd: " & SlitslängdTextBoxJacka.Text & vbNewLine & "• Fodrad: " & FodradComboBoxJacka.Text & vbNewLine & vbNewLine & "Se gärna våra andra jackor! Du hittar dem här: https://www.tradera.com/profile/items/4351205/ra-rely?categoryId=1633&expanded=1"

        Clipboard.SetText(text)

    End Sub

    Private Sub Button93_Click(sender As Object, e As EventArgs) Handles Button93.Click

        Dim text As String
        text = SV_Detaljerad_beskrivningTextBox3.Text & vbNewLine & vbNewLine & "• Ungefärlig storlek: " & Vår_storkelTextBoxKjol.Text & vbNewLine & "• Måttbeskrivning : Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag." & vbNewLine & "• Märke: " & MärkeTextBoxKjol.Text & vbNewLine & "• Material: " & MaterialTextBoxKjol.Text & vbNewLine & "• Midjemått: " & MidjaTextBoxKjol.Text & vbNewLine & "• Midjemått utan stretch: " & Midja_usTextBoxKjol.Text & vbNewLine & "• Midjemått med 'bekväm stretch': " & Midja_msTextBoxKjol.Text & vbNewLine & "• Längd: " & LängdTextBoxKjol.Text & vbNewLine & "• Fodrad: " & FodradComboBoxKjol.Text & vbNewLine & vbNewLine & "Se gärna våra andra fina kjolar som du hittar här: https://www.tradera.com/profile/items/4351205/ra-rely?categoryId=340347&expanded=1"

        Clipboard.SetText(text)

    End Sub

    Private Sub Button94_Click(sender As Object, e As EventArgs) Handles Button94.Click

        Dim text As String
        text = SV_Detaljerad_beskrivningTextBox2.Text & vbNewLine & vbNewLine & "• Ungefärlig storlek: " & Vår_storlekTextBoxKlanning.Text & vbNewLine & "• Måttbeskrivning : Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag." & vbNewLine & "• Märke: " & MärkeTextBoxKlanning.Text & vbNewLine & "• Material: " & MaterialTextBoxKlanning.Text & vbNewLine & "• Byst: " & BystTextBoxKlanning.Text & vbNewLine & "• Bystmått utan stretch: " & Byst_usTextBoxKlanning.Text & vbNewLine & "• Bystmått med 'bekväm stretch'" & Byst_msTextBoxKlanning.Text & vbNewLine & "• Höftmått: " & HöftTextBoxKlanning.Text & vbNewLine & "• Höftmått utan stretch: " & Höft_usTextBoxKlanning.Text & vbNewLine & "• Höftmått med stretch: " & Höft_msTextBoxKlanning.Text & vbNewLine & "• Rygglängd (mätt från plaggets nacke och ned): " & RygglängdTextBoxKlanning.Text & vbNewLine & "• Sidlängd (mätt från ärmsömm och ned): " & SidlängdTextBoxKlanning.Text & vbNewLine & "• Ärmlängd utsida (mätt från axelslut): " & Yttre_ärmTextBoxKlanning.Text & vbNewLine & "• Ärmlängd insida (mätt från ärmhålssöm utåt): " & Inre_ärmTextBoxKlanning.Text & vbNewLine & vbNewLine & "Se fler av våra fina klänningar här: https://www.tradera.com/profile/items/4351205/ra-rely?categoryId=301741&expanded=1"

        Clipboard.SetText(text)

    End Sub

    Private Sub Button95_Click(sender As Object, e As EventArgs) Handles Button95.Click

        Dim text As String
        text = SV_Detaljerad_beskrivningTextBox.Text & vbNewLine & vbNewLine & "• Ungefärlig storlek: " & Vår_storlekTextBoxTopp.Text & vbNewLine & "• Måttbeskrivning : Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag." & vbNewLine & "• Märke: " & MärkeTextBoxToppar.Text & vbNewLine & "• Material: " & MaterialTextBoxToppar.Text & vbNewLine & "• Byst: " & BystTextBoxTopparar.Text & vbNewLine & "• Bystmått utan stretch: " & Byst_usTextBoxToppar.Text & vbNewLine & "• Bystmått med 'bekväm stretch': " & Byst_msTextBoxToppar.Text & "• Midjemått: " & MidjaTextBoxToppar.Text & vbNewLine & "• Midjemått utan stretch: " & Midja_usTextBoxToppar.Text & vbNewLine & "• Midjemått med 'bekväm stretch': " & Midja_msTextBoxToppar.Text & vbNewLine & "• Rygglängd (mätt från plaggets nacke och ned): " & RygglängdTextBoxToppar.Text & vbNewLine & "• Sidlängd (mätt från ärmsömm och ned): " & SidlängdTextBoxToppar.Text & vbNewLine & "• Ärmlängd utsida (mätt från axelslut): " & Ärmlängd_utsidaTextBoxToppar.Text & vbNewLine & "• Ärmlängd insida (mätt från ärmhålssöm utåt): " & Ärmlängd_insidaTextBoxToppar.Text & vbNewLine & "• Fodrad: " & FodradComboBoxToppar.Text & vbNewLine & vbNewLine & "Se gärna våra andra fina toppar och kläder här: https://www.tradera.com/profile/items/4351205/ra-rely?categoryId=302073&expanded=1"

        Clipboard.SetText(text)

    End Sub

    Private Sub Button96_Click(sender As Object, e As EventArgs) Handles Button96.Click

        Dim text As String
        text = SV_Detaljerad_beskrivningTextBoxJumpsuit.Text & vbNewLine & vbNewLine & "• Ungefärlig storlek: " & Vår_storlekTextBoxJumpsuit.Text & vbNewLine & "• Måttbeskrivning : Samtliga mått är ungefärliga mått, angivna i cm och tagna med plagget liggandes på plant underlag." & vbNewLine & "• Märke: " & MärkeTextBoxJumpsuit.Text & "• Material: " & MaterialTextBoxJumpsuit.Text & vbNewLine & "• Byst: " & BystTextBoxJumpsuit.Text & vbNewLine & "• Bystmått utan stretch: " & Byst_usTextBoxJumpsuit.Text & vbNewLine & "• Bystmått med 'bekväm stretch': " & Byst_msTextBoxJumpsuit.Text & vbNewLine & "• Midjemått: " & MidjaTextBoxJumpsuit.Text & vbNewLine & "• Midjemått utan stretch: " & Midja_usTextBoxJumpsuit.Text & vbNewLine & "• Midjemått med 'bekväm stretch': " & Midja_msTextBoxJumpsuit.Text & vbNewLine & "• Rygglängd (mätt från plaggets nacke och ned): " & RygglängdTextBoxJumpsuit.Text & vbNewLine & "• Sidlängd (mätt från ärmsömm och ned): " & SidlängdTextBoxJumpsuit.Text & vbNewLine & "• Fodrad: " & FodradComboBoxJumpsuit.Text & vbNewLine & vbNewLine & "Se gärna våra andra korta och långa jumpsuits / byxdressar här: https://www.tradera.com/profile/items/4351205/ra-rely?categoryId=342757&expanded=1"

        Clipboard.SetText(text)

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub Button55_Click_1(sender As Object, e As EventArgs) Handles Button55.Click

        Me.Dam_ByxorTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Byxor)

        Panel2.Enabled = False
        Button51.Enabled = False
        Button52.Enabled = True

    End Sub

    Private Sub Button97_Click(sender As Object, e As EventArgs) Handles Button97.Click

        Me.Dam_JackaTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Jacka)

        Panel3.Enabled = False

        Button43.Enabled = True
        Button42.Enabled = False
        Button55.Enabled = False
        Button103.Enabled = True
        Button104.Enabled = True

    End Sub

    Private Sub Button98_Click(sender As Object, e As EventArgs) Handles Button98.Click

        Me.Dam_KjolTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Kjol)

        Panel4.Enabled = True
        Button17.Enabled = True
        Button16.Enabled = False
        Button98.Enabled = False
        Button97.Enabled = True
        Button105.Enabled = True

    End Sub

    Private Sub Button99_Click(sender As Object, e As EventArgs) Handles Button99.Click

        Me.Dam_KlänningTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Klänning)

        Panel5.Enabled = False

        Button24.Enabled = False
        Button25.Enabled = True
        Button99.Enabled = False
        Button106.Enabled = True

    End Sub

    Private Sub Button100_Click(sender As Object, e As EventArgs) Handles Button100.Click

        Me.Dam_ToppTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Topp)

        Panel6.Enabled = False

        Button2.Enabled = False
        Button100.Enabled = False
        Button1.Enabled = True
        Button107.Enabled = True

    End Sub

    Private Sub Button101_Click(sender As Object, e As EventArgs) Handles Button101.Click

        Me.Dam_JumpsuitTableAdapter.Fill(Me.CLOREGDBDataSet.Dam_Jumpsuit)

        Panel7.Enabled = False

        Button63.Enabled = False
        Button101.Enabled = False
        Button62.Enabled = True
        Button108.Enabled = True

    End Sub

    Private Sub Button102_Click(sender As Object, e As EventArgs) Handles Button102.Click

        Panel1.Enabled = True

        Button34.Enabled = False
        AvbrytAccButton.Enabled = True
        Button33.Enabled = True
        Button102.Enabled = False

    End Sub

    Private Sub Button103_Click(sender As Object, e As EventArgs) Handles Button103.Click

        Panel2.Enabled = True

        Button52.Enabled = False
        Button55.Enabled = True
        Button51.Enabled = True
        Button103.Enabled = False

    End Sub

    Private Sub Button104_Click(sender As Object, e As EventArgs) Handles Button104.Click

        Panel3.Enabled = True

        Button43.Enabled = False
        Button97.Enabled = True
        Button42.Enabled = True
        Button104.Enabled = False

    End Sub

    Private Sub Button105_Click(sender As Object, e As EventArgs) Handles Button105.Click

        Panel4.Enabled = True

        Button17.Enabled = False
        Button98.Enabled = True
        Button16.Enabled = True
        Button105.Enabled = False

    End Sub

    Private Sub Button106_Click(sender As Object, e As EventArgs) Handles Button106.Click

        Panel5.Enabled = True

        Button25.Enabled = False
        Button99.Enabled = True
        Button24.Enabled = True
        Button106.Enabled = False

    End Sub

    Private Sub Button107_Click(sender As Object, e As EventArgs) Handles Button107.Click

        Panel6.Enabled = True

        Button1.Enabled = False
        Button100.Enabled = True
        Button2.Enabled = True
        Button107.Enabled = False

    End Sub

    Private Sub Button108_Click(sender As Object, e As EventArgs) Handles Button108.Click

        Panel7.Enabled = True

        Button62.Enabled = False
        Button101.Enabled = True
        Button63.Enabled = True
        Button108.Enabled = False

    End Sub
End Class
