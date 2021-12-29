Imports MySql.Data.MySqlClient

Imports iTextSharp
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO

Public Class Home

    Function execCommand(ByVal cmd As MySqlCommand) As Boolean

        If constring.State = ConnectionState.Closed Then
            constring.Open()
        End If

        Try
            If cmd.ExecuteNonQuery() = 1 Then
                Return True

            Else
                Return False
            End If
        Catch ex As Exception

            MessageBox.Show("ERROR")
            Return False

        End Try

        If constring.State = ConnectionState.Open Then
            constring.Close()
        End If
    End Function

    Private Sub Home_Load(sender As Object, e As EventArgs) Handles Me.Load
        Panelgeneraldata.Hide()
        PanelgeneraldataPRD.Hide()
        PanelgeneraldataHE.Hide()
        PanelPOF.Hide()
        PanelPOFHE.Hide()
        PanelPOFPRD.Hide()
        PanelCOF.Hide()
        PanelCOFHE.Hide()

        Panel4.Hide()
        Button15.Text = ">"
        Panelnewanalysis.Show()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If Button15.Text = "<" Then
            Panel4.Hide()
            Button15.Text = ">"
        Else
            Button15.Text = ">"
            Panel4.Show()
            Button15.Text = "<"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Panel6.Visible = True
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        PictureBox2.Visible = True
        Button13.Visible = True
        Panel18.Visible = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Panel87.Visible = False Then
            Panel87.Visible = True
        ElseIf Panel87.Visible = True Then
            Panel87.Visible = False
        End If

        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False




        PictureBox2.Visible = False
        Button13.Visible = False
        Panel18.Visible = True


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = True
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Enabled = True
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabPage1.Text = "General Data"
        TabPage2.Text = ""
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabControl1.SelectedTab = TabPage1

        df1.Hide()
        df2.Hide()
        df3.Hide()
        df4.Hide()
        df5.Hide()
        df6.Hide()
        df7.Hide()
        df8.Hide()
        df9.Hide()
        df10.Hide()
        df11.Hide()
        df12.Hide()
        df13.Hide()
        df14.Hide()
        df15.Hide()
        df16.Hide()
        df17.Hide()
        df18.Hide()
        df19.Hide()
        df20.Hide()
        df21.Hide()

        Call tampildatamaterialcombobox()
        Call tampildatafluidacombobox()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = True
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Text = ""
        TabPage2.Text = "POF Analysis"
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabPage1.Enabled = False
        TabPage2.Enabled = True
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = True
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = True
        TabPage4.Enabled = False


        TabPage1.Text = ""
        TabPage2.Text = ""
        TabPage3.Text = "COF Analysis"
        TabPage4.Text = ""


        TabControl1.SelectedTab = TabPage3
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = True
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = True


        TabPage1.Text = ""
        TabPage2.Text = ""
        TabPage3.Text = ""
        TabPage4.Text = "Risk Calculation"


        TabControl1.SelectedTab = TabPage4
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = True
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabPage1.Text = ""
        TabPage2.Text = ""
        TabPage3.Text = ""
        TabPage4.Text = ""

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = True
        Panel14.Visible = False

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = True

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        PictureBox3.Visible = True
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("KODRUM")
        ComboBox62.Items.Add("DRUM")
        ComboBox62.Items.Add("REACTOR")
        ComboBox62.Items.Add("COLTOP")
        ComboBox62.Items.Add("COLMID")
        ComboBox62.Items.Add("COLBTM")
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = True
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("PIPE-1")
        ComboBox62.Items.Add("PIPE-3")
        ComboBox62.Items.Add("PIPE-4")
        ComboBox62.Items.Add("PIPE-6")
        ComboBox62.Items.Add("PIPE-8")
        ComboBox62.Items.Add("PIPE-10")
        ComboBox62.Items.Add("PIPE-13")
        ComboBox62.Items.Add("PIPE-16")
        ComboBox62.Items.Add("PIPEGT16")
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = True
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("TANKBOTTOM")
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = True
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("FINFAN")
        ComboBox62.Items.Add("FILTER")
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = True
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("COMPC")
        ComboBox62.Items.Add("COMPR")
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = True
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("PUMP3S")
        ComboBox62.Items.Add("PUMP1S")
        ComboBox62.Items.Add("PUMPR")
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = True
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("HEXSS")
        ComboBox62.Items.Add("HEXTS")
        ComboBox62.Items.Add("HEXTUBE")
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = True
        PictureBox11.Visible = False
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("Pressure Relief Devices")
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = True
        PictureBox12.Visible = False

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("COURSES-1-10")
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
        PictureBox12.Visible = True

        ComboBox62.Items.Clear()
        ComboBox62.Items.Add("Heat Exchanger Tube Bundles")
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        If PictureBox3.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        ElseIf PictureBox7.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        ElseIf PictureBox9.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        ElseIf PictureBox6.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        ElseIf PictureBox4.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        ElseIf PictureBox8.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        ElseIf PictureBox11.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf PictureBox5.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        ElseIf PictureBox10.Visible = True Then
            Panelgeneraldata.Hide()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Show()
            Panelnewanalysis.Hide()
        ElseIf PictureBox12.Visible = True Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Show()
            PanelgeneraldataPRD.Hide()
            Panelnewanalysis.Hide()
        End If

        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = True
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False
        Panel87.Visible = True


        TabPage1.Enabled = True
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabPage1.Text = "General Data"
        TabPage2.Text = ""
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabControl1.SelectedTab = TabPage1

        df1.Hide()
        df2.Hide()
        df3.Hide()
        df4.Hide()
        df5.Hide()
        df6.Hide()
        df7.Hide()
        df8.Hide()
        df9.Hide()
        df10.Hide()
        df11.Hide()
        df12.Hide()
        df13.Hide()
        df14.Hide()
        df15.Hide()
        df16.Hide()
        df17.Hide()
        df18.Hide()
        df19.Hide()
        df20.Hide()
        df21.Hide()

        Call tampildatamaterialcombobox()
        Call tampildatafluidacombobox()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
        Start.Close()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

    End Sub

    Private Sub UnitsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UnitsToolStripMenuItem.Click
        Units.ShowDialog()
    End Sub

    Private Sub FluidsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FluidsToolStripMenuItem.Click
        Risk_Matrix.ShowDialog()
    End Sub

    Private Sub GuidanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuidanceToolStripMenuItem.Click
        Fluids.ShowDialog()
    End Sub

    Private Sub LicenseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LicenseToolStripMenuItem.Click
        Materials.ShowDialog()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Industry_Units.ShowDialog()
    End Sub

    Private Sub EquipmentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EquipmentToolStripMenuItem.Click
        Equipment.ShowDialog()
    End Sub

    Private Sub FactorManagementSystemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FactorManagementSystemToolStripMenuItem.Click
        FMS.ShowDialog()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        FMS.ShowDialog()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = True
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        PictureBox2.Visible = False
        Button13.Visible = False
        Panel18.Visible = True

        TabPage1.Enabled = True
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabPage1.Text = "Equipment"
        TabPage2.Text = ""
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs)
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = True
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Text = ""
        TabPage2.Text = "General Data"
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabPage1.Enabled = False
        TabPage2.Enabled = True
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs)
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = True
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Text = ""
        TabPage2.Text = "General Data"
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabPage1.Enabled = False
        TabPage2.Enabled = True
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabControl1.SelectedTab = TabPage2
    End Sub

    'NEXT TO POF
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = True
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False

        PanelPOF.Show()
        PanelPOFHE.Hide()
        PanelPOFPRD.Hide()

        TabPage1.Text = ""
        TabPage2.Text = "POF Analysis"
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabPage1.Enabled = False
        TabPage2.Enabled = True
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabControl1.SelectedTab = TabPage2

        If RadioButton3.Checked = True Then
            df2.Show()
        ElseIf RadioButton4.Checked = True Then
            df2.Hide()
        End If

        If RadioButton1.Checked = True Then
            df1.Show()
        ElseIf RadioButton2.Checked = True Then
            df1.Show()
        End If

        If ComboBox1.Text = "Carbon Steel" Or ComboBox1.Text = "Low Alloy Steel" Then
            If RadioButton13.Checked = True Then
                df3.Show()
                df4.Hide()
                df5.Hide()
                df6.Hide()
                df7.Hide()
            ElseIf RadioButton12.Checked = True Then
                df3.Hide()
                df4.Show()
                df5.Hide()
                df6.Hide()
                df7.Hide()
            ElseIf RadioButton5.Checked = True Then
                df3.Hide()
                df4.Hide()
                df5.Show()
                df6.Show()
                df7.Hide()
            ElseIf RadioButton6.Checked = True Then
                df3.Hide()
                df4.Hide()
                df5.Hide()
                df6.Hide()
                df7.Show()
            ElseIf RadioButton7.Checked = True Then
                df3.Hide()
                df4.Hide()
                df5.Hide()
                df6.Hide()
                df7.Hide()
            ElseIf RadioButton8.Checked = True Then
                df10.Show()
                df11.Show()
            End If

            If RadioButton15.Checked = True Or RadioButton24.Checked = True Then
                df17.Show()
            ElseIf RadioButton14.Checked = True Or RadioButton25.Checked = True Then
                df17.Hide()
            End If

            If Val(TextBox5.Text) > 177 And RadioButton16.Checked = True Then
                df16.Show()
            Else
                df16.Hide()
            End If
        End If

        If ComboBox1.Text = "Austenitic Stainless Steel" Or ComboBox1.Text = "(<42%) Nickel Based Alloys" Then
            If RadioButton9.Checked = True Then
                df8.Show()
            ElseIf RadioButton10.Checked = True Then
                If TextBox5.Text > 38 Then
                    df9.Show()
                ElseIf Val(TextBox5.Text) >= 50 And Val(TextBox5.Text) <= 500 Then
                    df9.Show()
                    df14.Show()
                    If RadioButton23.Checked = True Then
                        df15.Show()
                    End If
                End If
            ElseIf TextBox5.Text >= 593 And TextBox5.Text <= 927 Then
                df20.Show()
            End If
        End If

        If ComboBox1.Text = "Cr-1/2 Mo or Cr-Mo Low Alloy Steel" Then
            If Val(TextBox5.Text) > 177 And RadioButton16.Checked = True Then
                df16.Show()
            Else
                df16.Hide()
            End If

            If Val(TextBox5.Text) >= 343 And Val(TextBox5.Text) <= 577 Then
                df18.Show()
            Else
                df18.Hide()
            End If
        End If

        If ComboBox1.Text = "High Chromium ( >12% Cp ) Ferritic Steel" Then
            If Val(TextBox5.Text) >= 371 And Val(TextBox5.Text) <= 566 Then
                df19.Show()
            Else
                df19.Hide()
            End If
        End If

        If RadioButton18.Checked = True Then
            df12.Show()
        ElseIf RadioButton19.Checked = True Then
            df12.Hide()
        End If

        If RadioButton20.Checked = True Then
            df13.Show()
        ElseIf RadioButton21.Checked = True Then
            df13.Hide()
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = True
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False

        PanelPOF.Hide()
        PanelPOFHE.Show()
        PanelPOFPRD.Hide()

        TabPage1.Text = ""
        TabPage2.Text = "POF Analysis"
        TabPage3.Text = ""
        TabPage4.Text = ""


        TabPage1.Enabled = False
        TabPage2.Enabled = True
        TabPage3.Enabled = False
        TabPage4.Enabled = False


        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = True
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = True
        TabPage4.Enabled = False


        TabPage1.Text = ""
        TabPage2.Text = ""
        TabPage3.Text = "COF Analysis"
        TabPage4.Text = ""


        TabControl1.SelectedTab = TabPage3

        PanelCOF.Show()
        PanelCOFHE.Hide()
        PanelCOFPRD.Hide()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = True
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = True


        TabPage1.Text = ""
        TabPage2.Text = ""
        TabPage3.Text = ""
        TabPage4.Text = "Risk Calculation"


        TabControl1.SelectedTab = TabPage4

        Call liningdf()
    End Sub

    Private Sub Button19_Click_1(sender As Object, e As EventArgs) Handles Button19.Click
        Panel6.Visible = False
        Panel7.Visible = True
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = True
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False


        TabPage1.Enabled = False
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = True


        TabPage1.Text = ""
        TabPage2.Text = ""
        TabPage3.Text = ""
        TabPage4.Text = "Risk Calculation"


        TabControl1.SelectedTab = TabPage4
    End Sub

    'POF HE
    Private Sub RadioButton44_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton44.CheckedChanged
        Panelreliabilitylibrary.Show()
        Panelweibull.Hide()
        Panelmttf.Hide()
        Panelhistory.Hide()
    End Sub

    Private Sub RadioButton45_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton45.CheckedChanged
        Panelreliabilitylibrary.Hide()
        Panelweibull.Show()
        Panelmttf.Hide()
        Panelhistory.Hide()
    End Sub

    Private Sub RadioButton46_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton46.CheckedChanged
        Panelreliabilitylibrary.Hide()
        Panelweibull.Hide()
        Panelmttf.Show()
        Panelhistory.Hide()
    End Sub

    Private Sub RadioButton47_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton47.CheckedChanged
        Panelreliabilitylibrary.Hide()
        Panelweibull.Hide()
        Panelmttf.Hide()
        Panelhistory.Show()
    End Sub
    Private Sub RadioButton14_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton14.CheckedChanged
        Label90.Visible = True
        RadioButton24.Visible = True
        RadioButton25.Visible = True
    End Sub

    Private Sub RadioButton15_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton15.CheckedChanged
        Label90.Visible = False
        RadioButton24.Visible = False
        RadioButton25.Visible = False
        RadioButton24.Checked = False
        RadioButton25.Checked = False
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        MessageBox.Show("a) Areas exposed to mist overspray from cooling towers,

b) Areas exposed to steam vents,

c) Areas exposed to deluge systems,

d) Areas subject to process spills, ingress of moisture, or acid vapors.

e) Carbon steel systems, operating between –12°C and 177°C (10°F and 350°F).
    External corrosion is particularly aggressive where operating temperatures
    cause frequent or continuous condensation and reevaporation of atmospheric 
    moisture,

f) Systems that do not normally operate between -12°C and 177°C (10°F and 
    350°F) but cool or heat into this range intermittently or are subjected
    to frequent outages,

g) Systems with deteriorated coating and/or wrappings,

h) Cold service equipment consistently operating below the atmospheric dew
    point.

i) Un-insulated nozzles or other protrusions components of insulated equipment
    in cold service conditions.")
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        MessageBox.Show("a) Penetrations
    1) All penetrations or breaches in the insulation jacketing systems, such as dead legs (vents, drains, and other
        similar items), hangers and other supports, valves and fittings, bolted-on pipe shoes, ladders, and platforms.
    2) Steam tracer tubing penetrations.
    3) Termination of insulation at flanges and other components.
    4) Poorly designed insulation support rings.
    5) Stiffener rings.

b) Damaged Insulation Areas
    1) Damaged or missing insulation jacketing.
    2) Termination of insulation in a vertical pipe or piece of equipment.
    3) Caulking that has hardened, has separated, or is missing.
    4) Bulges, staining of the jacketing system or missing bands (bulges may indicate corrosion product build-up).
    5) Low points in systems that have a known breach in the insulation system, including low points in long unsupported
        piping runs.
    6) Carbon or low alloy steel flanges, bolting, and other components under insulation in high alloy piping.

The following are some examples of other suspect areas that should be considered when performing inspection for CUI:
a) Areas exposed to mist overspray from cooling towers
b) Areas exposed to steam vents.
c) Areas exposed to deluge systems.
d) Areas subject to process spills, ingress of moisture, or acid vapors.
e) Insulation jacketing seams located on top of horizontal vessels or improperly lapped or sealed insulation systems,
f) Carbon steel systems, including those insulated for personnel protection, operating between -12°C and 175°C (10°F 
    and 350°F). CUI is particularly aggressive where operating temperatures cause frequent or continuous condensation
    and re-evaporation of atmospheric moisture.
g) Carbon steel systems that normally operate in services above 175°C (350°F) but are in intermittent service or are
    subjected to frequent outages.
h) Dead legs and attachments that protrude from the insulation and operate at a different temperature than the operating
    temperature of the active line, i.e., insulation support rings, piping/platform attachments.
i) Systems in which vibration has a tendency to inflict damage to insulation jacketing providing paths for water ingress.
j) Steam traced systems experiencing tracing leaks, especially at tubing fittings beneath the insulation
k) Systems with deteriorated coating and/or wrappings.
l) Cold service equipment consistently operating below the atmospheric dew point.
m) Inspection ports or plugs which are removed to permit thickness measurements on insulated systems represent a major
    contributor to possible leaks in insulated systems. Special attention should be paid to these locations. Promptly
    replacing and resealing of these plugs is imperative.
")
    End Sub

    'Lining DF --------------------------------------------------------------------------------------------------




    'SCC-Caustic Cracking DF ------------------------------------------------------------------------------------

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If ComboBox6.Text = "Yes" Then
            Label91.Visible = True
            Label92.Visible = False
            ComboBox7.Visible = True
            ComboBox103.Visible = False
            ComboBox105.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox6.Text = "No" Then
            Label91.Visible = False
            Label92.Visible = True
            ComboBox7.Visible = False
            ComboBox103.Visible = True
            ComboBox105.Text = ""
        Else
            Label91.Visible = False
            Label92.Visible = False
            ComboBox7.Visible = False
            ComboBox103.Visible = False
            ComboBox105.Text = ""
        End If
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If ComboBox7.Text = "Yes" Then
            ComboBox105.Text = "NONE SUSCEPTIBILITY"
            Panel42.Visible = False
        ElseIf ComboBox7.Text = "No" Then
            ComboBox105.Text = "PLEASE DETERMINE BY FFS"
            Panel42.Visible = False
        Else
            ComboBox105.Text = "HIGH SUSCEPTIBILITY"
            Panel42.Visible = False
        End If
    End Sub

    Private Sub ComboBox103_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox103.SelectedIndexChanged
        If ComboBox103.Text = "Yes" Then
            ComboBox105.Text = "NONE SUSCEPTIBILITY"
            Panel42.Visible = False
        ElseIf ComboBox103.Text = "No" Then
            ComboBox105.Text = ""
            Panel42.Visible = True
        Else
            ComboBox105.Text = ""
            Panel42.Visible = False
        End If
    End Sub

    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox26.Checked = True Then
            ComboBox105.Items.Clear()
            ComboBox105.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox105.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox105.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox105.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'SCC Amine Cracking DF --------------------------------------------------------------------------------------

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        If ComboBox8.Text = "Yes" Then
            Label102.Visible = True
            Label101.Visible = False
            ComboBox9.Visible = True
            ComboBox106.Visible = False
            ComboBox108.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox8.Text = "No" Then
            Label102.Visible = False
            Label101.Visible = True
            ComboBox9.Visible = False
            ComboBox106.Visible = True
            ComboBox108.Text = ""
        Else
            Label102.Visible = False
            Label101.Visible = False
            ComboBox9.Visible = False
            ComboBox106.Visible = False
            ComboBox108.Text = ""
        End If
    End Sub


    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        If ComboBox9.Text = "Yes" Then
            ComboBox108.Text = "NONE SUSCEPTIBILITY"
            Panel44.Visible = False
        ElseIf ComboBox9.Text = "No" Then
            ComboBox108.Text = "PLEASE DETERMINE BY FFS"
            Panel44.Visible = False
        Else
            ComboBox108.Text = "HIGH SUSCEPTIBILITY"
            Panel44.Visible = False
        End If
    End Sub

    Private Sub ComboBox106_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox106.SelectedIndexChanged
        If ComboBox106.Text = "Yes" Then
            ComboBox108.Text = "NONE SUSCEPTIBILITY"
            Panel44.Visible = False
        ElseIf ComboBox106.Text = "No" Then
            ComboBox108.Text = ""
            Panel44.Visible = True
        Else
            ComboBox108.Text = ""
            Panel44.Visible = False
        End If
    End Sub

    Private Sub CheckBox35_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox35.CheckedChanged
        If CheckBox35.Checked = True Then
            ComboBox108.Items.Clear()
            ComboBox108.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox108.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox108.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox108.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'SCC Sulfide Stress Cracking DF -----------------------------------------------------------------------------

    Private Sub ComboBox94_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox94.SelectedIndexChanged
        If ComboBox94.Text = "Yes" Then
            Label124.Visible = True
            ComboBox95.Visible = True
            ComboBox110.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox94.Text = "No" Then
            Label124.Visible = False
            ComboBox95.Visible = False
            ComboBox110.Text = ""
        Else
            Label124.Visible = False
            ComboBox95.Visible = False
            ComboBox110.Text = ""
        End If
    End Sub

    Private Sub CheckBox73_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox73.CheckedChanged
        If CheckBox73.Checked = True Then
            ComboBox110.Items.Clear()
            ComboBox110.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox110.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox110.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox110.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub


    'SCC-HIC/SOHIC-H2S DF --------------------------------------------------------------------------------------

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        If ComboBox13.Text = "Yes" Then
            Label125.Visible = True
            ComboBox17.Visible = True
            ComboBox112.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox13.Text = "No" Then
            Label125.Visible = False
            ComboBox17.Visible = False
            ComboBox112.Text = ""
        Else
            Label125.Visible = False
            ComboBox17.Visible = False
            ComboBox112.Text = ""
        End If
    End Sub

    Private Sub CheckBox76_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox76.CheckedChanged
        If CheckBox76.Checked = True Then
            ComboBox112.Items.Clear()
            ComboBox112.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox112.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox112.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox112.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'Scc Damage Factor – Alkaline Carbonate Stress Corrosion Cracking ------------------------------------------

    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox16.SelectedIndexChanged
        If ComboBox16.Text = "Yes" Then
            Label127.Visible = True
            ComboBox97.Visible = True
            ComboBox114.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox16.Text = "No" Then
            Label127.Visible = False
            ComboBox97.Visible = False
            ComboBox114.Text = ""
        Else
            Label125.Visible = False
            ComboBox17.Visible = False
            ComboBox114.Text = ""
        End If
    End Sub

    Private Sub CheckBox87_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox87.CheckedChanged
        If CheckBox87.Checked = True Then
            ComboBox114.Items.Clear()
            ComboBox114.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox114.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox114.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox114.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'Scc Damage Factor – Polythionic Acid Stress Corrosion Cracking --------------------------------------------

    Private Sub ComboBox115_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox115.SelectedIndexChanged
        If ComboBox115.Text = "Yes" Then
            Label140.Visible = True
            Label141.Visible = False
            ComboBox20.Visible = True
            ComboBox21.Visible = False
            ComboBox117.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox115.Text = "No" Then
            Label140.Visible = False
            Label141.Visible = True
            ComboBox20.Visible = False
            ComboBox21.Visible = True
            ComboBox117.Text = ""
        Else
            Label140.Visible = False
            Label141.Visible = False
            ComboBox20.Visible = False
            ComboBox21.Visible = False
            ComboBox117.Text = ""
        End If
    End Sub

    Private Sub ComboBox20_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox20.SelectedIndexChanged
        If ComboBox20.Text = "Yes" Then
            ComboBox117.Text = "NONE SUSCEPTIBILITY"
            Panel48.Visible = False
        ElseIf ComboBox20.Text = "No" Then
            ComboBox117.Text = "PLEASE DETERMINE BY FFS"
            Panel48.Visible = False
        Else
            ComboBox117.Text = "HIGH SUSCEPTIBILITY"
            Panel48.Visible = False
        End If
    End Sub

    Private Sub ComboBox21_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox21.SelectedIndexChanged
        If ComboBox21.Text = "During Operation" Then
            'ComboBox117.Text = ""
            Panel48.Visible = True
        ElseIf ComboBox21.Text = "During Shutdown" Then
            'ComboBox117.Text = ""
            Panel48.Visible = True
        ElseIf ComboBox21.Text = "None" Then
            ComboBox117.Text = "NONE SUSCEPTIBILITY"
            Panel48.Visible = False
        Else
            'ComboBox117.Text = ""
            Panel48.Visible = False
        End If
    End Sub

    Private Sub CheckBox97_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox97.CheckedChanged
        If CheckBox97.Checked = True Then
            ComboBox117.Items.Clear()
            ComboBox117.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox117.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox117.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox117.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'Scc Damage Factor – Chloride Stress Corrosion Cracking ----------------------------------------------------

    Private Sub ComboBox99_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox99.SelectedIndexChanged
        If ComboBox99.Text = "Yes" Then
            Label152.Visible = True
            ComboBox100.Visible = True
            ComboBox120.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox99.Text = "No" Then
            Label152.Visible = False
            ComboBox100.Visible = False
            ComboBox120.Text = ""
        Else
            Label152.Visible = False
            ComboBox100.Visible = False
            ComboBox120.Text = ""
        End If
    End Sub

    Private Sub ComboBox100_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox100.SelectedIndexChanged

    End Sub

    Private Sub CheckBox108_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox108.CheckedChanged
        If CheckBox108.Checked = True Then
            ComboBox120.Items.Clear()
            ComboBox120.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox120.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox120.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox120.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'Scc Damage Factor – Hydrogen Stress Cracking-Hf -----------------------------------------------------------

    Private Sub ComboBox24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox24.SelectedIndexChanged
        If ComboBox24.Text = "Yes" Then
            Label156.Visible = True
            ComboBox121.Visible = True
            ComboBox124.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox24.Text = "No" Then
            Label156.Visible = False
            ComboBox121.Visible = False
            ComboBox124.Text = ""
        Else
            Label156.Visible = False
            ComboBox121.Visible = False
            ComboBox124.Text = ""
        End If
    End Sub

    Private Sub CheckBox122_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox122.CheckedChanged
        If CheckBox122.Checked = True Then
            ComboBox124.Items.Clear()
            ComboBox124.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox124.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox124.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox124.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'Scc Damage Factor – Hic/Sohic-Hf --------------------------------------------------------------------------

    Private Sub ComboBox26_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox26.SelectedIndexChanged
        If ComboBox26.Text = "Yes" Then
            Label165.Visible = True
            Label167.Visible = False
            ComboBox27.Visible = True
            ComboBox126.Visible = False
            ComboBox125.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox26.Text = "No" Then
            Label165.Visible = False
            Label167.Visible = True
            ComboBox27.Visible = False
            ComboBox126.Visible = True
            ComboBox125.Text = ""
        Else
            Label140.Visible = False
            Label141.Visible = False
            ComboBox20.Visible = False
            ComboBox21.Visible = False
            ComboBox125.Text = ""
        End If
    End Sub

    Private Sub CheckBox127_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox127.CheckedChanged
        If CheckBox127.Checked = True Then
            ComboBox125.Items.Clear()
            ComboBox125.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox125.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox125.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox125.Items.Add("HIGH SUSCEPTIBILITY")
        End If
    End Sub

    'External Corrosion Damage Factor – Ferritic Component -----------------------------------------------------


    'Corrosion Under Insulation Damage Factor – Ferritic Component ---------------------------------------------


    'External Chloride Stress Corrosion Cracking Damage Factor – Austenitic Component --------------------------

    Private Sub ComboBox30_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox30.SelectedIndexChanged
        If ComboBox24.Text = "Yes" Then
            Label156.Visible = True
            ComboBox121.Visible = True
            ComboBox124.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox24.Text = "No" Then
            Label156.Visible = False
            ComboBox121.Visible = False
            ComboBox124.Text = ""
        Else
            Label156.Visible = False
            ComboBox121.Visible = False
            ComboBox124.Text = ""
        End If
    End Sub

    'External Chloride Stress Corrosion Cracking Under Insulation Damage Factor – Austenitic Component ---------

    Private Sub ComboBox32_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox32.SelectedIndexChanged
        If ComboBox32.Text = "Yes" Then
            Label190.Visible = True
            ComboBox129.Visible = True
            TextBox40.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox32.Text = "No" Then
            Label190.Visible = False
            ComboBox129.Visible = False
            TextBox40.Text = ""
        Else
            Label156.Visible = False
            ComboBox121.Visible = False
            TextBox40.Text = ""
        End If
    End Sub

    'High Temperature Hydrogen Attack Damage Factor ------------------------------------------------------------



    Private Sub CheckBox196_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox196.CheckedChanged
        If CheckBox196.Checked = True Then
            TextBox80.Text = FMS.TextBox2.Text
        End If
    End Sub

    'Brittle Fracture Damage Factor ----------------------------------------------------------------------------


    'Low Alloy Steel Embrittlement Damage Factor ---------------------------------------------------------------


    '885°F Embrittlement Damage Factor -------------------------------------------------------------------------


    'Sigma Phase Embrittlement Damage Factor -------------------------------------------------------------------


    'Piping Mechanical Fatigue Damage Factor -------------------------------------------------------------------

    'GFF -------------------------------------------------------------------------------------------------------

    Private Sub ComboBox62_SelectedIndexChanged(sender As Object, e As EventArgs)
        If ComboBox62.Text = "COMPC" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0
            totalfms.Text = 0.00003
        ElseIf ComboBox62.Text = "COMPR" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "HEXSS" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "HEXTS" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-1" Then
            small.Text = 0.000028
            medium.Text = 0
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-2" Then
            small.Text = 0.000028
            medium.Text = 0
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-4" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-6" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-8" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-10" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-12" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPE-16" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PIPEGT16" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PUMP2S" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PUMP1S" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "PUMPR" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "TANKBOTTOM" Then
            small.Text = 0.00072
            medium.Text = 0
            large.Text = 0
            rupture.Text = 0.000002
            totalfms.Text = 0.00072
        ElseIf ComboBox62.Text = "COURSES-1-10" Then
            small.Text = 0.00007
            medium.Text = 0.000025
            large.Text = 0.000005
            rupture.Text = 0.0000001
            totalfms.Text = 0.0001
        ElseIf ComboBox62.Text = "KODRUM" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "COLBTM" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "FINFAN" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "FILTER" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "DRUM" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "REACTOR" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "COLTOP" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox62.Text = "COLMID" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        End If

        Call tampildatamaterialcombobox()
        Call tampildatafluidacombobox()
    End Sub

    'Koneksi Database ------------------------------------------------------------------------------------------

    Public Sub tampildatamaterial()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT * FROM `tbl_material` WHERE `nama_material`=@nama AND `units`=@unit", constring)
        CMD.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text
        CMD.Parameters.Add("@nama", MySqlDbType.VarChar).Value = ComboBox3.Text

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader

        While rd.Read()
            TextBox13.Text = rd.GetDouble("yield_strength")
            TextBox14.Text = rd.GetDouble("tensile_strength")
            TextBox15.Text = rd.GetDouble("design_pressure")
            TextBox16.Text = rd.GetDouble("max_operating_pressure")
            TextBox17.Text = rd.GetDouble("design_temperature")
            TextBox18.Text = rd.GetDouble("max_design_temperature")
            TextBox19.Text = rd.GetDouble("cost_factor")
        End While
    End Sub

    Public Sub tampildatamaterialcombobox()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT `nama_material` FROM `tbl_material` WHERE `units`=@unit", constring)
        CMD.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader
        ComboBox3.Items.Clear()
        Do While rd.Read
            ComboBox3.Items.Add(rd.Item(0))
        Loop
    End Sub

    Public Sub tampildatafluida()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT * FROM `tbl_fluida` WHERE `nama_fluida`=@nama AND `units`=@unit", constring)
        CMD.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text
        CMD.Parameters.Add("@nama", MySqlDbType.VarChar).Value = ComboBox4.Text

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader

        While rd.Read()
            TextBox20.Text = rd.GetDouble("density")
            TextBox21.Text = rd.GetDouble("nbp")
            TextBox22.Text = rd.GetDouble("mw")
            TextBox23.Text = rd.GetDouble("ait")
            TextBox24.Text = rd.GetDouble("liquid_flow_velocity")
            TextBox25.Text = rd.GetDouble("viscosity")
            TextBox29.Text = rd.GetDouble("igc_a")
            TextBox30.Text = rd.GetDouble("igc_b")
            TextBox31.Text = rd.GetDouble("igc_c")
            TextBox32.Text = rd.GetDouble("igc_d")
            TextBox33.Text = rd.GetDouble("igc_e")
            TextBox26.Text = rd.GetString("fluid_type")
            TextBox27.Text = rd.GetString("ambient_state")
            TextBox28.Text = rd.GetDouble("liquid_dynamic_viscosity")
        End While
    End Sub

    Public Sub tampildatafluidacombobox()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT `nama_fluida` FROM `tbl_fluida` WHERE `units`=@unit", constring)
        CMD.Parameters.Add("@unit", MySqlDbType.VarChar).Value = Units.ComboBox1.Text

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader
        ComboBox4.Items.Clear()
        Do While rd.Read
            ComboBox4.Items.Add(rd.Item(0))
        Loop
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Call tampildatamaterial()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Call tampildatafluida()
    End Sub


    '------------------------------------------------------------------------
    'interpolasi


    Private Function Inter(ByVal temp As Double, ByVal Up1 As Double, ByVal Up2 As Double, ByVal Lo1 As Double, ByVal Lo2 As Double) As Double
        Inter = ((temp - Lo1) * (Up2 - Lo2) / (Up1 - Lo1)) + Lo2
    End Function

    'extrapolation lower

    Private Function Inter3b(ByVal upt As Double, ByVal Lot As Double, ByVal temp As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3b = ((temp - upt) * (x - z) / (Lot - upt)) + z
    End Function

    'extrapolation upper
    Private Function Inter3c(ByVal upt As Double, ByVal Lot As Double, ByVal temp As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3c = ((temp - Lot) * (z - x) / (upt - Lot)) + x
    End Function

    Private Function Intera(ByVal calhar As Double, ByVal up1a As Double, ByVal up2a As Double, ByVal lo1a As Double, ByVal lo2a As Double) As Double
        Intera = ((calhar - lo1a) * (up2a - lo2a) / (up1a - lo1a)) + lo2a
    End Function

    'exatrapolation lower

    Private Function Inter3ba(ByVal upta As Double, ByVal lota As Double, ByVal calhar As Double, ByVal xa As Double, ByVal za As Double) As Double
        Inter3ba = ((calhar - upta) * (xa - za) / (lota - upta)) + za
    End Function

    'exatrapolation upper
    Private Function Inter3ca(ByVal upta As Double, ByVal lota As Double, ByVal calhar As Double, ByVal xa As Double, ByVal za As Double) As Double
        Inter3ca = ((calhar - lota) * (za - xa) / (upta - lota)) + xa
    End Function

    Private Function Interb(ByVal moa As Double, ByVal up1b As Double, ByVal up2b As Double, ByVal lo1b As Double, ByVal lo2b As Double) As Double
        Interb = ((moa - lo1b) * (up2b - lo2b) / (up1b - lo1b)) + lo2b
    End Function

    'exbtrapolation lower

    Private Function Inter3bb(ByVal uptb As Double, ByVal lotb As Double, ByVal moa As Double, ByVal xb As Double, ByVal zb As Double) As Double
        Inter3bb = ((moa - uptb) * (xb - zb) / (lotb - uptb)) + zb
    End Function

    'exbtrapolation upper
    Private Function Inter3cb(ByVal uptb As Double, ByVal lotb As Double, ByVal moa As Double, ByVal xb As Double, ByVal zb As Double) As Double
        Inter3cb = ((moa - lotb) * (zb - xb) / (uptb - lotb)) + xb
    End Function


    Private Function Inter5(ByVal up1 As Double, ByVal Lo1 As Double, ByVal temp As Double, ByVal up2 As Double, ByVal lo2 As Double) As Double
        Inter5 = ((temp - up1) * (lo2 - up2) / (Lo1 - up1)) + up2
    End Function

    '--------------------------------------------------------------------------------------


    'Lining Coding perhitungan --------------------------------------------------------------------------------------------

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text = "Inorganic Lining Type" Then
            ComboBox76.Text = ""
            ComboBox76.Items.Clear()
            ComboBox76.Items.Add("Strip Lined Alloy (Resistant)")
            ComboBox76.Items.Add("Castable Refractory")
            ComboBox76.Items.Add("Castable Refractory Severe Conditions")
            ComboBox76.Items.Add("Glass Lined")
            ComboBox76.Items.Add("Acid Brick")
            ComboBox76.Items.Add("Fiberglass")
        ElseIf ComboBox5.Text = "Organic Lining Type" Then
            ComboBox76.Text = ""
            ComboBox76.Items.Clear()
            ComboBox76.Items.Add("Low Quality Immersion Grade Coating (Spray Applied, to 40 mils)")
            ComboBox76.Items.Add("Medium Quality Immersion Grade Coating (Filled, Trowel Applied, to 80 mils)")
            ComboBox76.Items.Add("High Quality Immersion Grade Coating (Reinforced, Trowel Applied, ≥ 80 mils)")
        End If
    End Sub

    Public Sub liningdf()
        '---dateyear
        Dim d As Double
        d = DateDiff(DateInterval.DayOfYear, DateTimePicker3.Value, DateTimePicker1.Value)

        Dim datetotal As Double
        datetotal = d / 365

        '---basevaluelining
        Dim basevalue As Double

        If Units.ComboBox1.Text = "SI" Then
            Dim chart2(,) As Double = {{1, 0.3, 0.5, 9, 3, 0.01, 1},
            {2, 0.5, 1, 40, 4, 0.03, 1},
            {3, 0.7, 2, 146, 6, 0.05, 1},
            {4, 1, 4, 428, 7, 0.15, 1},
            {5, 1, 9, 1017, 9, 1, 1},
            {6, 2, 16, 1978, 11, 1, 1},
            {7, 3, 30, 3000, 13, 1, 2},
            {8, 4, 53, 3000, 16, 1, 3},
            {9, 6, 89, 3000, 20, 2, 7},
            {10, 9, 146, 3000, 25, 3, 13},
            {11, 12, 230, 3000, 30, 4, 26},
            {12, 16, 351, 3000, 36, 5, 47},
            {13, 22, 518, 3000, 44, 7, 82},
            {14, 30, 738, 3000, 53, 9, 139},
            {15, 40, 1017, 3000, 63, 11, 228},
            {16, 53, 1358, 3000, 75, 15, 359},
            {17, 69, 1758, 3000, 89, 19, 548},
            {18, 89, 2209, 3000, 105, 25, 808},
            {19, 115, 2697, 3000, 124, 31, 1151},
            {20, 146, 3000, 3000, 146, 40, 1587},
            {21, 184, 3000, 3000, 170, 50, 2119},
            {22, 230, 3000, 3000, 199, 63, 2743},
            {23, 286, 3000, 3000, 230, 78, 3000},
            {24, 351, 3000, 3000, 266, 97, 3000},
            {25, 428, 3000, 3000, 306, 119, 3000}}

            Dim UpNum2(6) As Double
            Dim LowNum2(6) As Double

            Dim Lo1 As Double
            Dim Lo2 As Double
            Dim up1 As Double
            Dim up2 As Double

            Dim temp As Double
            Dim upt As Double
            Dim lot As Double
            Dim x As Double
            Dim z As Double


            temp = datetotal

            If ComboBox76.Text = "Strip Lined Alloy (Resistant)" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox76.Text = "Castable Refractory" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 2
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(2)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(2)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox76.Text = "Castable Refractory Severe Conditions" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 3
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(3)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(3)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox76.Text = "Glass Lined" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(4)
                                z = UpNum2(4)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(4)
                                z = UpNum2(4)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 4
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(4)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(4)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox76.Text = "Acid Brick" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(5)
                                z = UpNum2(5)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(5)
                                z = UpNum2(5)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 5
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(5)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(5)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox76.Text = "Fiberglass" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(6)
                                z = UpNum2(6)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(6)
                                z = UpNum2(6)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 6
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(6)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(6)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

        End If

        If Units.ComboBox1.Text = "US" = True Then
            Dim chart2(,) As Double = {{1, 30, 1, 0.1},
            {2, 89, 4, 0.13},
            {3, 230, 16, 0.15},
            {4, 518, 53, 0.17},
            {5, 1017, 146, 0.2},
            {6, 1758, 351, 1},
            {7, 2697, 738, 4},
            {8, 3000, 1358, 16},
            {9, 3000, 2209, 53},
            {10, 3000, 3000, 146},
            {11, 3000, 3000, 351},
            {12, 3000, 3000, 738},
            {13, 3000, 3000, 1358},
            {14, 3000, 3000, 2209},
            {15, 3000, 3000, 3000},
            {16, 3000, 3000, 3000},
            {17, 3000, 3000, 3000},
            {18, 3000, 3000, 3000},
            {19, 3000, 3000, 3000},
            {20, 3000, 3000, 3000},
            {21, 3000, 3000, 3000},
            {22, 3000, 3000, 3000},
            {23, 3000, 3000, 3000},
            {24, 3000, 3000, 3000},
            {25, 3000, 3000, 3000}}

            Dim UpNum2(6) As Double
            Dim LowNum2(6) As Double

            Dim Lo1 As Double
            Dim Lo2 As Double
            Dim up1 As Double
            Dim up2 As Double

            Dim temp As Double
            Dim upt As Double
            Dim lot As Double
            Dim x As Double
            Dim z As Double


            temp = datetotal

            If ComboBox76.Text = "Low Quality Immersion Grade Coating (Spray Applied, to 40 mils)" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox76.Text = "Medium Quality Immersion Grade Coating (Filled, Trowel Applied, to 80 mils)" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 2
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(2)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(2)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox76.Text = "High Quality Immersion Grade Coating (Reinforced, Trowel Applied, ≥ 80 mils)" Then
                If temp < 1 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 2 AndAlso chart2(r + 1, 0) >= 1 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = 2
                                lot = 1

                                basevalue = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 25 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 24 AndAlso chart2(r + 1, 0) >= 25 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = 25
                                lot = 24

                                basevalue = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= temp AndAlso chart2(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 3
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(3)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(3)

                            basevalue = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

        End If

        '---adjlining
        Dim valueadj As Double
        If ComboBox77.Text = "Poor" Then
            valueadj = 10
        ElseIf ComboBox77.Text = "Average" Then
            valueadj = 2
        ElseIf ComboBox77.Text = "Good" Then
            valueadj = 1
        End If

        '---adjonline
        Dim adjonline As Double
        If ComboBox78.Text = "Yes" Then
            adjonline = 0.1
        ElseIf ComboBox78.Text = "No" Then
            adjonline = 1.0
        End If

        '---basedf
        Dim dflining As Double

        dflining = basevalue * valueadj * adjonline
        Label325.Text = dflining

    End Sub


End Class
