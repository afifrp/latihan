﻿Imports MySql.Data.MySqlClient

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

    'Lining DF




    'SCC-Caustic Cracking DF
    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            CheckBox13.Enabled = False
            Label91.Visible = True
            CheckBox14.Visible = True
            CheckBox15.Visible = True
            ComboBox6.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox13.Enabled = True
            Label91.Visible = False
            CheckBox14.Visible = False
            CheckBox15.Visible = False
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            CheckBox12.Enabled = False
            Label92.Visible = True
            CheckBox16.Visible = True
            CheckBox17.Visible = True
        Else
            CheckBox12.Enabled = True
            Label92.Visible = False
            CheckBox16.Visible = False
            CheckBox17.Visible = False
        End If
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        If CheckBox17.Checked = True Then
            Panel41.Visible = True
            CheckBox16.Enabled = False
        Else
            Panel41.Visible = False
            CheckBox16.Enabled = True
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            CheckBox17.Enabled = False
            ComboBox6.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox17.Enabled = True
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        If CheckBox14.Checked = True Then
            CheckBox15.Enabled = False
            ComboBox6.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox15.Enabled = True
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = True Then
            CheckBox14.Enabled = False
            ComboBox6.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox14.Enabled = True
            ComboBox6.Text = ""
        End If
    End Sub

    'SCC Amine Cracking DF
    Private Sub CheckBox27_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox27.CheckedChanged
        If CheckBox27.Checked = True Then
            CheckBox28.Enabled = False
            Label102.Visible = True
            CheckBox29.Visible = True
            CheckBox30.Visible = True
            ComboBox8.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox28.Enabled = True
            Label102.Visible = False
            CheckBox29.Visible = False
            CheckBox30.Visible = False
            ComboBox8.Text = ""
        End If
    End Sub

    Private Sub CheckBox28_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox28.CheckedChanged
        If CheckBox28.Checked = True Then
            CheckBox27.Enabled = False
            Label101.Visible = True
            CheckBox31.Visible = True
            CheckBox32.Visible = True
        Else
            CheckBox27.Enabled = True
            Label101.Visible = False
            CheckBox31.Visible = False
            CheckBox32.Visible = False
        End If
    End Sub

    Private Sub CheckBox29_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox29.CheckedChanged
        If CheckBox29.Checked = True Then
            CheckBox30.Enabled = False
            ComboBox8.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox30.Enabled = True
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox30_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox30.CheckedChanged
        If CheckBox30.Checked = True Then
            CheckBox29.Enabled = False
            ComboBox8.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox29.Enabled = True
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox31_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox31.CheckedChanged
        If CheckBox31.Checked = True Then
            CheckBox32.Enabled = False
            ComboBox8.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox32.Enabled = True
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox32_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox32.CheckedChanged
        If CheckBox32.Checked = True Then
            CheckBox31.Enabled = False

        Else
            CheckBox31.Enabled = True

        End If
    End Sub

    'SCC Sulfide Stress Cracking DF


    Private Sub CheckBox64_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox64.Checked = True Then
            CheckBox65.Enabled = False
            ComboBox10.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox65.Enabled = True
            ComboBox10.Text = ""
        End If
    End Sub

    Private Sub CheckBox65_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox65.Checked = True Then
            CheckBox64.Enabled = False

        Else
            CheckBox64.Enabled = True
        End If
    End Sub

    'SCC-HIC/SOHIC-H2S DF


    Private Sub CheckBox69_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox69.Checked = True Then
            CheckBox70.Enabled = False
            ComboBox11.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox70.Enabled = True
            ComboBox11.Text = ""
        End If
    End Sub

    'Scc Damage Factor – Alkaline Carbonate Stress Corrosion Cracking
    Private Sub CheckBox77_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox77.Checked = True Then
            CheckBox78.Checked = False
            CheckBox79.Checked = False
            CheckBox80.Checked = False
        Else
            CheckBox78.Checked = True
            CheckBox79.Checked = True
            CheckBox80.Checked = True
        End If
    End Sub

    Private Sub CheckBox78_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox78.Checked = True Then
            CheckBox77.Checked = False
            CheckBox79.Checked = False
            CheckBox80.Checked = False
        Else
            CheckBox77.Checked = True
            CheckBox79.Checked = True
            CheckBox80.Checked = True
        End If
    End Sub

    Private Sub CheckBox79_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox79.Checked = True Then
            CheckBox78.Checked = False
            CheckBox77.Checked = False
            CheckBox80.Checked = False
        Else
            CheckBox78.Checked = True
            CheckBox77.Checked = True
            CheckBox80.Checked = True
        End If
    End Sub

    Private Sub CheckBox80_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox80.Checked = True Then
            CheckBox78.Checked = False
            CheckBox79.Checked = False
            CheckBox77.Checked = False
        Else
            CheckBox78.Checked = True
            CheckBox79.Checked = True
            CheckBox77.Checked = True
        End If
    End Sub

    Private Sub CheckBox81_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox81.Checked = True Then
            CheckBox82.Checked = False
        Else
            CheckBox82.Checked = True
        End If
    End Sub

    Private Sub CheckBox82_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox82.Checked = True Then
            CheckBox81.Checked = False
        Else
            CheckBox81.Checked = True
        End If
    End Sub

    Private Sub CheckBox83_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox83.Checked = True Then
            CheckBox84.Checked = False
            Label127.Visible = True
            CheckBox85.Visible = True
            CheckBox86.Visible = True
            ComboBox15.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox84.Checked = True
            Label127.Visible = False
            Label127.Visible = False
            CheckBox85.Visible = False
            CheckBox86.Visible = False
            ComboBox15.Text = ""
        End If
    End Sub

    Private Sub CheckBox84_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox84.Checked = True Then
            CheckBox83.Checked = False
        Else
            CheckBox83.Checked = True
        End If
    End Sub

    'Scc Damage Factor – Polythionic Acid Stress Corrosion Cracking
    Private Sub CheckBox88_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox88.CheckedChanged
        If CheckBox88.Checked = True Then
            CheckBox89.Checked = False
            Label140.Visible = True
            CheckBox90.Visible = True
            CheckBox91.Visible = True
            ComboBox20.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox89.Checked = True
            Label140.Visible = False
            CheckBox90.Visible = False
            CheckBox91.Visible = False
            ComboBox20.Text = ""
        End If
    End Sub

    Private Sub CheckBox89_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox89.CheckedChanged
        If CheckBox89.Checked = True Then
            CheckBox88.Checked = False
            Label141.Visible = True
            CheckBox92.Visible = True
            CheckBox93.Visible = True
            CheckBox94.Visible = True
        Else
            CheckBox88.Checked = True
            Label141.Visible = False
            CheckBox92.Visible = False
            CheckBox93.Visible = False
            CheckBox94.Visible = False
        End If
    End Sub

    Private Sub CheckBox90_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox90.CheckedChanged
        If CheckBox90.Checked = True Then
            CheckBox91.Checked = False

        Else
            CheckBox91.Checked = True

        End If
    End Sub

    Private Sub CheckBox91_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox91.CheckedChanged
        If CheckBox91.Checked = True Then
            CheckBox90.Checked = False
            ComboBox20.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox90.Checked = True
            ComboBox20.Text = ""
        End If
    End Sub

    Private Sub CheckBox92_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox92.CheckedChanged
        If CheckBox92.Checked = True Then
            CheckBox93.Checked = False
            CheckBox94.Checked = False
            Panel48.Show()
        Else
            CheckBox93.Checked = True
            CheckBox94.Checked = True
            Panel48.Hide()
        End If

    End Sub

    Private Sub CheckBox93_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox93.CheckedChanged
        If CheckBox93.Checked = True Then
            CheckBox92.Checked = False
            CheckBox94.Checked = False
            Panel48.Show()
        Else
            CheckBox92.Checked = True
            CheckBox94.Checked = True
            Panel48.Hide()
        End If
    End Sub

    Private Sub CheckBox94_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox94.CheckedChanged
        If CheckBox94.Checked = True Then
            CheckBox92.Checked = False
            CheckBox93.Checked = False
            ComboBox20.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox92.Checked = True
            CheckBox93.Checked = True
            ComboBox20.Text = ""
        End If
    End Sub

    Private Sub CheckBox95_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox95.CheckedChanged
        If CheckBox95.Checked = True Then
            CheckBox96.Checked = False
        Else
            CheckBox96.Checked = True
        End If
    End Sub

    Private Sub CheckBox96_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox96.CheckedChanged
        If CheckBox96.Checked = True Then
            CheckBox95.Checked = False
        Else
            CheckBox95.Checked = True
        End If
    End Sub


    'Scc Damage Factor – Chloride Stress Corrosion Cracking

    Private Sub CheckBox98_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox98.Checked = True Then
            CheckBox99.Checked = False
            CheckBox100.Checked = False
            CheckBox101.Checked = False
        Else
            CheckBox99.Checked = True
            CheckBox100.Checked = True
            CheckBox101.Checked = True
        End If
    End Sub

    Private Sub CheckBox99_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox99.Checked = True Then
            CheckBox98.Checked = False
            CheckBox100.Checked = False
            CheckBox101.Checked = False
        Else
            CheckBox98.Checked = True
            CheckBox100.Checked = True
            CheckBox101.Checked = True
        End If
    End Sub

    Private Sub CheckBox100_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox100.Checked = True Then
            CheckBox98.Checked = False
            CheckBox99.Checked = False
            CheckBox101.Checked = False
        Else
            CheckBox98.Checked = True
            CheckBox99.Checked = True
            CheckBox101.Checked = True
        End If
    End Sub

    Private Sub CheckBox101_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox101.Checked = True Then
            CheckBox98.Checked = False
            CheckBox100.Checked = False
            CheckBox99.Checked = False
        Else
            CheckBox98.Checked = True
            CheckBox100.Checked = True
            CheckBox99.Checked = True
        End If
    End Sub

    Private Sub CheckBox102_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox102.Checked = True Then
            CheckBox103.Checked = False
        Else
            CheckBox103.Checked = True
        End If
    End Sub

    Private Sub CheckBox103_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox103.Checked = True Then
            CheckBox102.Checked = False
        Else
            CheckBox102.Checked = True
        End If
    End Sub

    Private Sub CheckBox104_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox104.Checked = True Then
            CheckBox105.Checked = False
            Label152.Visible = True
            CheckBox106.Visible = True
            CheckBox107.Visible = True
            ComboBox22.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox105.Checked = True
            Label152.Visible = False
            CheckBox106.Visible = False
            CheckBox107.Visible = False
            ComboBox22.Text = ""
        End If
    End Sub

    Private Sub CheckBox105_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox105.Checked = True Then
            CheckBox104.Checked = False
        Else
            CheckBox104.Checked = True
        End If
    End Sub

    Private Sub CheckBox106_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox106.Checked = True Then
            CheckBox107.Checked = False
        Else
            CheckBox107.Checked = True
        End If
    End Sub

    Private Sub CheckBox107_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox107.Checked = True Then
            CheckBox106.Checked = False
            ComboBox22.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox106.Checked = True
            ComboBox22.Text = ""
        End If
    End Sub

    'Scc Damage Factor – Hydrogen Stress Cracking-Hf
    Private Sub CheckBox109_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox109.CheckedChanged
        If CheckBox109.Checked = True Then
            CheckBox110.Checked = False
            Label156.Visible = True
            CheckBox111.Visible = True
            CheckBox112.Visible = True
            ComboBox24.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox110.Checked = True
            Label156.Visible = False
            CheckBox111.Visible = False
            CheckBox112.Visible = False
            ComboBox24.Text = ""
        End If
    End Sub

    Private Sub CheckBox110_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox110.CheckedChanged
        If CheckBox110.Checked = True Then
            CheckBox109.Checked = False
            Label158.Visible = True
            CheckBox113.Visible = True
            CheckBox114.Visible = True

        Else
            CheckBox109.Checked = True
            Label158.Visible = False
            CheckBox113.Visible = False
            CheckBox114.Visible = False

        End If
    End Sub

    Private Sub CheckBox111_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox111.CheckedChanged
        If CheckBox111.Checked = True Then
            CheckBox112.Checked = False
        Else
            CheckBox112.Checked = True
        End If
    End Sub

    Private Sub CheckBox112_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox112.CheckedChanged
        If CheckBox112.Checked = True Then
            CheckBox111.Checked = False
            ComboBox24.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox111.Checked = True
            ComboBox24.Text = ""
        End If
    End Sub

    Private Sub CheckBox113_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox113.CheckedChanged
        If CheckBox113.Checked = True Then
            CheckBox114.Checked = False
            Label159.Visible = True
            CheckBox115.Visible = True
            CheckBox116.Visible = True

        Else
            CheckBox114.Checked = True
            Label159.Visible = False
            CheckBox115.Visible = False
            CheckBox116.Visible = False

        End If
    End Sub

    Private Sub CheckBox114_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox114.CheckedChanged
        If CheckBox114.Checked = True Then
            CheckBox113.Checked = False
            ComboBox24.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox113.Checked = True
            ComboBox24.Text = ""
        End If
    End Sub

    Private Sub CheckBox115_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox115.CheckedChanged
        If CheckBox115.Checked = True Then
            CheckBox116.Checked = False
            Label163.Visible = True
            Label164.Visible = True
            ComboBox101.Visible = True
            ComboBox102.Visible = True
        Else
            CheckBox116.Checked = True
            Label163.Visible = False
            Label164.Visible = False
            ComboBox101.Visible = True
            ComboBox102.Visible = True
        End If
    End Sub

    Private Sub CheckBox116_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox116.CheckedChanged
        If CheckBox116.Checked = True Then
            CheckBox115.Checked = False
            ComboBox24.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox115.Checked = True
            ComboBox24.Text = ""
        End If
    End Sub



    'Scc Damage Factor – Hic/Sohic-Hf
    Private Sub CheckBox123_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox123.CheckedChanged
        If CheckBox123.Checked = True Then
            CheckBox124.Checked = False
            Label165.Visible = True
            CheckBox125.Visible = True
            CheckBox126.Visible = True
            ComboBox27.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox124.Checked = True
            Label165.Visible = False
            CheckBox125.Visible = False
            CheckBox126.Visible = False
            ComboBox27.Text = ""
        End If
    End Sub

    Private Sub CheckBox124_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox124.CheckedChanged
        If CheckBox124.Checked = True Then
            CheckBox123.Checked = False
            Label167.Visible = True
            ComboBox26.Visible = True
        Else
            CheckBox123.Checked = True
            Label167.Visible = False
            ComboBox26.Visible = False
        End If
    End Sub

    Private Sub CheckBox125_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox125.CheckedChanged
        If CheckBox125.Checked = True Then
            CheckBox126.Checked = False

        Else
            CheckBox126.Checked = True

        End If
    End Sub

    Private Sub CheckBox126_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox126.CheckedChanged
        If CheckBox126.Checked = True Then
            CheckBox125.Checked = False
            ComboBox27.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox125.Checked = True
            ComboBox27.Text = ""
        End If
    End Sub

    'External Chloride Stress Corrosion Cracking Damage Factor – Austenitic Component
    Private Sub CheckBox128_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox128.CheckedChanged
        If CheckBox128.Checked = True Then
            CheckBox129.Checked = False
            Label184.Visible = True
            CheckBox132.Visible = True
            CheckBox133.Visible = True
            TextBox39.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox129.Checked = True
            Label184.Visible = False
            CheckBox132.Visible = False
            CheckBox133.Visible = False
            TextBox39.Text = ""
        End If
    End Sub

    Private Sub CheckBox129_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox129.CheckedChanged
        If CheckBox129.Checked = True Then
            CheckBox128.Checked = False
        Else
            CheckBox128.Checked = True
        End If
    End Sub

    Private Sub CheckBox132_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox132.CheckedChanged
        If CheckBox132.Checked = True Then
            CheckBox133.Checked = False
        Else
            CheckBox133.Checked = True
        End If
    End Sub

    Private Sub CheckBox133_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox133.CheckedChanged
        If CheckBox133.Checked = True Then
            CheckBox132.Checked = False
            TextBox39.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox125.Checked = True
            TextBox39.Text = ""
        End If
    End Sub

    'External Chloride Stress Corrosion Cracking Under Insulation Damage Factor – Austenitic Component
    Private Sub CheckBox130_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox130.CheckedChanged
        If CheckBox130.Checked = True Then
            CheckBox131.Checked = False
            Label190.Visible = True
            CheckBox134.Visible = True
            CheckBox135.Visible = True
            TextBox40.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox131.Checked = True
            Label190.Visible = False
            CheckBox134.Visible = False
            CheckBox135.Visible = False
            TextBox40.Text = ""
        End If
    End Sub

    Private Sub CheckBox131_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox131.CheckedChanged
        If CheckBox131.Checked = True Then
            CheckBox130.Checked = False
            Panel51.Show()
        Else
            CheckBox130.Checked = True
            Label190.Visible = False
            Panel51.Hide()
        End If
    End Sub

    Private Sub CheckBox134_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox134.CheckedChanged
        If CheckBox134.Checked = True Then
            CheckBox135.Checked = False

        Else
            CheckBox135.Checked = True


        End If
    End Sub

    Private Sub CheckBox135_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox135.CheckedChanged
        If CheckBox135.Checked = True Then
            CheckBox134.Checked = False
            TextBox40.Text = "PLEASE DETERMINE BY FFS"
        Else
            CheckBox134.Checked = True
            TextBox40.Text = ""
        End If
    End Sub

    Private Sub CheckBox136_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox136.CheckedChanged
        If CheckBox136.Checked = True Then
            CheckBox137.Enabled = True
            CheckBox138.Enabled = True
            CheckBox139.Enabled = True
        Else
            CheckBox137.Enabled = False
            CheckBox138.Enabled = False
            CheckBox139.Enabled = False
        End If
    End Sub

    Private Sub CheckBox137_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox137.CheckedChanged
        If CheckBox137.Checked = True Then
            CheckBox138.Checked = False
            CheckBox139.Checked = False
        Else
            CheckBox138.Checked = True
            CheckBox139.Checked = True

        End If
    End Sub

    Private Sub CheckBox138_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox138.CheckedChanged
        If CheckBox138.Checked = True Then
            CheckBox137.Checked = False
            CheckBox139.Checked = False
        Else
            CheckBox137.Checked = True
            CheckBox139.Checked = True

        End If
    End Sub

    Private Sub CheckBox139_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox139.CheckedChanged
        If CheckBox139.Checked = True Then
            CheckBox137.Checked = False
            CheckBox138.Checked = False
        Else
            CheckBox137.Checked = True
            CheckBox138.Checked = True

        End If
    End Sub

    Private Sub CheckBox140_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox140.CheckedChanged
        If CheckBox140.Checked = True Then
            CheckBox141.Enabled = True
            CheckBox142.Enabled = True
            CheckBox143.Enabled = True
        Else
            CheckBox141.Enabled = False
            CheckBox142.Enabled = False
            CheckBox143.Enabled = False
        End If
    End Sub

    Private Sub CheckBox141_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox141.CheckedChanged
        If CheckBox141.Checked = True Then
            CheckBox142.Checked = False
            CheckBox143.Checked = False
        Else
            CheckBox142.Checked = True
            CheckBox143.Checked = True

        End If
    End Sub

    Private Sub CheckBox142_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox142.CheckedChanged
        If CheckBox142.Checked = True Then
            CheckBox141.Checked = False
            CheckBox143.Checked = False
        Else
            CheckBox141.Checked = True
            CheckBox143.Checked = True

        End If
    End Sub

    Private Sub CheckBox143_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox143.CheckedChanged
        If CheckBox143.Checked = True Then
            CheckBox141.Checked = False
            CheckBox142.Checked = False
        Else
            CheckBox141.Checked = True
            CheckBox142.Checked = True

        End If
    End Sub

    'High Temperature Hydrogen Attack Damage Factor
    Private Sub CheckBox145_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox145.CheckedChanged
        If CheckBox145.Checked = True Then
            CheckBox146.Checked = False

        Else
            CheckBox146.Checked = True


        End If
    End Sub

    Private Sub CheckBox146_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox146.CheckedChanged
        If CheckBox146.Checked = True Then
            CheckBox145.Checked = False

        Else
            CheckBox145.Checked = True


        End If
    End Sub

    Private Sub CheckBox147_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox147.CheckedChanged
        If CheckBox147.Checked = True Then
            CheckBox148.Checked = False

        Else
            CheckBox148.Checked = True


        End If
    End Sub

    Private Sub CheckBox148_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox148.CheckedChanged
        If CheckBox148.Checked = True Then
            CheckBox149.Checked = False

        Else
            CheckBox149.Checked = True


        End If
    End Sub

    Private Sub CheckBox196_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox196.CheckedChanged
        If CheckBox196.Checked = True Then
            TextBox80.Text = FMS.TextBox2.Text
        End If
    End Sub

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


    'Lining
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
