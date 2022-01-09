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
        Call tampildataindustrycombobox()
    End Sub

    'Home ------------------------------------------------------------------------------------

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

        Label391.Text = "Pressure Vessel"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("KODRUM")
        'ComboBox62.Items.Add("DRUM")
        'ComboBox62.Items.Add("REACTOR")
        'ComboBox62.Items.Add("COLTOP")
        'ComboBox62.Items.Add("COLMID")
        'ComboBox62.Items.Add("COLBTM")
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

        Label391.Text = "Pipes and Tubes"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("PIPE-1")
        'ComboBox62.Items.Add("PIPE-3")
        'ComboBox62.Items.Add("PIPE-4")
        'ComboBox62.Items.Add("PIPE-6")
        'ComboBox62.Items.Add("PIPE-8")
        'ComboBox62.Items.Add("PIPE-10")
        'ComboBox62.Items.Add("PIPE-13")
        'ComboBox62.Items.Add("PIPE-16")
        'ComboBox62.Items.Add("PIPEGT16")
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

        Label391.Text = "Atmospheric Storage Tank - Bottom Plates"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("TANKBOTTOM")
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

        Label391.Text = "AirFin Heat Exchanger Header Boxes"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("FINFAN")
        'ComboBox62.Items.Add("FILTER")
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

        Label391.Text = "Compressors"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("COMPC")
        'ComboBox62.Items.Add("COMPR")
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

        Label391.Text = "Pumps"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("PUMP3S")
        'ComboBox62.Items.Add("PUMP1S")
        'ComboBox62.Items.Add("PUMPR")
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

        Label391.Text = "Heat Exchanger"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("HEXSS")
        'ComboBox62.Items.Add("HEXTS")
        'ComboBox62.Items.Add("HEXTUBE")
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

        Label391.Text = "Pressure Relief Devices"
        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("Pressure Relief Devices")
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

        Label391.Text = "Atmospheric Storage Tank - Shell Courses"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("COURSES-1-10")
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

        Label391.Text = "Heat Exchanger Tube Bundles"

        Call tampildatacomponentcombobox()

        'ComboBox62.Items.Clear()
        'ComboBox62.Items.Add("Heat Exchanger Tube Bundles")
    End Sub

    Private Sub ComboBox62_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox62.SelectedIndexChanged
        Call tampildatacodeequipmentcombobox()
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

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
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

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
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

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
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

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
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

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
        Start.Close()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Login.ShowDialog()

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

        df1.Show()
        df2.Show()
        df3.Show()
        df4.Show()
        df5.Show()
        df6.Show()
        df7.Show()
        df8.Show()
        df9.Show()
        df10.Show()
        df11.Show()
        df12.Show()
        df13.Show()
        df14.Show()
        df15.Show()
        df16.Show()
        df17.Show()
        df18.Show()
        df19.Show()
        df20.Show()
        df21.Show()

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

        Call liningdf()
        Call scccausticdf()
        Call sccaminedf()
        Call sccsulfidedf()
        Call scchicsohich2s()
        Call sccalkalinecarbonatedf()
        Call sccpolythionicaciddf()
        Call sccclsccdf()
        Call scchydrogenschfdf()
        Call scchicsohichfdf()


        Call externalsccausteniticdf()
        Call externalscccuiausteniticdf()
        Call brittledf()
        Call lowalloydf()
        Call delapandelapanlimafembrittlementdf()
        Call sigmaphasedf()
        Call hthadf()
        Call mechanicalfatiguedf()
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

    '==========================================================================================================================
    '                                                     BUTTON DF

    'Lining DF --------------------------------------------------------------------------------------------------

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

    Private Sub TextBox35_TextChanged(sender As Object, e As EventArgs) Handles TextBox35.TextChanged
        Call plotarea()
    End Sub

    Private Sub TextBox36_TextChanged(sender As Object, e As EventArgs) Handles TextBox36.TextChanged
        Call plotarea()
    End Sub

    Private Sub ComboBox81_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox81.SelectedIndexChanged
        If TextBox37.Text = "A" Then
            If ComboBox81.Text = "Yes" Then
                ComboBox82.Enabled = True
                ComboBox83.Enabled = True
            ElseIf ComboBox81.Text = "No" Then
                ComboBox82.Enabled = True
                ComboBox83.Enabled = True
            End If
        Else
            If ComboBox81.Text = "Yes" Then
                ComboBox82.Enabled = False
                ComboBox83.Enabled = False
            ElseIf ComboBox81.Text = "No" Then
                ComboBox82.Enabled = False
                ComboBox83.Enabled = False
            End If
        End If

        Call susceptibilitycaustic()
    End Sub

    Private Sub ComboBox82_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox82.SelectedIndexChanged
        Call susceptibilitycaustic()
    End Sub

    Private Sub ComboBox83_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox83.SelectedIndexChanged
        Call susceptibilitycaustic()
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

    Private Sub ComboBox85_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox85.SelectedIndexChanged
        If ComboBox85.Text = "Yes" Then
            Panel46.Visible = True
        ElseIf ComboBox85.Text = "No" Then
            ComboBox108.Text = "NONE SUSCEPTIBILITY"
            Panel46.Visible = False
        Else
            Panel46.Visible = False
        End If
    End Sub

    Private Sub ComboBox86_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox86.SelectedIndexChanged
        If ComboBox86.Text = "Yes" Then
            ComboBox87.Enabled = False
        ElseIf ComboBox86.Text = "No" Then
            ComboBox87.Enabled = True
        End If

        Call susceptibilityamine()
    End Sub

    Private Sub ComboBox87_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox87.SelectedIndexChanged
        If ComboBox87.Text = "Yes" Then
            ComboBox89.Enabled = False
            ComboBox90.Enabled = True
        ElseIf ComboBox87.Text = "No" Then
            ComboBox89.Enabled = False
            ComboBox90.Enabled = False
        End If

        Call susceptibilityamine()
    End Sub

    Private Sub ComboBox88_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox88.SelectedIndexChanged
        Call susceptibilityamine()
    End Sub

    Private Sub ComboBox89_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox89.SelectedIndexChanged
        Call susceptibilityamine()
    End Sub

    Private Sub ComboBox90_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox90.SelectedIndexChanged
        Call susceptibilityamine()
    End Sub

    Private Sub ComboBox91_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox91.SelectedIndexChanged
        Call susceptibilityamine()
    End Sub

    Private Sub ComboBox92_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox92.SelectedIndexChanged
        Call susceptibilityamine()
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

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        Call enviromentalseveritysulfide()
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        Call enviromentalseveritysulfide()
    End Sub

    Private Sub ComboBox80_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox80.SelectedIndexChanged
        Call susceptibilitysulfide()
    End Sub

    Private Sub ComboBox93_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox93.SelectedIndexChanged
        Call susceptibilitysulfide()
    End Sub

    Private Sub TextBox38_TextChanged(sender As Object, e As EventArgs) Handles TextBox38.TextChanged
        Call susceptibilitysulfide()
    End Sub

    Private Sub ComboBox94_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox94.SelectedIndexChanged
        If ComboBox94.Text = "Yes" Then
            Label124.Visible = True
            ComboBox95.Visible = True
            ComboBox110.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox94.Text = "No" Then
            Label124.Visible = False
            ComboBox95.Visible = False
            'ComboBox110.Text = ""
        Else
            Label124.Visible = False
            ComboBox95.Visible = False
            'ComboBox110.Text = ""
        End If
    End Sub

    Private Sub ComboBox95_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox95.SelectedIndexChanged
        If ComboBox95.Text = "Yes" Then
        ElseIf ComboBox95.Text = "No" Then
            ComboBox110.Text = "PLEASE DETERMINE BY FFS"
        Else
        End If
    End Sub

    Private Sub CheckBox73_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox73.CheckedChanged
        If CheckBox73.Checked = True Then
            ComboBox110.Items.Clear()
            ComboBox110.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox110.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox110.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox110.Items.Add("HIGH SUSCEPTIBILITY")
        Else
            ComboBox110.Items.Clear()
            If ComboBox94.Text = "Yes" Then
                ComboBox110.Text = "HIGH SUSCEPTIBILITY"
                If ComboBox95.Text = "Yes" Then
                    Call susceptibilitysulfide()
                ElseIf ComboBox95.Text = "No" Then
                    ComboBox110.Text = "PLEASE DETERMINE BY FFS"
                End If
            ElseIf ComboBox94.Text = "No" Then
                Call susceptibilitysulfide()
            End If
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
            'ComboBox112.Text = ""
        Else
            Label125.Visible = False
            ComboBox17.Visible = False
            'ComboBox112.Text = ""
        End If
    End Sub

    Private Sub ComboBox17_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox17.SelectedIndexChanged
        If ComboBox17.Text = "Yes" Then
        ElseIf ComboBox17.Text = "No" Then
            ComboBox112.Text = "PLEASE DETERMINE BY FFS"
        Else
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

    Private Sub ComboBox14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox14.SelectedIndexChanged
        Call susceptibilitysccalkalinecarbonate()
    End Sub

    Private Sub ComboBox15_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox15.SelectedIndexChanged
        If ComboBox15.Text = "PWHT" Then
            ComboBox98.Enabled = False
            ComboBox98.Text = ""
            Call susceptibilitysccalkalinecarbonate()
        Else
            ComboBox98.Enabled = True
            Call susceptibilitysccalkalinecarbonate()
        End If
    End Sub

    Private Sub ComboBox98_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox98.SelectedIndexChanged
        Call susceptibilitysccalkalinecarbonate()
    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox16.SelectedIndexChanged
        If ComboBox16.Text = "Yes" Then
            Label127.Visible = True
            ComboBox97.Visible = True
            ComboBox114.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox16.Text = "No" Then
            Label127.Visible = False
            ComboBox97.Visible = False
            ComboBox97.Text = ""
            Call susceptibilitysccalkalinecarbonate()
        Else
            Label125.Visible = False
            ComboBox17.Visible = False
            ComboBox97.Text = ""
        End If
    End Sub

    Private Sub ComboBox97_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox97.SelectedIndexChanged
        If ComboBox97.Text = "Yes" Then
            Call susceptibilitysccalkalinecarbonate()
            CheckBox87.Enabled = True
            ComboBox113.Enabled = True
            NumericUpDown5.Enabled = True
        ElseIf ComboBox97.Text = "No" Then
            ComboBox114.Text = "PLEASE DETERMINE BY FFS"
            CheckBox87.Enabled = False
            ComboBox113.Enabled = False
            NumericUpDown5.Enabled = False
        Else
            Call susceptibilitysccalkalinecarbonate()
            CheckBox87.Enabled = True
            ComboBox113.Enabled = True
            NumericUpDown5.Enabled = True
        End If
    End Sub

    Private Sub CheckBox87_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox87.CheckedChanged
        If CheckBox87.Checked = True Then
            ComboBox114.Items.Clear()
            ComboBox114.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox114.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox114.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox114.Items.Add("HIGH SUSCEPTIBILITY")
        Else
            ComboBox114.Items.Clear()
            If ComboBox16.Text = "Yes" Then
                ComboBox114.Text = "HIGH SUSCEPTIBILITY"
                If ComboBox97.Text = "Yes" Then
                    Call susceptibilitysccalkalinecarbonate()
                ElseIf ComboBox97.Text = "No" Then
                    ComboBox114.Text = "PLEASE DETERMINE BY FFS"
                End If
            ElseIf ComboBox16.Text = "No" Then
                Call susceptibilitysccalkalinecarbonate()
            End If
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

    Private Sub ComboBox18_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox18.SelectedIndexChanged
        Call susceptibilitypolythionicacid()
    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        Call susceptibilitypolythionicacid()
    End Sub

    Private Sub ComboBox118_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox118.SelectedIndexChanged
        If ComboBox118.Text = "Yes" Then
            If ComboBox117.Text = "HIGH SUSCEPTIBILITY" Then
                ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf ComboBox117.Text = "MEDIUM SUSCEPTIBILITY" Then
                ComboBox117.Text = "LOW SUSCEPTIBILITY"
            ElseIf ComboBox117.Text = "LOW SUSCEPTIBILITY" Then
                ComboBox117.Text = "NONE SUSCEPTIBILITY"
            End If
        ElseIf ComboBox118.Text = "No" Then

        End If
    End Sub

    Private Sub CheckBox97_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox97.CheckedChanged
        If CheckBox97.Checked = True Then
            ComboBox117.Items.Clear()
            ComboBox117.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox117.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox117.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox117.Items.Add("HIGH SUSCEPTIBILITY")
        Else
            ComboBox117.Items.Clear()
            If ComboBox115.Text = "Yes" Then
                ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                If ComboBox20.Text = "Yes" Then
                    Call susceptibilitypolythionicacid()
                    If ComboBox118.Text = "Yes" Then
                        If ComboBox117.Text = "HIGH SUSCEPTIBILITY" Then
                            ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                        ElseIf ComboBox117.Text = "MEDIUM SUSCEPTIBILITY" Then
                            ComboBox117.Text = "LOW SUSCEPTIBILITY"
                        ElseIf ComboBox117.Text = "LOW SUSCEPTIBILITY" Then
                            ComboBox117.Text = "NONE SUSCEPTIBILITY"
                        End If
                    ElseIf ComboBox118.Text = "No" Then

                    End If
                ElseIf ComboBox20.Text = "No" Then
                    ComboBox117.Text = "PLEASE DETERMINE BY FFS"
                End If
            ElseIf ComboBox115.Text = "No" Then
                Call susceptibilitypolythionicacid()
                If ComboBox118.Text = "Yes" Then
                    If ComboBox117.Text = "HIGH SUSCEPTIBILITY" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox117.Text = "MEDIUM SUSCEPTIBILITY" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox117.Text = "LOW SUSCEPTIBILITY" Then
                        ComboBox117.Text = "NONE SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox118.Text = "No" Then

                End If
            End If
        End If
    End Sub

    'Scc Damage Factor – Chloride Stress Corrosion Cracking ----------------------------------------------------

    Private Sub ComboBox23_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox23.SelectedIndexChanged
        Call susceptibilitysccchloride()
    End Sub

    Private Sub ComboBox22_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox22.SelectedIndexChanged
        Call susceptibilitysccchloride()
    End Sub

    Private Sub ComboBox99_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox99.SelectedIndexChanged
        If ComboBox99.Text = "Yes" Then
            Label152.Visible = True
            ComboBox100.Visible = True
            ComboBox120.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox99.Text = "No" Then
            Label152.Visible = False
            ComboBox100.Visible = False
            ComboBox100.Text = ""
            Call susceptibilitysccchloride()
        Else
            Label152.Visible = False
            ComboBox100.Visible = False
            ComboBox100.Text = ""
            Call susceptibilitysccchloride()
        End If
    End Sub

    Private Sub ComboBox100_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox100.SelectedIndexChanged
        If ComboBox100.Text = "Yes" Then
            Call susceptibilitysccchloride()
            CheckBox108.Enabled = True
            ComboBox119.Enabled = True
            NumericUpDown7.Enabled = True
        ElseIf ComboBox100.Text = "No" Then
            ComboBox120.Text = "PLEASE DETERMINE BY FFS"
            CheckBox108.Enabled = False
            ComboBox119.Enabled = False
            NumericUpDown7.Enabled = False
        Else
            Call susceptibilitysccchloride()
            CheckBox108.Enabled = True
            ComboBox119.Enabled = True
            NumericUpDown7.Enabled = True
        End If
    End Sub

    Private Sub CheckBox108_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox108.CheckedChanged
        If CheckBox108.Checked = True Then
            ComboBox120.Items.Clear()
            ComboBox120.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox120.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox120.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox120.Items.Add("HIGH SUSCEPTIBILITY")
        Else
            ComboBox120.Items.Clear()
            If ComboBox99.Text = "Yes" Then
                ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                If ComboBox100.Text = "Yes" Then
                    Call susceptibilitysccchloride()
                ElseIf ComboBox100.Text = "No" Then
                    ComboBox120.Text = "PLEASE DETERMINE BY FFS"
                End If
            ElseIf ComboBox99.Text = "No" Then
                Call susceptibilitysccchloride()
            End If
        End If

    End Sub

    'Scc Damage Factor – Hydrogen Stress Cracking-Hf -----------------------------------------------------------

    Private Sub ComboBox24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox24.SelectedIndexChanged
        If ComboBox24.Text = "Yes" Then
            Label156.Visible = True
            Label158.Visible = False
            ComboBox121.Visible = True
            ComboBox25.Visible = False
            Label159.Visible = False
            ComboBox122.Visible = False
            ComboBox124.Text = "HIGH SUSCEPTIBILITY"
            ComboBox25.Text = ""
            ComboBox122.Text = ""
            CheckBox122.Enabled = True
            ComboBox123.Enabled = True
            NumericUpDown8.Enabled = True
            Panel49.Visible = False
            ComboBox101.Text = ""
            ComboBox102.Text = ""
        ElseIf ComboBox24.Text = "No" Then
            Label156.Visible = False
            Label158.Visible = True
            ComboBox121.Visible = False
            ComboBox25.Visible = True
            ComboBox124.Text = ""
            ComboBox121.Text = ""
            CheckBox122.Enabled = True
            ComboBox123.Enabled = True
            NumericUpDown8.Enabled = True
            Panel49.Visible = False
            ComboBox101.Text = ""
            ComboBox102.Text = ""
        Else
            Label156.Visible = False
            Label158.Visible = False
            ComboBox121.Visible = False
            ComboBox25.Visible = False
            ComboBox124.Text = ""
            ComboBox25.Text = ""
            ComboBox121.Text = ""
            ComboBox122.Text = ""
            CheckBox122.Enabled = True
            ComboBox123.Enabled = True
            NumericUpDown8.Enabled = True
            Panel49.Visible = False
            ComboBox101.Text = ""
            ComboBox102.Text = ""
        End If
    End Sub

    Private Sub ComboBox121_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox121.SelectedIndexChanged
        If ComboBox121.Text = "Yes" Then
            ComboBox124.Text = "NONE SUSCEPTIBILITY"
            CheckBox122.Enabled = True
            ComboBox123.Enabled = True
            NumericUpDown8.Enabled = True
        ElseIf ComboBox121.Text = "No" Then
            ComboBox124.Text = "PLEASE DETERMINE BY FFS"
            CheckBox122.Enabled = False
            ComboBox123.Enabled = False
            NumericUpDown8.Enabled = False
        Else
            CheckBox122.Enabled = True
            ComboBox123.Enabled = True
            NumericUpDown8.Enabled = True
        End If
    End Sub

    Private Sub ComboBox25_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox25.SelectedIndexChanged
        If ComboBox25.Text = "Yes" Then
            Label159.Visible = True
            ComboBox122.Visible = True
            ComboBox124.Text = ""
        ElseIf ComboBox25.Text = "No" Then
            Label159.Visible = False
            ComboBox122.Visible = False
            Panel49.Visible = False
            ComboBox124.Text = "NONE SUSCEPTIBILITY"
            ComboBox122.Text = ""
            ComboBox101.Text = ""
            ComboBox102.Text = ""
        Else
            Label152.Visible = False
            ComboBox100.Visible = False
            Panel49.Visible = False
            ComboBox124.Text = ""
            ComboBox122.Text = ""
            ComboBox101.Text = ""
            ComboBox102.Text = ""
        End If
    End Sub

    Private Sub ComboBox122_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox122.SelectedIndexChanged
        If ComboBox122.Text = "Yes" Then
            Panel49.Visible = True
            ComboBox124.Text = ""
        ElseIf ComboBox122.Text = "No" Then
            Panel49.Visible = False
            ComboBox124.Text = "NONE SUSCEPTIBILITY"
            ComboBox101.Text = ""
            ComboBox102.Text = ""
        Else
            Panel49.Visible = False
            ComboBox124.Text = ""
        End If
    End Sub

    Private Sub ComboBox101_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox101.SelectedIndexChanged
        Call susceptibilityscchydrogenschf()
    End Sub

    Private Sub ComboBox102_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox102.SelectedIndexChanged
        Call susceptibilityscchydrogenschf()
    End Sub

    Private Sub CheckBox122_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox122.CheckedChanged
        If CheckBox122.Checked = True Then
            ComboBox124.Items.Clear()
            ComboBox124.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox124.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox124.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox124.Items.Add("HIGH SUSCEPTIBILITY")
        Else
            ComboBox124.Items.Clear()
            If ComboBox24.Text = "Yes" Then
                ComboBox124.Text = "HIGH SUSCEPTIBILITY"
                If ComboBox121.Text = "Yes" Then
                    Call susceptibilityscchydrogenschf()
                ElseIf ComboBox121.Text = "No" Then
                    ComboBox124.Text = "PLEASE DETERMINE BY FFS"
                End If
            ElseIf ComboBox24.Text = "No" Then
                If ComboBox25.Text = "Yes" Then
                    If ComboBox122.Text = "Yes" Then
                        Call susceptibilityscchydrogenschf()
                    ElseIf ComboBox122.Text = "No" Then
                        ComboBox124.Text = "NONE SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox25.Text = "No" Then
                    ComboBox124.Text = "NONE SUSCEPTIBILITY"
                End If
            End If
        End If
    End Sub

    'Scc Damage Factor – Hic/Sohic-Hf --------------------------------------------------------------------------

    Private Sub ComboBox26_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox26.SelectedIndexChanged
        If ComboBox26.Text = "Yes" Then
            Label165.Visible = True
            Label167.Visible = False
            ComboBox27.Visible = True
            ComboBox126.Visible = False
            Label388.Visible = False
            ComboBox125.Visible = False
            ComboBox165.Text = "HIGH SUSCEPTIBILITY"
            ComboBox126.Text = ""
            ComboBox125.Text = ""
            CheckBox127.Enabled = True
            ComboBox164.Enabled = True
            ComboBox168.Enabled = True
            NumericUpDown9.Enabled = True
            Panel97.Visible = False
            ComboBox28.Text = ""
            ComboBox163.Text = ""
        ElseIf ComboBox26.Text = "No" Then
            Label165.Visible = False
            Label167.Visible = True
            ComboBox27.Visible = False
            ComboBox126.Visible = True
            ComboBox165.Text = ""
            ComboBox27.Text = ""
            CheckBox127.Enabled = True
            ComboBox164.Enabled = True
            ComboBox168.Enabled = True
            NumericUpDown9.Enabled = True
            Panel97.Visible = False
            ComboBox28.Text = ""
            ComboBox163.Text = ""
        Else
            Label140.Visible = False
            Label141.Visible = False
            ComboBox20.Visible = False
            ComboBox21.Visible = False
            ComboBox165.Text = ""
            ComboBox126.Text = ""
            ComboBox27.Text = ""
            ComboBox125.Text = ""
            CheckBox127.Enabled = True
            ComboBox164.Enabled = True
            ComboBox168.Enabled = True
            NumericUpDown9.Enabled = True
            Panel97.Visible = False
            ComboBox28.Text = ""
            ComboBox163.Text = ""
        End If
    End Sub

    Private Sub ComboBox27_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox27.SelectedIndexChanged
        If ComboBox27.Text = "Yes" Then
            ComboBox165.Text = "NONE SUSCEPTIBILITY"
            CheckBox127.Enabled = True
            ComboBox164.Enabled = True
            ComboBox168.Enabled = True
            NumericUpDown9.Enabled = True
        ElseIf ComboBox27.Text = "No" Then
            ComboBox165.Text = "PLEASE DETERMINE BY FFS"
            CheckBox127.Enabled = False
            ComboBox164.Enabled = False
            ComboBox168.Enabled = False
            NumericUpDown9.Enabled = False
        Else
            CheckBox127.Enabled = True
            ComboBox164.Enabled = True
            ComboBox168.Enabled = True
            NumericUpDown9.Enabled = True
        End If
    End Sub

    Private Sub ComboBox126_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox126.SelectedIndexChanged
        If ComboBox126.Text = "Yes" Then
            Label388.Visible = True
            ComboBox125.Visible = True
            ComboBox165.Text = ""
        ElseIf ComboBox126.Text = "No" Then
            Panel97.Visible = False
            Label388.Visible = False
            ComboBox125.Visible = False
            Panel97.Visible = False
            ComboBox165.Text = "NONE SUSCEPTIBILITY"
            ComboBox125.Text = ""
            ComboBox28.Text = ""
            ComboBox163.Text = ""
        Else
            Panel97.Visible = False
            ComboBox165.Text = ""
            ComboBox125.Text = ""
            ComboBox28.Text = ""
            ComboBox163.Text = ""
        End If
    End Sub

    Private Sub ComboBox125_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox125.SelectedIndexChanged
        If ComboBox125.Text = "Yes" Then
            Panel97.Visible = True
            ComboBox165.Text = ""
        ElseIf ComboBox125.Text = "No" Then
            Panel97.Visible = False
            ComboBox165.Text = "NONE SUSCEPTIBILITY"
            ComboBox28.Text = ""
            ComboBox163.Text = ""
        Else
            Panel97.Visible = False
            ComboBox165.Text = ""
        End If
    End Sub

    Private Sub ComboBox28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox28.SelectedIndexChanged
        Call susceptibilityscchicsohichf()
    End Sub

    Private Sub ComboBox163_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox163.SelectedIndexChanged
        Call susceptibilityscchicsohichf()
    End Sub

    Private Sub CheckBox127_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox127.CheckedChanged
        If CheckBox127.Checked = True Then
            ComboBox165.Items.Clear()
            ComboBox165.Items.Add("NONE SUSCEPTIBILITY")
            ComboBox165.Items.Add("LOW SUSCEPTIBILITY")
            ComboBox165.Items.Add("MEDIUM SUSCEPTIBILITY")
            ComboBox165.Items.Add("HIGH SUSCEPTIBILITY")
        Else
            ComboBox165.Items.Clear()
            If ComboBox26.Text = "Yes" Then
                ComboBox165.Text = "HIGH SUSCEPTIBILITY"
                If ComboBox27.Text = "Yes" Then
                    Call susceptibilityscchicsohichf()
                ElseIf ComboBox27.Text = "No" Then
                    ComboBox165.Text = "PLEASE DETERMINE BY FFS"
                End If
            ElseIf ComboBox26.Text = "No" Then
                If ComboBox126.Text = "Yes" Then
                    If ComboBox125.Text = "Yes" Then
                        Call susceptibilityscchicsohichf()
                    ElseIf ComboBox125.Text = "No" Then
                        ComboBox165.Text = "NONE SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox126.Text = "No" Then
                    ComboBox165.Text = "NONE SUSCEPTIBILITY"
                End If
            End If
        End If
    End Sub


    'External Corrosion Damage Factor – Ferritic Component -----------------------------------------------------

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Label334.Visible = True
            TextBox85.Visible = True
        Else
            Label334.Visible = False
            TextBox85.Visible = False
        End If
    End Sub

    'Corrosion Under Insulation Damage Factor – Ferritic Component ---------------------------------------------

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Label347.Visible = True
            TextBox88.Visible = True
        Else
            Label347.Visible = False
            TextBox88.Visible = False
        End If
    End Sub

    'External Chloride Stress Corrosion Cracking Damage Factor – Austenitic Component --------------------------

    Private Sub ComboBox29_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox29.SelectedIndexChanged
        Call susceptibilityexternalsccaustenitic()
    End Sub

    Private Sub ComboBox30_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox30.SelectedIndexChanged
        If ComboBox30.Text = "Yes" Then
            Label184.Visible = True
            ComboBox31.Visible = True
            TextBox39.Text = "HIGH SUSCEPTIBILITY"
        ElseIf ComboBox30.Text = "No" Then
            Label184.Visible = False
            ComboBox31.Visible = False
            ComboBox31.Text = ""
            Call susceptibilityexternalsccaustenitic()
        Else
            Label184.Visible = False
            ComboBox31.Visible = False
            ComboBox31.Text = ""
            Call susceptibilityexternalsccaustenitic()
        End If
    End Sub

    Private Sub ComboBox31_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox31.SelectedIndexChanged
        If ComboBox31.Text = "Yes" Then
            Call susceptibilityexternalsccaustenitic()
            DateTimePicker4.Enabled = True
            ComboBox128.Enabled = True
            ComboBox127.Enabled = True
            NumericUpDown10.Enabled = True
        ElseIf ComboBox100.Text = "No" Then
            ComboBox120.Text = "PLEASE DETERMINE BY FFS"
            DateTimePicker4.Enabled = False
            ComboBox108.Enabled = False
            ComboBox127.Enabled = False
            NumericUpDown10.Enabled = False
        Else
            Call susceptibilityexternalsccaustenitic()
            DateTimePicker4.Enabled = True
            CheckBox108.Enabled = True
            ComboBox119.Enabled = True
            NumericUpDown7.Enabled = True
        End If
    End Sub

    Private Sub CheckBox136_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox136.CheckedChanged
        If CheckBox136.Checked = True Then
            ComboBox131.Enabled = True
        Else
            ComboBox131.Enabled = False
            If ComboBox132.Text IsNot "" Then
                Call susceptibilityexternalscccuiaustenitic()
                Call insulationcondition()
            Else
                Call susceptibilityexternalscccuiaustenitic()
            End If
            ComboBox131.Text = ""
            If CheckBox140.Checked = False AndAlso CheckBox144.Checked = False Then
                Call susceptibilityexternalscccuiaustenitic()
            End If
        End If
    End Sub

    Private Sub CheckBox140_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox140.CheckedChanged
        If CheckBox140.Checked = True Then
            ComboBox132.Enabled = True
        Else
            ComboBox132.Enabled = False
            If ComboBox131.Text IsNot "" Then
                Call susceptibilityexternalscccuiaustenitic()
                Call pipingcomplexity()
            Else
                Call susceptibilityexternalscccuiaustenitic()
            End If
            ComboBox132.Text = ""
            If CheckBox136.Checked = False AndAlso CheckBox144.Checked = False Then
                Call susceptibilityexternalscccuiaustenitic()
            End If
        End If
    End Sub

    Private Sub CheckBox144_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox144.CheckedChanged
        If CheckBox144.Checked = True Then
            If TextBox40.Text = "LOW SUSCEPTIBILITY" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "MEDIUM SUSCEPTIBILITY" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "HIGH SUSCEPTIBILITY" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"
            End If
        Else
            If ComboBox132.Text IsNot "" AndAlso ComboBox131.Text IsNot "" Then
                Call susceptibilityexternalscccuiaustenitic()
                Call pipingcomplexity()
                Call insulationcondition()
            ElseIf ComboBox132.Text IsNot "" Then
                Call susceptibilityexternalscccuiaustenitic()
                Call insulationcondition()
            ElseIf ComboBox131.Text IsNot "" Then
                Call susceptibilityexternalscccuiaustenitic()
                Call pipingcomplexity()
            Else
                Call susceptibilityexternalscccuiaustenitic()
            End If

            If CheckBox136.Checked = False AndAlso CheckBox140.Checked = False Then
                Call susceptibilityexternalscccuiaustenitic()
            End If

        End If
    End Sub

    Private Sub ComboBox131_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox131.SelectedIndexChanged
        If CheckBox136.Checked = True Then
            Call pipingcomplexity()
        Else
            If ComboBox132.Enabled = True Then

            Else
                Call susceptibilityexternalscccuiaustenitic()
            End If
        End If

    End Sub

    Private Sub ComboBox132_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox132.SelectedIndexChanged
        If CheckBox140.Checked = True Then
            Call insulationcondition()
        Else
            If ComboBox131.Enabled = True Then

            Else
                Call susceptibilityexternalscccuiaustenitic()
            End If
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

    Private Sub TextBox39_TextChanged(sender As Object, e As EventArgs) Handles TextBox39.TextChanged
        TextBox40.Text = TextBox39.Text
    End Sub

    'High Temperature Hydrogen Attack Damage Factor ------------------------------------------------------------

    Private Sub ComboBox133_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox133.SelectedIndexChanged
        If ComboBox133.Text = "Yes" Then
            Label192.Visible = True
            ComboBox134.Visible = True
        Else
            Label192.Visible = False
            ComboBox134.Visible = False
        End If
    End Sub

    Private Sub ComboBox134_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox134.SelectedIndexChanged
        If ComboBox134.Text = "Yes, replace in kind" Then
            TextBox45.Text = "High Susceptibility"
            TextBox41.ReadOnly = True
            TextBox42.ReadOnly = True
        ElseIf ComboBox134.Text = "No, not replaced" Then
            TextBox45.Text = "Damage Observed"
            TextBox41.ReadOnly = True
            TextBox42.ReadOnly = True
        ElseIf ComboBox134.Text = "No, replaced with upgraded material" Then
            TextBox45.Text = ""
            TextBox41.ReadOnly = False
            TextBox42.ReadOnly = False
        Else
            TextBox45.Text = ""
            TextBox41.ReadOnly = False
            TextBox42.ReadOnly = False
        End If
    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged
        If ComboBox1.Text = "Cr-Mo Low Alloy Steel" Then
            If Units.ComboBox1.Text = "SI" Then
                If Val(TextBox41.Text) > 177 And Val(TextBox42.Text) > 0.345 Then
                    Panel98.Visible = True
                Else
                    Panel98.Visible = False
                End If
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                If Val(TextBox41.Text) > 350 And Val(TextBox42.Text) > 50 Then
                    Panel98.Visible = True
                Else
                    Panel98.Visible = False
                End If
            End If

        ElseIf ComboBox1.Text = "Carbon Steel" Then
            If Units.ComboBox1.Text = "SI" Then
                If Val(TextBox41.Text) > 177 And Val(TextBox42.Text) > 0.345 Then
                    TextBox45.Text = "High Susceptibility"
                Else
                    TextBox45.Text = "No Susceptibility"
                End If
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                If Val(TextBox41.Text) > 350 And Val(TextBox42.Text) > 50 Then
                    TextBox45.Text = "High Susceptibility"
                Else
                    TextBox45.Text = "No Susceptibility"
                End If
            End If

        ElseIf ComboBox1.Text = "C-1/2 Mo Alloy Steel" Then
            If Units.ComboBox1.Text = "SI" Then
                If Val(TextBox41.Text) > 177 And Val(TextBox42.Text) > 0.345 Then
                    TextBox45.Text = "High Susceptibility"
                Else
                    TextBox45.Text = "No Susceptibility"
                End If
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                If Val(TextBox41.Text) > 350 And Val(TextBox42.Text) > 50 Then
                    TextBox45.Text = "High Susceptibility"
                Else
                    TextBox45.Text = "No Susceptibility"
                End If
            End If
        End If
    End Sub

    Private Sub TextBox42_TextChanged(sender As Object, e As EventArgs) Handles TextBox42.TextChanged
        If ComboBox1.Text = "Cr-Mo Low Alloy Steel" Then
            If Units.ComboBox1.Text = "SI" Then
                If Val(TextBox41.Text) > 177 And Val(TextBox42.Text) > 0.345 Then
                    Panel98.Visible = True
                Else
                    Panel98.Visible = False
                End If
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                If Val(TextBox41.Text) > 350 And Val(TextBox42.Text) > 50 Then
                    Panel98.Visible = True
                Else
                    Panel98.Visible = False
                End If
            End If

        ElseIf textbox45.Text = "Carbon Steel" Then
            If Units.ComboBox1.Text = "SI" Then
                If Val(TextBox41.Text) > 177 And Val(TextBox42.Text) > 0.345 Then
                    TextBox45.Text = "High Susceptibility"
                Else
                    TextBox45.Text = "No Susceptibility"
                End If
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                If Val(TextBox41.Text) > 350 And Val(TextBox42.Text) > 50 Then
                    textbox45.Text = "High Susceptibility"
                Else
                    textbox45.Text = "No Susceptibility"
                End If
            End If

        ElseIf textbox45.Text = "C-1/2 Mo Alloy Steel" Then
            If Units.ComboBox1.Text = "SI" Then
                If Val(TextBox41.Text) > 177 And Val(TextBox42.Text) > 0.345 Then
                    TextBox45.Text = "High Susceptibility"
                Else
                    TextBox45.Text = "No Susceptibility"
                End If
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                If Val(TextBox41.Text) > 350 And Val(TextBox42.Text) > 50 Then
                    textbox45.Text = "High Susceptibility"
                Else
                    textbox45.Text = "No Susceptibility"
                End If
            End If
        End If
    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click
        Dim exposure As Double
        Dim partialpressure As Double
        Dim temponcurve As Double
        Dim deltatempproximity As Double

        If Units.ComboBox1.Text = "SI" Then
            exposure = Val(TextBox41.Text) * (9 / 5) + 32
            partialpressure = Val(TextBox42.Text) * 0.0980665 * 145.038
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            exposure = Val(TextBox41.Text)
            partialpressure = Val(TextBox42.Text)
        End If

        If ComboBox134.Text = "6.0Cr-0.5Mo Steel" Then
            Dim a As Double
            a = -0.1714 * partialpressure + 1300

            If partialpressure <= 1050 Then
                temponcurve = a
            ElseIf partialpressure >= 1050 Then
                temponcurve = 1120.03
            End If
        End If

        If ComboBox134.Text = "3.0Cr-1.0Mo Steel" Then
            Dim b As Double
            b = -0.1648 * partialpressure + 1250

            If partialpressure <= 1820 Then
                temponcurve = b
            ElseIf partialpressure >= 1820 Then
                temponcurve = 950.064
            End If
        End If

        If ComboBox134.Text = "2.25Cr-1.0Mo-V Steel" Then
            Dim c As Double
            c = -0.1975 * partialpressure + 1200

            If partialpressure <= 2000 Then
                temponcurve = c
            ElseIf partialpressure >= 2000 Then
                temponcurve = 805
            End If
        End If

        If ComboBox134.Text = "2.25Cr-1.0Mo Steel" Then
            Dim c As Double
            c = -0.1975 * partialpressure + 1200

            If partialpressure <= 2000 Then
                temponcurve = c
            ElseIf partialpressure >= 2000 Then
                temponcurve = 805
            End If
        End If

        If ComboBox134.Text = "1.25Cr-0.5Mo Steel" Then
            Dim d As Double
            d = -0.1646 * partialpressure + 1150

            If partialpressure <= 1225 Then
                temponcurve = d
            End If

            Dim g As Double
            g = -3415 * Math.Log(partialpressure) + 25236

            If partialpressure >= 1225 And partialpressure <= 1280 Then
                temponcurve = g
            End If

            Dim h As Double
            h = -836.9 * Math.Log(partialpressure) + 6788

            If partialpressure >= 1280 And partialpressure <= 1400 Then
                temponcurve = h
            End If

            Dim i As Double
            i = -374.4 * Math.Log(partialpressure) + 3437.6

            If partialpressure >= 1400 And partialpressure <= 1600 Then
                temponcurve = i
            End If

            Dim j As Double
            j = -145.5 * Math.Log(partialpressure) + 1748.3

            If partialpressure >= 1600 And partialpressure <= 1900 Then
                temponcurve = j
            End If

            Dim k As Double
            k = -71.14 * Math.Log(partialpressure) + 1187.1

            If partialpressure >= 1900 And partialpressure <= 2700 Then
                temponcurve = k
            ElseIf partialpressure >= 2700 Then
                temponcurve = 625
            End If
        End If

        If ComboBox134.Text = "1.0Cr-0.5Mo Steel" Then
            Dim f As Double
            f = -0.2941 * partialpressure + 1100

            If partialpressure <= 680 Then
                temponcurve = f
            ElseIf partialpressure >= 680 Then
                temponcurve = 900
            End If

            Dim l As Double
            l = -2277 * Math.Log(partialpressure) + 17090

            If partialpressure >= 1225 And partialpressure <= 1280 Then
                temponcurve = l
            End If

            Dim h As Double
            h = -836.9 * Math.Log(partialpressure) + 6788

            If partialpressure >= 1280 And partialpressure <= 1400 Then
                temponcurve = h
            End If

            Dim i As Double
            i = -374.4 * Math.Log(partialpressure) + 3437.6

            If partialpressure >= 1400 And partialpressure <= 1600 Then
                temponcurve = i
            End If

            Dim j As Double
            j = -145.5 * Math.Log(partialpressure) + 1748.3

            If partialpressure >= 1600 And partialpressure <= 1900 Then
                temponcurve = j
            End If

            Dim k As Double
            k = -71.14 * Math.Log(partialpressure) + 1187.1

            If partialpressure >= 1900 And partialpressure <= 2700 Then
                temponcurve = k
            ElseIf partialpressure >= 2700 And partialpressure <= 13000 Then
                temponcurve = 625
            End If
        End If

        TextBox43.Text = temponcurve

        deltatempproximity = temponcurve - exposure

        TextBox44.Text = deltatempproximity

        If ComboBox1.Text = "Cr-Mo Low Alloy Steel" Then
            If deltatempproximity <= 0 Then
                TextBox45.Text = "High Susceptibility"
            ElseIf deltatempproximity >= 0 And deltatempproximity <= 50 Then
                TextBox45.Text = "Medium Susceptibility"
            ElseIf deltatempproximity >= 50 And deltatempproximity <= 100 Then
                TextBox45.Text = "Low Susceptibility"
            ElseIf deltatempproximity >= 100 Then
                TextBox45.Text = "No Susceptibility"
            End If
        End If

    End Sub

    'Brittle Fracture Damage Factor ----------------------------------------------------------------------------


    'Low Alloy Steel Embrittlement Damage Factor ---------------------------------------------------------------
    Private Sub ComboBox160_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox160.SelectedIndexChanged
        If ComboBox160.Text = "Yes" Then
            TextBox91.ReadOnly = False
            ComboBox161.Visible = False
            ComboBox161.Text = ""
            TextBox46.Text = ""
        ElseIf ComboBox160.Text = "No" Then
            TextBox91.ReadOnly = True
            ComboBox161.Visible = True
        Else
            TextBox91.ReadOnly = True
        End If
    End Sub

    Private Sub ComboBox161_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox161.SelectedIndexChanged
        If ComboBox161.Text = "Minimum Design Temperature" Then
            TextBox91.ReadOnly = False
            TextBox91.Text = ""
        ElseIf ComboBox169.Text = "Minimum Pressurization Temperature" Then
            TextBox91.ReadOnly = False
            TextBox91.Text = ""
        Else
            TextBox91.ReadOnly = True
        End If
    End Sub

    Private Sub ComboBox162_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox162.SelectedIndexChanged
        If ComboBox162.Text = "Yes" Then
            TextBox91.ReadOnly = False
            ComboBox171.Visible = False
            Panel95.Visible = False
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
            ComboBox171.Text = ""
            TextBox103.Text = ""
        ElseIf ComboBox162.Text = "No" Then
            TextBox103.ReadOnly = True
            ComboBox171.Visible = True
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
        Else
            TextBox91.ReadOnly = True
            ComboBox171.Visible = True
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
        End If
    End Sub

    Private Sub ComboBox171_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox171.SelectedIndexChanged
        If ComboBox171.Text = "Engineering analysis or Impact Test" Then
            TextBox91.ReadOnly = True
            Panel95.Visible = True
            FATT1.Visible = True
            FATT2.Visible = False
            FATT3.Visible = False
            FATT4.Visible = False
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
            ComboBox171.Text = ""
            TextBox103.Text = ""
        ElseIf ComboBox171.Text = "Step Cooling Embrittlement (SCE) test" Then
            TextBox91.ReadOnly = True
            Panel95.Visible = True
            FATT1.Visible = False
            FATT2.Visible = True
            FATT3.Visible = False
            FATT4.Visible = False
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
            ComboBox171.Text = ""
            TextBox103.Text = ""
        ElseIf ComboBox171.Text = "J-factor correlation, X-bar factor" Then
            TextBox91.ReadOnly = True
            Panel95.Visible = True
            FATT1.Visible = False
            FATT2.Visible = False
            FATT3.Visible = True
            FATT4.Visible = False
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
            ComboBox171.Text = ""
            TextBox103.Text = ""
        ElseIf ComboBox171.Text = "Conservative Assumption based on year of fabrication" Then
            TextBox91.ReadOnly = True
            Panel95.Visible = True
            FATT1.Visible = False
            FATT2.Visible = False
            FATT3.Visible = False
            FATT4.Visible = True
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
            ComboBox171.Text = ""
            TextBox103.Text = ""
        Else
            Panel95.Visible = False
            FATT1.Visible = False
            FATT2.Visible = False
            FATT3.Visible = False
            FATT4.Visible = False
            TextBox91.ReadOnly = True
            TextBox92.Text = ""
            TextBox93.Text = ""
            TextBox94.Text = ""
            TextBox95.Text = ""
            TextBox96.Text = ""
            TextBox97.Text = ""
            TextBox98.Text = ""
            TextBox99.Text = ""
            TextBox100.Text = ""
            TextBox101.Text = ""
            TextBox102.Text = ""
            ComboBox172.Text = ""
        End If
    End Sub

    Private Sub TextBox92_TextChanged(sender As Object, e As EventArgs) Handles TextBox92.TextChanged
        Call fatt()
    End Sub

    Private Sub TextBox93_TextChanged(sender As Object, e As EventArgs) Handles TextBox93.TextChanged
        If TextBox93.Text IsNot "" AndAlso TextBox94.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If
    End Sub

    Private Sub TextBox94_TextChanged(sender As Object, e As EventArgs) Handles TextBox94.TextChanged
        If TextBox93.Text IsNot "" AndAlso TextBox94.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If
    End Sub

    Private Sub TextBox95_TextChanged(sender As Object, e As EventArgs) Handles TextBox95.TextChanged
        If TextBox95.Text = "" AndAlso TextBox96.Text = "" AndAlso TextBox97.Text = "" AndAlso TextBox98.Text = "" Then
            Label402.Enabled = True
            Label403.Enabled = True
            Label408.Enabled = True
            Label409.Enabled = True
            Label410.Enabled = True
            Label411.Enabled = True
            TextBox99.Enabled = True
            TextBox100.Enabled = True
            TextBox101.Enabled = True
            TextBox102.Enabled = True
        Else
            Label402.Enabled = False
            Label403.Enabled = False
            Label408.Enabled = False
            Label409.Enabled = False
            Label410.Enabled = False
            Label411.Enabled = False
            TextBox99.Enabled = False
            TextBox100.Enabled = False
            TextBox101.Enabled = False
            TextBox102.Enabled = False
        End If

        If TextBox95.Text IsNot "" AndAlso TextBox96.Text IsNot "" AndAlso TextBox97.Text IsNot "" AndAlso TextBox98.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub TextBox96_TextChanged(sender As Object, e As EventArgs) Handles TextBox96.TextChanged
        If TextBox95.Text = "" AndAlso TextBox96.Text = "" AndAlso TextBox97.Text = "" AndAlso TextBox98.Text = "" Then
            Label402.Enabled = True
            Label403.Enabled = True
            Label408.Enabled = True
            Label409.Enabled = True
            Label410.Enabled = True
            Label411.Enabled = True
            TextBox99.Enabled = True
            TextBox100.Enabled = True
            TextBox101.Enabled = True
            TextBox102.Enabled = True
        Else
            Label402.Enabled = False
            Label403.Enabled = False
            Label408.Enabled = False
            Label409.Enabled = False
            Label410.Enabled = False
            Label411.Enabled = False
            TextBox99.Enabled = False
            TextBox100.Enabled = False
            TextBox101.Enabled = False
            TextBox102.Enabled = False
        End If

        If TextBox95.Text IsNot "" AndAlso TextBox96.Text IsNot "" AndAlso TextBox97.Text IsNot "" AndAlso TextBox98.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub TextBox97_TextChanged(sender As Object, e As EventArgs) Handles TextBox97.TextChanged
        If TextBox95.Text = "" AndAlso TextBox96.Text = "" AndAlso TextBox97.Text = "" AndAlso TextBox98.Text = "" Then
            Label402.Enabled = True
            Label403.Enabled = True
            Label408.Enabled = True
            Label409.Enabled = True
            Label410.Enabled = True
            Label411.Enabled = True
            TextBox99.Enabled = True
            TextBox100.Enabled = True
            TextBox101.Enabled = True
            TextBox102.Enabled = True
        Else
            Label402.Enabled = False
            Label403.Enabled = False
            Label408.Enabled = False
            Label409.Enabled = False
            Label410.Enabled = False
            Label411.Enabled = False
            TextBox99.Enabled = False
            TextBox100.Enabled = False
            TextBox101.Enabled = False
            TextBox102.Enabled = False
        End If

        If TextBox95.Text IsNot "" AndAlso TextBox96.Text IsNot "" AndAlso TextBox97.Text IsNot "" AndAlso TextBox98.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub TextBox98_TextChanged(sender As Object, e As EventArgs) Handles TextBox98.TextChanged
        If TextBox95.Text = "" AndAlso TextBox96.Text = "" AndAlso TextBox97.Text = "" AndAlso TextBox98.Text = "" Then
            Label402.Enabled = True
            Label403.Enabled = True
            Label408.Enabled = True
            Label409.Enabled = True
            Label410.Enabled = True
            Label411.Enabled = True
            TextBox99.Enabled = True
            TextBox100.Enabled = True
            TextBox101.Enabled = True
            TextBox102.Enabled = True
        Else
            Label402.Enabled = False
            Label403.Enabled = False
            Label408.Enabled = False
            Label409.Enabled = False
            Label410.Enabled = False
            Label411.Enabled = False
            TextBox99.Enabled = False
            TextBox100.Enabled = False
            TextBox101.Enabled = False
            TextBox102.Enabled = False
        End If

        If TextBox95.Text IsNot "" AndAlso TextBox96.Text IsNot "" AndAlso TextBox97.Text IsNot "" AndAlso TextBox98.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub TextBox99_TextChanged(sender As Object, e As EventArgs) Handles TextBox99.TextChanged
        If TextBox99.Text = "" AndAlso TextBox100.Text = "" AndAlso TextBox101.Text = "" AndAlso TextBox102.Text = "" Then
            Label400.Enabled = True
            Label401.Enabled = True
            Label404.Enabled = True
            Label405.Enabled = True
            Label406.Enabled = True
            Label407.Enabled = True
            TextBox95.Enabled = True
            TextBox96.Enabled = True
            TextBox97.Enabled = True
            TextBox98.Enabled = True
        Else
            Label400.Enabled = False
            Label401.Enabled = False
            Label404.Enabled = False
            Label405.Enabled = False
            Label406.Enabled = False
            Label407.Enabled = False
            TextBox95.Enabled = False
            TextBox96.Enabled = False
            TextBox97.Enabled = False
            TextBox98.Enabled = False
        End If

        If TextBox99.Text IsNot "" AndAlso TextBox100.Text IsNot "" AndAlso TextBox101.Text IsNot "" AndAlso TextBox102.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub TextBox100_TextChanged(sender As Object, e As EventArgs) Handles TextBox100.TextChanged
        If TextBox99.Text = "" AndAlso TextBox100.Text = "" AndAlso TextBox101.Text = "" AndAlso TextBox102.Text = "" Then
            Label400.Enabled = True
            Label401.Enabled = True
            Label404.Enabled = True
            Label405.Enabled = True
            Label406.Enabled = True
            Label407.Enabled = True
            TextBox95.Enabled = True
            TextBox96.Enabled = True
            TextBox97.Enabled = True
            TextBox98.Enabled = True
        Else
            Label400.Enabled = False
            Label401.Enabled = False
            Label404.Enabled = False
            Label405.Enabled = False
            Label406.Enabled = False
            Label407.Enabled = False
            TextBox95.Enabled = False
            TextBox96.Enabled = False
            TextBox97.Enabled = False
            TextBox98.Enabled = False
        End If

        If TextBox99.Text IsNot "" AndAlso TextBox100.Text IsNot "" AndAlso TextBox101.Text IsNot "" AndAlso TextBox102.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub TextBox101_TextChanged(sender As Object, e As EventArgs) Handles TextBox101.TextChanged
        If TextBox99.Text = "" AndAlso TextBox100.Text = "" AndAlso TextBox101.Text = "" AndAlso TextBox102.Text = "" Then
            Label400.Enabled = True
            Label401.Enabled = True
            Label404.Enabled = True
            Label405.Enabled = True
            Label406.Enabled = True
            Label407.Enabled = True
            TextBox95.Enabled = True
            TextBox96.Enabled = True
            TextBox97.Enabled = True
            TextBox98.Enabled = True
        Else
            Label400.Enabled = False
            Label401.Enabled = False
            Label404.Enabled = False
            Label405.Enabled = False
            Label406.Enabled = False
            Label407.Enabled = False
            TextBox95.Enabled = False
            TextBox96.Enabled = False
            TextBox97.Enabled = False
            TextBox98.Enabled = False
        End If

        If TextBox99.Text IsNot "" AndAlso TextBox100.Text IsNot "" AndAlso TextBox101.Text IsNot "" AndAlso TextBox102.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub TextBox102_TextChanged(sender As Object, e As EventArgs) Handles TextBox102.TextChanged
        If TextBox99.Text = "" AndAlso TextBox100.Text = "" AndAlso TextBox101.Text = "" AndAlso TextBox102.Text = "" Then
            Label400.Enabled = True
            Label401.Enabled = True
            Label404.Enabled = True
            Label405.Enabled = True
            Label406.Enabled = True
            Label407.Enabled = True
            TextBox95.Enabled = True
            TextBox96.Enabled = True
            TextBox97.Enabled = True
            TextBox98.Enabled = True
        Else
            Label400.Enabled = False
            Label401.Enabled = False
            Label404.Enabled = False
            Label405.Enabled = False
            Label406.Enabled = False
            Label407.Enabled = False
            TextBox95.Enabled = False
            TextBox96.Enabled = False
            TextBox97.Enabled = False
            TextBox98.Enabled = False
        End If

        If TextBox99.Text IsNot "" AndAlso TextBox100.Text IsNot "" AndAlso TextBox101.Text IsNot "" AndAlso TextBox102.Text IsNot "" Then
            Call fatt()
        Else
            TextBox103.Text = ""
        End If

    End Sub

    Private Sub ComboBox172_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox172.SelectedIndexChanged
        Call fatt()
    End Sub

    '885°F Embrittlement Damage Factor -------------------------------------------------------------------------
    Private Sub ComboBox34_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox34.SelectedIndexChanged
        If ComboBox34.Text = "Yes" Then
            TextBox46.ReadOnly = False
            ComboBox169.Visible = False
            ComboBox169.Text = ""
            TextBox46.Text = ""
        ElseIf ComboBox34.Text = "No" Then
            TextBox46.ReadOnly = True
            ComboBox169.Visible = True
        Else
            TextBox46.ReadOnly = True
        End If
    End Sub

    Private Sub ComboBox169_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox169.SelectedIndexChanged
        If ComboBox169.Text = "Minimum Design Temperature" Then
            TextBox46.ReadOnly = False
            TextBox46.Text = ""
        ElseIf ComboBox169.Text = "Minimum Operating Temperature" Then
            TextBox46.ReadOnly = False
            TextBox46.Text = ""
        Else
            TextBox46.ReadOnly = True
        End If
    End Sub

    Private Sub ComboBox84_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox84.SelectedIndexChanged
        If ComboBox84.Text = "Type 405" Then
            TextBox48.Text = "11.5 - 14.5"
        ElseIf ComboBox84.Text = "Type 430" Then
            TextBox48.Text = "16 - 18"
        ElseIf ComboBox84.Text = "Type 430F" Then
            TextBox48.Text = "16 - 18"
        ElseIf ComboBox84.Text = "Type 442" Then
            TextBox48.Text = "18 - 23"
        ElseIf ComboBox84.Text = "Type 446" Then
            TextBox48.Text = "23 - 27"
        End If
    End Sub

    'Sigma Phase Embrittlement Damage Factor -------------------------------------------------------------------


    'Piping Mechanical Fatigue Damage Factor -------------------------------------------------------------------

    'FMS -------------------------------------------------------------------------------------------------------

    Private Sub CheckBox196_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox196.CheckedChanged
        If CheckBox196.Checked = True Then
            TextBox80.Text = FMS.TextBox2.Text
        End If
    End Sub

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

    Public Sub tampildataindustrycombobox()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT `industry_name` FROM `tbl_industry`", constring)

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader
        ComboBox2.Items.Clear()
        Do While rd.Read
            ComboBox2.Items.Add(rd.Item(0))
        Loop
    End Sub

    Public Sub tampildatacomponentcombobox()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT `component_type` FROM `tbl_equipment` WHERE `equipment_type`=@type AND `industry_name`=@name", constring)
        CMD.Parameters.Add("@type", MySqlDbType.VarChar).Value = Label391.Text
        CMD.Parameters.Add("@name", MySqlDbType.VarChar).Value = ComboBox2.Text

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader
        ComboBox62.Items.Clear()
        Do While rd.Read
            ComboBox62.Items.Add(rd.Item(0))
        Loop
    End Sub

    Public Sub tampildatacodeequipmentcombobox()
        Call koneksi()
        Dim CMD As New MySqlCommand("SELECT `equipment_code` FROM `tbl_equipment` WHERE `equipment_type`=@type AND `component_type`=@comp AND `industry_name`=@name", constring)
        CMD.Parameters.Add("@type", MySqlDbType.VarChar).Value = Label391.Text
        CMD.Parameters.Add("@comp", MySqlDbType.VarChar).Value = ComboBox62.Text
        CMD.Parameters.Add("@name", MySqlDbType.VarChar).Value = ComboBox2.Text

        Dim adapter As New MySqlDataAdapter(CMD)
        Dim rd As MySqlDataReader

        rd = CMD.ExecuteReader
        ComboBox63.Items.Clear()
        Do While rd.Read
            ComboBox63.Items.Add(rd.Item(0))
        Loop
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Call tampildatamaterial()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Call tampildatafluida()
    End Sub


    '=========================================================================================================================
    '                                                      INTERPOLASI


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

    '=============================================================================================================================
    '                                          INTERPOLASI BRITTLE DAMAGE FACTOR

    Private Function Interx(ByVal temp As Double, ByVal Up1 As Double, ByVal Up2 As Double, ByVal Lo1 As Double, ByVal Lo2 As Double) As Double
        Interx = ((temp - Lo1) * (Up2 - Lo2) / (Up1 - Lo1)) + Lo2
    End Function

    Private Function Interx1(ByVal temp As Double, ByVal Up3 As Double, ByVal Up4 As Double, ByVal Lo3 As Double, ByVal Lo4 As Double) As Double
        Interx1 = ((temp - Lo3) * (Up4 - Lo4) / (Up3 - Lo3)) + Lo4
    End Function

    Private Function Inter3a1(ByVal upt As Double, ByVal Lot As Double, ByVal ph As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3a1 = ((ph - Lot) * (z - x) / (upt - Lot)) + x
    End Function

    'extrapolation lower

    Private Function Inter3b1(ByVal upt As Double, ByVal Lot As Double, ByVal ph As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3b1 = ((ph - upt) * (x - z) / (Lot - upt)) + z
    End Function

    'extrapolation upper
    Private Function Inter3c1(ByVal upt As Double, ByVal Lot As Double, ByVal ph As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3c1 = ((ph - Lot) * (z - x) / (upt - Lot)) + x
    End Function

    Private Function Inter3da(ByVal upta As Double, ByVal Lota As Double, ByVal temp As Double, ByVal xa As Double, ByVal za As Double) As Double
        Inter3da = ((temp - upta) * (xa - za) / (Lota - upta)) + za
    End Function

    Private Function Inter3db(ByVal uptb As Double, ByVal Lotb As Double, ByVal temp As Double, ByVal xb As Double, ByVal zb As Double) As Double
        Inter3db = ((temp - uptb) * (xb - zb) / (Lotb - uptb)) + zb
    End Function

    Private Function Inter3ex(ByVal uptx As Double, ByVal Lotx As Double, ByVal temp As Double, ByVal xx As Double, ByVal zx As Double) As Double
        Inter3ex = ((temp - Lotx) * (zx - xx) / (uptx - Lotx)) + xx
    End Function

    Private Function Inter3ez(ByVal uptz As Double, ByVal Lotz As Double, ByVal temp As Double, ByVal xz As Double, ByVal zz As Double) As Double
        Inter3ez = ((temp - Lotz) * (zz - xz) / (uptz - Lotz)) + xz
    End Function

    '=============================================================================================================================
    Function Cumnorm(x As Double) As Double
        Dim xabs As Double
        Dim exponential As Double
        Dim build As Double


        xabs = Math.Abs(x)
        If xabs > 37 Then
            Cumnorm = 0
        Else
            exponential = Math.Exp(-xabs ^ 2 / 2)
            If xabs < 7.07106781186547 Then
                build = 0.0352624965998911 * xabs + 0.700383064443688
                build = build * xabs + 6.37396220353165
                build = build * xabs + 33.912866078383
                build = build * xabs + 112.079291497871
                build = build * xabs + 221.213596169931
                build = build * xabs + 220.206867912376
                Cumnorm = exponential * build
                build = 0.0883883476483184 * xabs + 1.75566716318264
                build = build * xabs + 16.064177579207
                build = build * xabs + 86.7807322029461
                build = build * xabs + 296.564248779674
                build = build * xabs + 637.333633378831
                build = build * xabs + 793.826512519948
                build = build * xabs + 440.413735824752
                Cumnorm = Cumnorm / build
            Else
                build = xabs + 0.65
                build = xabs + 4 / build
                build = xabs + 3 / build
                build = xabs + 2 / build
                build = xabs + 1 / build
                Cumnorm = exponential / build / 2.506628274631
            End If
        End If

        If x > 0 Then
            Cumnorm = 1 - Cumnorm
        End If
    End Function

    '=============================================================================================================================
    '                                                   PERHITUNGAN DF

    'Thinning Coding perhitungan ------------------------------------------------------------------------------------------



    'Lining Coding perhitungan --------------------------------------------------------------------------------------------
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

    'Scc – Caustic Cracking Coding perhitungan ----------------------------------------------------------------------------
    Private Sub plotarea()
        Dim X As Double
        Dim Y As Double
        Dim Y1 As Double 'equation of curve 1, lower curve
        Dim Y2 As Double 'equation of curve 2, higher curve for 0<x<20
        Dim Y3 As Double 'equation of curve 2, higher curve for 20<x<35
        Dim Y4 As Double 'equation of curve 2, higher curve for 30<x<50

        X = Val(TextBox35.Text)

        If Units.ComboBox1.Text = "SI" Then
            Y = Val(TextBox36.Text) * (9 / 5) + 32
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            Y = Val(TextBox36.Text)
        End If


        Y1 = 0.00001 * Math.Pow(X, 4) - 0.0001 * Math.Pow(X, 3) - 0.0347 * Math.Pow(X, 2) - 0.4521 * X + 179.69
        Y2 = -0.00007 * Math.Pow(X, 4) - 0.0009 * Math.Pow(X, 3) + 0.1056 * Math.Pow(X, 2) - 0.269 * X + 212.55
        Y3 = 0.0115 * Math.Pow(X, 3) - 1.018 * Math.Pow(X, 2) + 26.683 * X + 13.8
        Y4 = 0.0026 * Math.Pow(X, 3) - 0.2557 * Math.Pow(X, 2) + 6.2922 * X + 175.5


        If Y < Y1 Then
            TextBox37.Text = "A"
        ElseIf X >= 0 And X < 20 Then
            If Y > Y1 And Y < Y2 Then
                TextBox37.Text = "B"
            Else
                TextBox37.Text = "C"
            End If
        ElseIf X >= 20 And X < 35 Then
            If Y > Y1 And Y < Y3 Then
                TextBox37.Text = "B"
            Else
                TextBox37.Text = "C"
            End If
        ElseIf X >= 35 Then
            If Y > Y1 And Y < Y4 Then
                TextBox37.Text = "B"
            Else
                TextBox37.Text = "C"
            End If
        End If

        If TextBox35.Text = "" Or TextBox36.Text = "" Then
            TextBox37.Text = ""
        End If
    End Sub

    Private Sub susceptibilitycaustic()
        If TextBox37.Text = "A" Then
            If ComboBox81.Text = "Yes" Then
                ComboBox82.Enabled = True
                ComboBox83.Enabled = True
                If ComboBox82.Text = "Yes" Then
                    ComboBox105.Text = "MEDIUM SUSCEPTIBILITY"
                    ComboBox83.Enabled = False
                ElseIf ComboBox82.Text = "No" Then
                    ComboBox83.Enabled = True
                    If ComboBox83.Text = "Yes" Then
                        ComboBox105.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox83.Text = "No" Then
                        ComboBox105.Text = "NONE SUSCEPTIBILITY"
                    End If
                End If
            ElseIf ComboBox81.Text = "No" Then
                ComboBox82.Enabled = True
                ComboBox83.Enabled = True
                If ComboBox82.Text = "Yes" Then
                    ComboBox105.Text = "HIGH SUSCEPTIBILITY"
                    ComboBox83.Enabled = False
                ElseIf ComboBox82.Text = "No" Then
                    ComboBox83.Enabled = True
                    If ComboBox83.Text = "Yes" Then
                        ComboBox105.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox83.Text = "No" Then
                        ComboBox105.Text = "NONE SUSCEPTIBILITY"
                    End If
                End If
            End If
        Else
            If ComboBox81.Text = "Yes" Then
                ComboBox82.Enabled = False
                ComboBox83.Enabled = False
                ComboBox105.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf ComboBox81.Text = "No" Then
                ComboBox82.Enabled = False
                ComboBox83.Enabled = False
                ComboBox105.Text = "HIGH SUSCEPTIBILITY"
            End If
        End If
    End Sub

    Public Sub scccausticdf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox104.Text = "A" Or ComboBox104.Text = "B" Or ComboBox104.Text = "C" AndAlso ComboBox6.Text = "No" Or ComboBox103.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox105.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 5000
        ElseIf ComboBox105.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 500
        ElseIf ComboBox105.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 50
        ElseIf ComboBox105.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown1.Value = 1 And ComboBox104.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "A" And severityindex = "50" Then
            basedf = "3"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "A" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "A" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "B" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "B" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "B" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "C" And severityindex = "50" Then
            basedf = "17"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "C" And severityindex = "500" Then
            basedf = "170"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "C" And severityindex = “5000” Then
            basedf = "1670"

        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "D" And severityindex = "50" Then
            basedf = "40"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "D" And severityindex = "500" Then
            basedf = "400"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "D" And severityindex = “5000” Then
            basedf = "4000"

        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown1.Value = 1 And ComboBox104.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "A" And severityindex = "500" Then
            basedf = "5"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "A" And severityindex = “5000” Then
            basedf = "50"

        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "B" And severityindex = "50" Then
            basedf = "2"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "B" And severityindex = "500" Then
            basedf = "20"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "B" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "C" And severityindex = "50" Then
            basedf = "10"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "C" And severityindex = "500" Then
            basedf = "100"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "C" And severityindex = “5000” Then
            basedf = "1000"

        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "D" And severityindex = "50" Then
            basedf = "30"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "D" And severityindex = "500" Then
            basedf = "300"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "D" And severityindex = “5000” Then
            basedf = "3000"

        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown1.Value = 2 And ComboBox104.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "A" And severityindex = “5000” Then
            basedf = "10"

        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "B" And severityindex = "500" Then
            basedf = "8"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "B" And severityindex = “5000” Then
            basedf = "80"

        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "C" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "C" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "C" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "D" And severityindex = "50" Then
            basedf = "20"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "D" And severityindex = "500" Then
            basedf = "200"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "D" And severityindex = “5000” Then
            basedf = "2000"

        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown1.Value = 3 And ComboBox104.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "A" And severityindex = “5000” Then
            basedf = "2"

        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "B" And severityindex = "500" Then
            basedf = "2"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "B" And severityindex = “5000” Then
            basedf = "25"

        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "C" And severityindex = "50" Then
            basedf = "2"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "C" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "C" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "D" And severityindex = "50" Then
            basedf = "10"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "D" And severityindex = "500" Then
            basedf = "100"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "D" And severityindex = “5000” Then
            basedf = "1000"

        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown1.Value = 4 And ComboBox104.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "A" And severityindex = “5000” Then
            basedf = "1"

        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "B" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "B" And severityindex = “5000” Then
            basedf = "5"

        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "C" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "C" And severityindex = "500" Then
            basedf = "10"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "C" And severityindex = “5000” Then
            basedf = "125"

        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "D" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "D" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "D" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown1.Value = 5 And ComboBox104.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "A" And severityindex = “5000” Then
            basedf = "1"

        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "B" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "B" And severityindex = “5000” Then
            basedf = "2"

        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "C" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "C" And severityindex = "500" Then
            basedf = "5"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "C" And severityindex = “5000” Then
            basedf = "50"

        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "D" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "D" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "D" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown1.Value = 6 And ComboBox104.Text = "E" And severityindex = “5000” Then
            basedf = “5000”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = basedf * (maxage1 ^ 1.1)
            Label377.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = basedf * (maxage2 ^ 1.1)
            Label377.Text = totaldf
        End If
    End Sub

    'Scc – Amine Cracking Coding perhitungan ------------------------------------------------------------------------------
    Private Sub susceptibilityamine()
        If ComboBox86.Text = "Yes" Then
            ComboBox87.Enabled = False
            If ComboBox88.Text = "Yes" Then
                ComboBox108.Text = "HIGH SUSCEPTIBILITY"
                ComboBox89.Enabled = False
                ComboBox90.Enabled = False
                ComboBox91.Enabled = False
                ComboBox92.Enabled = False
            ElseIf ComboBox88.Text = "No" Then
                ComboBox89.Enabled = True
                ComboBox90.Enabled = False
                ComboBox91.Enabled = True
                ComboBox92.Enabled = True
                If ComboBox89.Text = "Yes" Then
                    ComboBox108.Text = "MEDIUM SUSCEPTIBILITY"
                    ComboBox91.Enabled = False
                    ComboBox92.Enabled = False
                ElseIf ComboBox89.Text = "No" Then
                    ComboBox91.Enabled = True
                    ComboBox92.Enabled = True
                    If ComboBox91.Text = "Yes" Then
                        ComboBox108.Text = "MEDIUM SUSCEPTIBILITY"
                        ComboBox92.Enabled = False
                    ElseIf ComboBox91.Text = "No" Then
                        ComboBox92.Enabled = True
                        If ComboBox92.Text = "Yes" Then
                            ComboBox108.Text = "MEDIUM SUSCEPTIBILITY"
                        ElseIf ComboBox92.Text = "No" Then
                            ComboBox108.Text = "LOW SUSCEPTIBILITY"
                        End If
                    End If
                End If
            End If
        ElseIf ComboBox86.Text = "No" Then
            ComboBox87.Enabled = True
            If ComboBox87.Text = "Yes" Then
                If ComboBox88.Text = "Yes" Then
                    ComboBox108.Text = "MEDIUM SUSCEPTIBILITY"
                    ComboBox89.Enabled = False
                    ComboBox90.Enabled = False
                    ComboBox91.Enabled = False
                    ComboBox92.Enabled = False
                ElseIf ComboBox88.Text = "No" Then
                    ComboBox89.Enabled = False
                    ComboBox90.Enabled = True
                    ComboBox91.Enabled = True
                    ComboBox92.Enabled = True
                    If ComboBox90.Text = "Yes" Then
                        ComboBox108.Text = "LOW SUSCEPTIBILITY"
                        ComboBox91.Enabled = False
                        ComboBox92.Enabled = False
                    ElseIf ComboBox89.Text = "No" Then
                        ComboBox91.Enabled = True
                        ComboBox92.Enabled = True
                        If ComboBox91.Text = "Yes" Then
                            ComboBox108.Text = "LOW SUSCEPTIBILITY"
                            ComboBox92.Enabled = False
                        ElseIf ComboBox91.Text = "No" Then
                            ComboBox92.Enabled = True
                            If ComboBox92.Text = "Yes" Then
                                ComboBox108.Text = "LOW SUSCEPTIBILITY"
                            ElseIf ComboBox92.Text = "No" Then
                                ComboBox108.Text = "NONE SUSCEPTIBILITY"
                            End If
                        End If
                    End If
                End If
            ElseIf ComboBox87.Text = "No" Then
                If ComboBox88.Text = "Yes" Then
                    ComboBox108.Text = "LOW SUSCEPTIBILITY"
                    ComboBox89.Enabled = False
                    ComboBox90.Enabled = False
                    ComboBox91.Enabled = False
                    ComboBox92.Enabled = False
                ElseIf ComboBox88.Text = "No" Then
                    ComboBox89.Enabled = False
                    ComboBox90.Enabled = False
                    ComboBox91.Enabled = True
                    ComboBox92.Enabled = True
                    If ComboBox91.Text = "Yes" Then
                        ComboBox108.Text = "LOW SUSCEPTIBILITY"
                        ComboBox92.Enabled = False
                    ElseIf ComboBox91.Text = "No" Then
                        ComboBox92.Enabled = True
                        If ComboBox92.Text = "Yes" Then
                            ComboBox108.Text = "LOW SUSCEPTIBILITY"
                        ElseIf ComboBox92.Text = "No" Then
                            ComboBox108.Text = "NONE SUSCEPTIBILITY"
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub sccaminedf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox107.Text = "A" Or ComboBox107.Text = "B" Or ComboBox107.Text = "C" AndAlso ComboBox8.Text = "No" Or ComboBox9.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox108.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 1000
        ElseIf ComboBox108.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 100
        ElseIf ComboBox108.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf ComboBox108.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown2.Value = 1 And ComboBox107.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "A" And severityindex = "100" Then
            basedf = "5"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "A" And severityindex = “1000” Then
            basedf = "50"

        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "B" And severityindex = "100" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "B" And severityindex = “1000” Then
            basedf = "100"

        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "C" And severityindex = "100" Then
            basedf = "33"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "C" And severityindex = “1000” Then
            basedf = "330"

        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "D" And severityindex = "100" Then
            basedf = "80"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "D" And severityindex = “1000” Then
            basedf = "800"

        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown2.Value = 1 And ComboBox107.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "A" And severityindex = “1000” Then
            basedf = "10"

        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "B" And severityindex = "100" Then
            basedf = "4"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "B" And severityindex = “1000” Then
            basedf = "40"

        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "C" And severityindex = "100" Then
            basedf = "20"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "C" And severityindex = “1000” Then
            basedf = "200"

        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "D" And severityindex = "100" Then
            basedf = "60"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "D" And severityindex = “1000” Then
            basedf = "600"

        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown2.Value = 2 And ComboBox107.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "A" And severityindex = “1000” Then
            basedf = "2"

        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "B" And severityindex = "100" Then
            basedf = "2"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "B" And severityindex = “1000” Then
            basedf = "16"

        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "C" And severityindex = "100" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "C" And severityindex = “1000” Then
            basedf = "100"

        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "D" And severityindex = "100" Then
            basedf = "40"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "D" And severityindex = “1000” Then
            basedf = "400"

        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown2.Value = 3 And ComboBox107.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "A" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "B" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "B" And severityindex = “1000” Then
            basedf = "5"

        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "C" And severityindex = "100" Then
            basedf = "5"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "C" And severityindex = “1000” Then
            basedf = "50"

        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "D" And severityindex = "100" Then
            basedf = "20"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "D" And severityindex = “1000” Then
            basedf = "200"

        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown2.Value = 4 And ComboBox107.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "A" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "B" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "B" And severityindex = “1000” Then
            basedf = "2"

        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "C" And severityindex = "100" Then
            basedf = "2"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "C" And severityindex = “1000” Then
            basedf = "25"

        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "D" And severityindex = "100" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "D" And severityindex = “1000” Then
            basedf = "100"

        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown2.Value = 5 And ComboBox107.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "A" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "B" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "B" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "C" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "C" And severityindex = “1000” Then
            basedf = "10"

        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "D" And severityindex = "100" Then
            basedf = "5"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "D" And severityindex = “1000” Then
            basedf = "50"

        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown2.Value = 6 And ComboBox107.Text = "E" And severityindex = “1000” Then
            basedf = “1000”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = basedf * (maxage1 ^ 1.1)
            Label368.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = basedf * (maxage2 ^ 1.1)
            Label368.Text = totaldf
        End If
    End Sub

    'Scc – Sulfide Stress Cracking Coding perhitungan ---------------------------------------------------------------------
    Private Sub enviromentalseveritysulfide()
        If ComboBox10.Text = "< 5.5" Then
            If ComboBox12.Text = "< 50 ppm" Then
                TextBox38.Text = "LOW"
            ElseIf ComboBox12.Text = "50 to 1,000 ppm" Then
                TextBox38.Text = "MODERATE"
            ElseIf ComboBox12.Text = "1,000 to 10,000 ppm" Then
                TextBox38.Text = "HIGH"
            ElseIf ComboBox12.Text = "> 10,000 ppm" Then
                TextBox38.Text = "HIGH"
            End If

        ElseIf ComboBox10.Text = "5.5 to 7.5" Then
            If ComboBox12.Text = "< 50 ppm" Then
                TextBox38.Text = "LOW"
            ElseIf ComboBox12.Text = "50 to 1,000 ppm" Then
                TextBox38.Text = "LOW"
            ElseIf ComboBox12.Text = "1,000 to 10,000 ppm" Then
                TextBox38.Text = "LOW"
            ElseIf ComboBox12.Text = "> 10,000 ppm" Then
                TextBox38.Text = "MODERATE"
            End If

        ElseIf ComboBox10.Text = "7.6 to 8.3" Then
            If ComboBox12.Text = "< 50 ppm" Then
                TextBox38.Text = "LOW"
            ElseIf ComboBox12.Text = "50 to 1,000 ppm" Then
                TextBox38.Text = "MODERATE"
            ElseIf ComboBox12.Text = "1,000 to 10,000 ppm" Then
                TextBox38.Text = "MODERATE"
            ElseIf ComboBox12.Text = "> 10,000 ppm" Then
                TextBox38.Text = "MODERATE"
            End If

        ElseIf ComboBox10.Text = "8.4 to 8.9" Then
            If ComboBox12.Text = "< 50 ppm" Then
                TextBox38.Text = "LOW"
            ElseIf ComboBox12.Text = "50 to 1,000 ppm" Then
                TextBox38.Text = "MODERATE"
            ElseIf ComboBox12.Text = "1,000 to 10,000 ppm" Then
                TextBox38.Text = "MODERATE"
            ElseIf ComboBox12.Text = "> 10,000 ppm" Then
                TextBox38.Text = "HIGH"
            End If

        ElseIf ComboBox10.Text = "> 9.0" Then
            If ComboBox12.Text = "< 50 ppm" Then
                TextBox38.Text = "LOW"
            ElseIf ComboBox12.Text = "50 to 1,000 ppm" Then
                TextBox38.Text = "MODERATE"
            ElseIf ComboBox12.Text = "1,000 to 10,000 ppm" Then
                TextBox38.Text = "HIGH"
            ElseIf ComboBox12.Text = "> 10,000 ppm" Then
                TextBox38.Text = "HIGH"
            End If
        End If
    End Sub

    Private Sub susceptibilitysulfide()
        If TextBox38.Text = "HIGH" Then
            If ComboBox80.Text = "As-Welded" Then
                If ComboBox93.Text = "< 200" Then
                    ComboBox110.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "200-237" Then
                    ComboBox110.Text = "MEDIUM SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "> 237" Then
                    ComboBox110.Text = "HIGH SUSCEPTIBILITY"
                End If
            ElseIf ComboBox80.Text = "PWHT" Then
                If ComboBox93.Text = "< 200" Then
                    ComboBox110.Text = "NONE SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "200-237" Then
                    ComboBox110.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "> 237" Then
                    ComboBox110.Text = "MEDIUM SUSCEPTIBILITY"
                End If
            End If
        ElseIf TextBox38.Text = "MODERATE" Then
            If ComboBox80.Text = "As-Welded" Then
                If ComboBox93.Text = "< 200" Then
                    ComboBox110.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "200-237" Then
                    ComboBox110.Text = "MEDIUM SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "> 237" Then
                    ComboBox110.Text = "HIGH SUSCEPTIBILITY"
                End If
            ElseIf ComboBox80.Text = "PWHT" Then
                If ComboBox93.Text = "< 200" Then
                    ComboBox110.Text = "NONE SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "200-237" Then
                    ComboBox110.Text = "NONE SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "> 237" Then
                    ComboBox110.Text = "LOW SUSCEPTIBILITY"
                End If
            End If
        ElseIf TextBox38.Text = "LOW" Then
            If ComboBox80.Text = "As-Welded" Then
                If ComboBox93.Text = "< 200" Then
                    ComboBox110.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "200-237" Then
                    ComboBox110.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "> 237" Then
                    ComboBox110.Text = "MEDIUM SUSCEPTIBILITY"
                End If
            ElseIf ComboBox80.Text = "PWHT" Then
                If ComboBox93.Text = "< 200" Then
                    ComboBox110.Text = "NONE SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "200-237" Then
                    ComboBox110.Text = "NONE SUSCEPTIBILITY"
                ElseIf ComboBox93.Text = "> 237" Then
                    ComboBox110.Text = "NONE SUSCEPTIBILITY"
                End If
            End If
        End If
    End Sub

    Public Sub sccsulfidedf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox109.Text = "A" Or ComboBox109.Text = "B" Or ComboBox109.Text = "C" AndAlso ComboBox94.Text = "No" Or ComboBox95.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox110.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 100
        ElseIf ComboBox110.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf ComboBox110.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 1
        ElseIf ComboBox110.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown3.Value = 1 And ComboBox109.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "A" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "B" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "C" And severityindex = “100” Then
            basedf = "33"

        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "D" And severityindex = “100” Then
            basedf = "80"

        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown3.Value = 1 And ComboBox109.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "B" And severityindex = “100” Then
            basedf = "4"

        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "C" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "D" And severityindex = “100” Then
            basedf = "60"

        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown3.Value = 2 And ComboBox109.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "B" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "C" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "D" And severityindex = “100” Then
            basedf = "40"

        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown3.Value = 3 And ComboBox109.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "C" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "D" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown3.Value = 4 And ComboBox109.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "C" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "D" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown3.Value = 5 And ComboBox109.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "C" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "D" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown3.Value = 6 And ComboBox109.Text = "E" And severityindex = “100” Then
            basedf = “100”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = basedf * (maxage1 ^ 1.1)
            Label369.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = basedf * (maxage2 ^ 1.1)
            Label369.Text = totaldf
        End If
    End Sub

    'Scc – Hic/Sohic-H2S Coding perhitungan -------------------------------------------------------------------------------
    Private Sub susceptibilityhicsohich2s()
        If TextBox38.Text = "HIGH" Then
            If ComboBox80.Text = "As-Welded" Then
                If ComboBox11.Text = "High Sulfur Steel > 0.01% S" Then
                    ComboBox112.Text = "HIGH SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Low Sulfur Steel < 0.01% S" Then
                    ComboBox112.Text = "HIGH SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Product Form – Seamless/Extruded Pipe" Then
                    ComboBox112.Text = "MEDIUM SUSCEPTIBILITY"
                End If
            ElseIf ComboBox80.Text = "PWHT" Then
                If ComboBox11.Text = "High Sulfur Steel > 0.01% S" Then
                    ComboBox112.Text = "HIGH SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Low Sulfur Steel < 0.01% S" Then
                    ComboBox112.Text = "MEDIUM SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Product Form – Seamless/Extruded Pipe" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                End If
            End If
        ElseIf TextBox38.Text = "MODERATE" Then
            If ComboBox80.Text = "As-Welded" Then
                If ComboBox11.Text = "High Sulfur Steel > 0.01% S" Then
                    ComboBox112.Text = "HIGH SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Low Sulfur Steel < 0.01% S" Then
                    ComboBox112.Text = "MEDIUM SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Product Form – Seamless/Extruded Pipe" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                End If
            ElseIf ComboBox80.Text = "PWHT" Then
                If ComboBox11.Text = "High Sulfur Steel > 0.01% S" Then
                    ComboBox112.Text = "MEDIUM SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Low Sulfur Steel < 0.01% S" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Product Form – Seamless/Extruded Pipe" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                End If
            End If
        ElseIf TextBox38.Text = "LOW" Then
            If ComboBox80.Text = "As-Welded" Then
                If ComboBox11.Text = "High Sulfur Steel > 0.01% S" Then
                    ComboBox112.Text = "MEDIUM SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Low Sulfur Steel < 0.01% S" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Product Form – Seamless/Extruded Pipe" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                End If
            ElseIf ComboBox80.Text = "PWHT" Then
                If ComboBox11.Text = "High Sulfur Steel > 0.01% S" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Low Sulfur Steel < 0.01% S" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox11.Text = "Product Form – Seamless/Extruded Pipe" Then
                    ComboBox112.Text = "LOW SUSCEPTIBILITY"
                End If
            End If
        End If
    End Sub

    Public Sub scchicsohich2s()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox111.Text = "A" Or ComboBox111.Text = "B" Or ComboBox111.Text = "C" AndAlso ComboBox13.Text = "No" Or ComboBox17.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox112.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 100
        ElseIf ComboBox112.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf ComboBox112.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 1
        ElseIf ComboBox112.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown4.Value = 1 And ComboBox111.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "A" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "B" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "C" And severityindex = “100” Then
            basedf = "33"

        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "D" And severityindex = “100” Then
            basedf = "80"

        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown4.Value = 1 And ComboBox111.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "B" And severityindex = “100” Then
            basedf = "4"

        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "C" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "D" And severityindex = “100” Then
            basedf = "60"

        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown4.Value = 2 And ComboBox111.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "B" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "C" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "D" And severityindex = “100” Then
            basedf = "40"

        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown4.Value = 3 And ComboBox111.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "C" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "D" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown4.Value = 4 And ComboBox111.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "C" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "D" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown4.Value = 5 And ComboBox111.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "C" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "D" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown4.Value = 6 And ComboBox111.Text = "E" And severityindex = “100” Then
            basedf = “100”
        End If

        Dim onlineadjfactor As Double

        If ComboBox96.Text = "Key Process Variables" Then
            onlineadjfactor = 2
        ElseIf ComboBox96.Text = "Hydrogen Probes" Then
            onlineadjfactor = 2
        ElseIf ComboBox96.Text = "Key Process Variables and Hydrogen Probes" Then
            onlineadjfactor = 4
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = (basedf * (maxage1 ^ 1.1)) / onlineadjfactor
            Label370.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = (basedf * (maxage2 ^ 1.1)) / onlineadjfactor
            Label370.Text = totaldf
        End If
    End Sub

    'Scc – Alkaline Carbonate Stress Corrosion Cracking Coding perhitungan ------------------------------------------------
    Private Sub susceptibilitysccalkalinecarbonate()
        If ComboBox14.Text = "< 7.5" Then
            If ComboBox15.Text = "PWHT" Then
                ComboBox114.Text = "NONE SUSCEPTIBILITY"
            ElseIf ComboBox15.Text = "No PWHT" Then
                ComboBox114.Text = ""
                If ComboBox98.Text = "CO3 < 100 ppm" Then
                    ComboBox114.Text = "NONE SUSCEPTIBILITY"
                ElseIf ComboBox98.Text = "CO3 ≥ 100 ppm" Then
                    ComboBox114.Text = "NONE SUSCEPTIBILITY"
                End If
            End If
        ElseIf ComboBox14.Text = "≥7.5 to 8.0" Then
            If ComboBox15.Text = "PWHT" Then
                ComboBox114.Text = "NONE SUSCEPTIBILITY"
            ElseIf ComboBox15.Text = "No PWHT" Then
                ComboBox114.Text = ""
                If ComboBox98.Text = "CO3 < 100 ppm" Then
                    ComboBox114.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox98.Text = "CO3 ≥ 100 ppm" Then
                    ComboBox114.Text = "MEDIUM SUSCEPTIBILITY"
                End If
            End If
        ElseIf ComboBox14.Text = "≥8.0 to 9.0" Then
            If ComboBox15.Text = "PWHT" Then
                ComboBox114.Text = "NONE SUSCEPTIBILITY"
            ElseIf ComboBox15.Text = "No PWHT" Then
                ComboBox114.Text = ""
                If ComboBox98.Text = "CO3 < 100 ppm" Then
                    ComboBox114.Text = "LOW SUSCEPTIBILITY"
                ElseIf ComboBox98.Text = "CO3 ≥ 100 ppm" Then
                    ComboBox114.Text = "HIGH SUSCEPTIBILITY"
                End If
            End If
        ElseIf ComboBox14.Text = "≥ 9.0" Then
            If ComboBox15.Text = "PWHT" Then
                ComboBox114.Text = "NONE SUSCEPTIBILITY"
            ElseIf ComboBox15.Text = "No PWHT" Then
                ComboBox114.Text = ""
                If ComboBox98.Text = "CO3 < 100 ppm" Then
                    ComboBox114.Text = "HIGH SUSCEPTIBILITY"
                ElseIf ComboBox98.Text = "CO3 ≥ 100 ppm" Then
                    ComboBox114.Text = "HIGH SUSCEPTIBILITY"
                End If
            End If
        End If
    End Sub

    Public Sub sccalkalinecarbonatedf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox113.Text = "A" Or ComboBox113.Text = "B" Or ComboBox113.Text = "C" AndAlso ComboBox16.Text = "No" Or ComboBox97.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox114.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 1000
        ElseIf ComboBox114.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 100
        ElseIf ComboBox114.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf ComboBox114.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown5.Value = 1 And ComboBox113.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "A" And severityindex = "100" Then
            basedf = "5"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "A" And severityindex = “1000” Then
            basedf = "50"

        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "B" And severityindex = "100" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "B" And severityindex = “1000” Then
            basedf = "100"

        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "C" And severityindex = "100" Then
            basedf = "33"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "C" And severityindex = “1000” Then
            basedf = "330"

        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "D" And severityindex = "100" Then
            basedf = "80"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "D" And severityindex = “1000” Then
            basedf = "800"

        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown5.Value = 1 And ComboBox113.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "A" And severityindex = “1000” Then
            basedf = "10"

        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "B" And severityindex = "100" Then
            basedf = "4"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "B" And severityindex = “1000” Then
            basedf = "40"

        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "C" And severityindex = "100" Then
            basedf = "20"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "C" And severityindex = “1000” Then
            basedf = "200"

        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "D" And severityindex = "100" Then
            basedf = "60"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "D" And severityindex = “1000” Then
            basedf = "600"

        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown5.Value = 2 And ComboBox113.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "A" And severityindex = “1000” Then
            basedf = "2"

        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "B" And severityindex = "100" Then
            basedf = "2"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "B" And severityindex = “1000” Then
            basedf = "16"

        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "C" And severityindex = "100" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "C" And severityindex = “1000” Then
            basedf = "100"

        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "D" And severityindex = "100" Then
            basedf = "40"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "D" And severityindex = “1000” Then
            basedf = "400"

        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown5.Value = 3 And ComboBox113.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "A" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "B" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "B" And severityindex = “1000” Then
            basedf = "5"

        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "C" And severityindex = "100" Then
            basedf = "5"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "C" And severityindex = “1000” Then
            basedf = "50"

        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "D" And severityindex = "100" Then
            basedf = "20"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "D" And severityindex = “1000” Then
            basedf = "200"

        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown5.Value = 4 And ComboBox113.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "A" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "B" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "B" And severityindex = “1000” Then
            basedf = "2"

        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "C" And severityindex = "100" Then
            basedf = "2"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "C" And severityindex = “1000” Then
            basedf = "25"

        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "D" And severityindex = "100" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "D" And severityindex = “1000” Then
            basedf = "100"

        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown5.Value = 5 And ComboBox113.Text = "E" And severityindex = “1000” Then
            basedf = “1000”


        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "A" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "A" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "B" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "B" And severityindex = “1000” Then
            basedf = "1"

        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "C" And severityindex = "100" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "C" And severityindex = “1000” Then
            basedf = "10"

        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "D" And severityindex = "100" Then
            basedf = "5"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "D" And severityindex = “1000” Then
            basedf = "50"

        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "E" And severityindex = "100" Then
            basedf = "100"
        ElseIf NumericUpDown5.Value = 6 And ComboBox113.Text = "E" And severityindex = “1000” Then
            basedf = “1000”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = basedf * (maxage1 ^ 1.1)
            Label371.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = basedf * (maxage2 ^ 1.1)
            Label371.Text = totaldf
        End If
    End Sub

    'Scc – Polythionic Acid Stress Corrosion Cracking Coding perhitungan --------------------------------------------------
    Private Sub susceptibilitypolythionicacid()
        If Units.ComboBox1.Text = "SI" Then
            If TextBox5.Text < 427 Then
                If ComboBox19.Text = "Solution Annealed (default)" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized Before Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized After Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If
            ElseIf TextBox5.Text >= 427 Then
                If ComboBox19.Text = "Solution Annealed (default)" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized Before Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized After Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If
            End If
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If TextBox5.Text < 800 Then
                If ComboBox19.Text = "Solution Annealed (default)" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized Before Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized After Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If
            ElseIf TextBox5.Text >= 800 Then
                If ComboBox19.Text = "Solution Annealed (default)" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "MEDIUM SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized Before Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox19.Text = "Stabilized After Welding" Then
                    If ComboBox18.Text = "All regular 300 series Stainless Steels and Alloys 600 and 800" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "H Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "L Grade 300 series SS" Then
                        ComboBox117.Text = ""
                    ElseIf ComboBox18.Text = "321 Stainless Steel" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox18.Text = "347 Stainless Steel, Alloy 20, Alloy 625, All austenitic weld overlay" Then
                        ComboBox117.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If
            End If
        End If


    End Sub

    Public Sub sccpolythionicaciddf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox116.Text = "A" Or ComboBox116.Text = "B" Or ComboBox116.Text = "C" AndAlso ComboBox115.Text = "No" Or ComboBox20.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox20.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 5000
        ElseIf ComboBox20.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 500
        ElseIf ComboBox20.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 50
        ElseIf ComboBox20.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown6.Value = 1 And ComboBox116.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "A" And severityindex = "50" Then
            basedf = "3"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "A" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "A" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "B" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "B" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "B" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "C" And severityindex = "50" Then
            basedf = "17"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "C" And severityindex = "500" Then
            basedf = "170"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "C" And severityindex = “5000” Then
            basedf = "1670"

        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "D" And severityindex = "50" Then
            basedf = "40"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "D" And severityindex = "500" Then
            basedf = "400"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "D" And severityindex = “5000” Then
            basedf = "4000"

        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown6.Value = 1 And ComboBox116.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "A" And severityindex = "500" Then
            basedf = "5"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "A" And severityindex = “5000” Then
            basedf = "50"

        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "B" And severityindex = "50" Then
            basedf = "2"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "B" And severityindex = "500" Then
            basedf = "20"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "B" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "C" And severityindex = "50" Then
            basedf = "10"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "C" And severityindex = "500" Then
            basedf = "100"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "C" And severityindex = “5000” Then
            basedf = "1000"

        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "D" And severityindex = "50" Then
            basedf = "30"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "D" And severityindex = "500" Then
            basedf = "300"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "D" And severityindex = “5000” Then
            basedf = "3000"

        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown6.Value = 2 And ComboBox116.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "A" And severityindex = “5000” Then
            basedf = "10"

        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "B" And severityindex = "500" Then
            basedf = "8"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "B" And severityindex = “5000” Then
            basedf = "80"

        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "C" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "C" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "C" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "D" And severityindex = "50" Then
            basedf = "20"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "D" And severityindex = "500" Then
            basedf = "200"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "D" And severityindex = “5000” Then
            basedf = "2000"

        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown6.Value = 3 And ComboBox116.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "A" And severityindex = “5000” Then
            basedf = "2"

        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "B" And severityindex = "500" Then
            basedf = "2"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "B" And severityindex = “5000” Then
            basedf = "25"

        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "C" And severityindex = "50" Then
            basedf = "2"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "C" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "C" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "D" And severityindex = "50" Then
            basedf = "10"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "D" And severityindex = "500" Then
            basedf = "100"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "D" And severityindex = “5000” Then
            basedf = "1000"

        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown6.Value = 4 And ComboBox116.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "A" And severityindex = “5000” Then
            basedf = "1"

        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "B" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "B" And severityindex = “5000” Then
            basedf = "5"

        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "C" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "C" And severityindex = "500" Then
            basedf = "10"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "C" And severityindex = “5000” Then
            basedf = "125"

        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "D" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "D" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "D" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown6.Value = 5 And ComboBox116.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "A" And severityindex = “5000” Then
            basedf = "1"

        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "B" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "B" And severityindex = “5000” Then
            basedf = "2"

        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "C" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "C" And severityindex = "500" Then
            basedf = "5"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "C" And severityindex = “5000” Then
            basedf = "50"

        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "D" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "D" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "D" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown6.Value = 6 And ComboBox116.Text = "E" And severityindex = “5000” Then
            basedf = “5000”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = basedf * (maxage1 ^ 1.1)
            Label372.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = basedf * (maxage2 ^ 1.1)
            Label372.Text = totaldf
        End If
    End Sub

    'Scc – Chloride Stress Corrosion Cracking Coding perhitungan ----------------------------------------------------------
    Private Sub susceptibilitysccchloride()
        If Units.ComboBox1.Text = "SI" Then

            If TextBox5.Text <= 38 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 38 AndAlso TextBox5.Text <= 66 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 66 AndAlso TextBox5.Text <= 93 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 93 AndAlso TextBox5.Text <= 149 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 149 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                End If
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then

            If TextBox5.Text <= 100 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "NONE SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 100 AndAlso TextBox5.Text <= 150 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 150 AndAlso TextBox5.Text <= 200 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 200 AndAlso TextBox5.Text <= 300 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "LOW SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    End If
                End If

            ElseIf TextBox5.Text > 300 Then
                If ComboBox23.Text = "pH ≤ 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                ElseIf ComboBox23.Text = "pH > 10" Then
                    If ComboBox22.Text = "1-10" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "11-100" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "101-1000" Then
                        ComboBox120.Text = "MEDIUM SUSCEPTIBILITY"
                    ElseIf ComboBox22.Text = "> 1000" Then
                        ComboBox120.Text = "HIGH SUSCEPTIBILITY"
                    End If
                End If
            End If

        End If
    End Sub

    Public Sub sccclsccdf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox119.Text = "A" Or ComboBox119.Text = "B" Or ComboBox119.Text = "C" AndAlso ComboBox99.Text = "No" Or ComboBox100.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox120.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 5000
        ElseIf ComboBox120.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 500
        ElseIf ComboBox120.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 50
        ElseIf ComboBox120.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown7.Value = 1 And ComboBox119.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "A" And severityindex = "50" Then
            basedf = "3"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "A" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "A" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "B" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "B" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "B" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "C" And severityindex = "50" Then
            basedf = "17"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "C" And severityindex = "500" Then
            basedf = "170"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "C" And severityindex = “5000” Then
            basedf = "1670"

        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "D" And severityindex = "50" Then
            basedf = "40"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "D" And severityindex = "500" Then
            basedf = "400"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "D" And severityindex = “5000” Then
            basedf = "4000"

        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown7.Value = 1 And ComboBox119.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "A" And severityindex = "500" Then
            basedf = "5"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "A" And severityindex = “5000” Then
            basedf = "50"

        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "B" And severityindex = "50" Then
            basedf = "2"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "B" And severityindex = "500" Then
            basedf = "20"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "B" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "C" And severityindex = "50" Then
            basedf = "10"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "C" And severityindex = "500" Then
            basedf = "100"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "C" And severityindex = “5000” Then
            basedf = "1000"

        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "D" And severityindex = "50" Then
            basedf = "30"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "D" And severityindex = "500" Then
            basedf = "300"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "D" And severityindex = “5000” Then
            basedf = "3000"

        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown7.Value = 2 And ComboBox119.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "A" And severityindex = “5000” Then
            basedf = "10"

        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "B" And severityindex = "500" Then
            basedf = "8"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "B" And severityindex = “5000” Then
            basedf = "80"

        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "C" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "C" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "C" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "D" And severityindex = "50" Then
            basedf = "20"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "D" And severityindex = "500" Then
            basedf = "200"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "D" And severityindex = “5000” Then
            basedf = "2000"

        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown7.Value = 3 And ComboBox119.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "A" And severityindex = “5000” Then
            basedf = "2"

        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "B" And severityindex = "500" Then
            basedf = "2"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "B" And severityindex = “5000” Then
            basedf = "25"

        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "C" And severityindex = "50" Then
            basedf = "2"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "C" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "C" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "D" And severityindex = "50" Then
            basedf = "10"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "D" And severityindex = "500" Then
            basedf = "100"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "D" And severityindex = “5000” Then
            basedf = "1000"

        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown7.Value = 4 And ComboBox119.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "A" And severityindex = “5000” Then
            basedf = "1"

        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "B" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "B" And severityindex = “5000” Then
            basedf = "5"

        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "C" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "C" And severityindex = "500" Then
            basedf = "10"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "C" And severityindex = “5000” Then
            basedf = "125"

        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "D" And severityindex = "50" Then
            basedf = "5"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "D" And severityindex = "500" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "D" And severityindex = “5000” Then
            basedf = "500"

        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown7.Value = 5 And ComboBox119.Text = "E" And severityindex = “5000” Then
            basedf = “5000”


        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "A" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "A" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "A" And severityindex = “5000” Then
            basedf = "1"

        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "B" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "B" And severityindex = "500" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "B" And severityindex = “5000” Then
            basedf = "2"

        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "C" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "C" And severityindex = "500" Then
            basedf = "5"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "C" And severityindex = “5000” Then
            basedf = "50"

        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "D" And severityindex = "50" Then
            basedf = "1"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "D" And severityindex = "500" Then
            basedf = "25"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "D" And severityindex = “5000” Then
            basedf = "250"

        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "E" And severityindex = "50" Then
            basedf = "50"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "E" And severityindex = "500" Then
            basedf = "500"
        ElseIf NumericUpDown7.Value = 6 And ComboBox119.Text = "E" And severityindex = “5000” Then
            basedf = “5000”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = basedf * (maxage1 ^ 1.1)
            Label373.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = basedf * (maxage2 ^ 1.1)
            Label373.Text = totaldf
        End If
    End Sub

    'Scc – Hydrogen Stress Cracking-Hf Coding perhitungan -----------------------------------------------------------------
    Private Sub susceptibilityscchydrogenschf()
        If ComboBox101.Text = "PWHT" Then
            If ComboBox102.Text = "< 200" Then
                ComboBox124.Text = "NONE SUSCEPTIBILITY"
            ElseIf ComboBox102.Text = "200-237" Then
                ComboBox124.Text = "LOW SUSCEPTIBILITY"
            ElseIf ComboBox102.Text = "> 237" Then
                ComboBox124.Text = "HIGH SUSCEPTIBILITY"
            End If
        ElseIf ComboBox101.Text = "As-Welded" Then
            If ComboBox102.Text = "< 200" Then
                ComboBox124.Text = "LOW SUSCEPTIBILITY"
            ElseIf ComboBox102.Text = "200-237" Then
                ComboBox124.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf ComboBox102.Text = "> 237" Then
                ComboBox124.Text = "HIGH SUSCEPTIBILITY"
            End If
        End If
    End Sub

    Public Sub scchydrogenschfdf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox123.Text = "A" Or ComboBox123.Text = "B" Or ComboBox123.Text = "C" AndAlso ComboBox24.Text = "No" Or ComboBox121.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox124.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 100
        ElseIf ComboBox124.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf ComboBox124.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 1
        ElseIf ComboBox124.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown8.Value = 1 And ComboBox123.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "A" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "B" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "C" And severityindex = “100” Then
            basedf = "33"

        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "D" And severityindex = “100” Then
            basedf = "80"

        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown8.Value = 1 And ComboBox123.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "B" And severityindex = “100” Then
            basedf = "4"

        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "C" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "D" And severityindex = “100” Then
            basedf = "60"

        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown8.Value = 2 And ComboBox123.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "B" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "C" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "D" And severityindex = “100” Then
            basedf = "40"

        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown8.Value = 3 And ComboBox123.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "C" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "D" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown8.Value = 4 And ComboBox123.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "C" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "D" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown8.Value = 5 And ComboBox123.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "C" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "D" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown8.Value = 6 And ComboBox123.Text = "E" And severityindex = “100” Then
            basedf = “100”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = basedf * (maxage1 ^ 1.1)
            Label374.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = basedf * (maxage2 ^ 1.1)
            Label374.Text = totaldf
        End If
    End Sub

    'Scc – Hic/Sohic-Hf Coding perhitungan --------------------------------------------------------------------------------
    Private Sub susceptibilityscchicsohichf()
        If ComboBox28.Text = "PWHT" Then
            If ComboBox163.Text = "High Sulfur Steel > 0.01% S" Then
                ComboBox165.Text = "HIGH SUSCEPTIBILITY"
            ElseIf ComboBox163.Text = "Low Sulfur Steel ≤ 0.01% S" Then
                ComboBox165.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf ComboBox163.Text = "Product Form – Seamless/Extruded Pipe" Then
                ComboBox165.Text = "LOW SUSCEPTIBILITY"
            End If
        ElseIf ComboBox28.Text = "Non-PWHT" Then
            If ComboBox163.Text = "High Sulfur Steel > 0.01% S" Then
                ComboBox165.Text = "HIGH SUSCEPTIBILITY"
            ElseIf ComboBox163.Text = "Low Sulfur Steel ≤ 0.01% S" Then
                ComboBox165.Text = "HIGH SUSCEPTIBILITY"
            ElseIf ComboBox163.Text = "Product Form – Seamless/Extruded Pipe" Then
                ComboBox165.Text = "LOW SUSCEPTIBILITY"
            End If
        End If
    End Sub

    Public Sub scchicsohichfdf()
        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim dateforcalculation As Double

        If ComboBox164.Text = "A" Or ComboBox164.Text = "B" Or ComboBox164.Text = "C" AndAlso ComboBox26.Text = "No" Or ComboBox27.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            dateforcalculation = datelastforgood
        Else
            dateforcalculation = datestartforpoor
        End If

        Dim severityindex As Double

        If ComboBox165.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 100
        ElseIf ComboBox165.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf ComboBox165.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 1
        ElseIf ComboBox165.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        End If

        Dim basedf As Double

        If NumericUpDown9.Value = 1 And ComboBox164.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "A" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "B" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "C" And severityindex = “100” Then
            basedf = "33"

        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "D" And severityindex = “100” Then
            basedf = "80"

        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown9.Value = 1 And ComboBox164.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "B" And severityindex = “100” Then
            basedf = "4"

        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "C" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "D" And severityindex = “100” Then
            basedf = "60"

        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown9.Value = 2 And ComboBox164.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "B" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "C" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "D" And severityindex = “100” Then
            basedf = "40"

        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown9.Value = 3 And ComboBox164.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "C" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "D" And severityindex = “100” Then
            basedf = "20"

        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown9.Value = 4 And ComboBox164.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "C" And severityindex = “100” Then
            basedf = "2"

        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "D" And severityindex = “100” Then
            basedf = "10"

        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown9.Value = 5 And ComboBox164.Text = "E" And severityindex = “100” Then
            basedf = “100”


        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "A" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "B" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "C" And severityindex = “100” Then
            basedf = "1"

        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "D" And severityindex = “100” Then
            basedf = "5"

        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown9.Value = 6 And ComboBox164.Text = "E" And severityindex = “100” Then
            basedf = “100”
        End If

        Dim onlineadjfactor As Double

        If ComboBox168.Text = "Key Process Variables" Then
            onlineadjfactor = 2
        ElseIf ComboBox168.Text = "Hydrogen Probes" Then
            onlineadjfactor = 2
        ElseIf ComboBox168.Text = "Key Process Variables and Hydrogen Probes" Then
            onlineadjfactor = 4
        ElseIf ComboBox168.Text = "No Online Monitoring Used" Then
            onlineadjfactor = 1
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim totaldf As Double

        If dateforcalculation > 1 Then
            maxage1 = dateforcalculation
            totaldf = (basedf * (maxage1 ^ 1.1)) / onlineadjfactor
            Label375.Text = totaldf
        ElseIf dateforcalculation <= 1 Then
            maxage2 = 1
            totaldf = (basedf * (maxage2 ^ 1.1)) / onlineadjfactor
            Label375.Text = totaldf
        End If
    End Sub

    'External Corrosion – Ferritic Component Coding perhitungan -----------------------------------------------------------
    Public Sub externalcorrosionferritic()
        Dim basecr As Double

        If Units.ComboBox1.Text = "SI" Then
            Dim chart2(,) As Double = {{-12, 0, 0, 0, 0},
            {-8, 0.025, 0, 0, 0},
            {6, 0.127, 0.076, 0.025, 0.254},
            {32, 0.127, 0.076, 0.025, 0.254},
            {71, 0.127, 0.051, 0.025, 0.254},
            {107, 0.025, 0, 0, 0.051},
            {121, 0, 0, 0, 0}}

            Dim UpNum2(4) As Double
            Dim LowNum2(4) As Double

            Dim Lo1 As Double
            Dim Lo2 As Double
            Dim up1 As Double
            Dim up2 As Double

            Dim temp As Double
            Dim upt As Double
            Dim lot As Double
            Dim x As Double
            Dim z As Double


            temp = Val(TextBox5.Text)


            If ComboBox138.Text = "Marine / Cooling Tower Drift Area" Then
                If temp < -12 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= -8 AndAlso chart2(r + 1, 0) >= -12 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = -8
                                lot = -12

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 121 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 107 AndAlso chart2(r + 1, 0) >= 121 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 121
                                lot = 107

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox138.Text = "Temperate" Then
                If temp < -12 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= -8 AndAlso chart2(r + 1, 0) >= -12 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = -8
                                lot = -12

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 121 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 107 AndAlso chart2(r + 1, 0) >= 121 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = 121
                                lot = 107

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox138.Text = "Arid / Dry" Then
                If temp < -12 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= -8 AndAlso chart2(r + 1, 0) >= -12 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = -8
                                lot = -12

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 121 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 107 AndAlso chart2(r + 1, 0) >= 121 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = 121
                                lot = 107

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox138.Text = "Severe" Then
                If temp < -12 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= -8 AndAlso chart2(r + 1, 0) >= -12 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(4)
                                z = UpNum2(4)
                                upt = -8
                                lot = -12

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 121 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 107 AndAlso chart2(r + 1, 0) >= 121 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(4)
                                z = UpNum2(4)
                                upt = 121
                                lot = 107

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            Dim chart2(,) As Double = {{10, 0, 0, 0, 0},
            {18, 1, 0, 0, 3},
            {43, 5, 3, 1, 10},
            {90, 5, 3, 1, 10},
            {160, 5, 2, 1, 10},
            {225, 1, 0, 0, 2},
            {250, 0, 0, 0, 0}}

            Dim UpNum2(4) As Double
            Dim LowNum2(4) As Double

            Dim Lo1 As Double
            Dim Lo2 As Double
            Dim up1 As Double
            Dim up2 As Double

            Dim temp As Double
            Dim upt As Double
            Dim lot As Double
            Dim x As Double
            Dim z As Double


            temp = Val(TextBox5.Text)


            If ComboBox138.Text = "Marine / Cooling Tower Drift Area" Then
                If temp < 10 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 18 AndAlso chart2(r + 1, 0) >= 10 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 18
                                lot = 10

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 250 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 225 AndAlso chart2(r + 1, 0) >= 250 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 250
                                lot = 225

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox138.Text = "Temperate" Then
                If temp < 10 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 18 AndAlso chart2(r + 1, 0) >= 10 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = 18
                                lot = 10

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 250 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 225 AndAlso chart2(r + 1, 0) >= 250 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(2)
                                z = UpNum2(2)
                                upt = 250
                                lot = 225

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox138.Text = "Arid / Dry" Then
                If temp < 10 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 18 AndAlso chart2(r + 1, 0) >= 10 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = 18
                                lot = 10

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 250 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 225 AndAlso chart2(r + 1, 0) >= 250 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(3)
                                z = UpNum2(3)
                                upt = 250
                                lot = 225

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If

            If ComboBox138.Text = "Severe" Then
                If temp < 10 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 18 AndAlso chart2(r + 1, 0) >= 10 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(4)
                                z = UpNum2(4)
                                upt = 18
                                lot = 10

                                basecr = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 250 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 225 AndAlso chart2(r + 1, 0) >= 250 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(4)
                                z = UpNum2(4)
                                upt = 250
                                lot = 225

                                basecr = Inter3c(upt, lot, temp, x, z)

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

                            basecr = Inter(temp, up1, up2, Lo1, Lo2)


                        Next
                        Exit For
                    End If
                Next
            End If
        End If

        Dim feq As Double
        If ComboBox142.Text = "Yes" Then
            feq = 2
        ElseIf ComboBox142.Text = "No" Then
            feq = 1
        End If

        Dim fif As Double
        If ComboBox139.Text = "Yes" Then
            fif = 2
        ElseIf ComboBox139.Text = "No" Then
            fif = 1
        End If

        Dim cr As Double

        If feq > fif Then
            cr = basecr * feq
        End If
        If feq < fif Then
            cr = basecr * fif
        End If
        If feq = fif Then
            cr = basecr * feq
        End If

        Dim p As Double
        Dim R1 As Double
        Dim s As Double
        Dim ejo As Double
        Dim ca As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim t3 As Double
        Dim tpipe As Double
        Dim y As Double
        Dim ttanksi As Double
        Dim ttankus As Double
        Dim d As Double
        Dim h As Double
        Dim gsi As Double

        gsi = Val(TextBox59.Text)

        p = Val(TextBox54.Text) * 0.0980665
        R1 = Val(TextBox53.Text)
        s = Val(TextBox6.Text)
        ejo = Val(TextBox7.Text)
        ca = Val(TextBox55.Text)
        d = Val(TextBox41.Text)
        h = Val(TextBox40.Text)


        If ComboBox4.Text = "Pipe" Then
            If RadioButton23.Checked = True Then
                If ComboBox3.Text = "Ferritic Steels" Then
                    Dim chart2(,) As Double = {{900, 0.4},
                {950, 0.5},
                {1000, 0.7},
                {1050, 0.7},
                {1100, 0.7},
                {1150, 0.7},
                {1200, 0.7},
                {1250, 0.7}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim temp As Double

                    temp = Val(TextBox14.Text)

                    If temp < 900 Then
                        y = 0.4

                    ElseIf temp > 1250 Then
                        y = 0.7
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

                                y = Inter(temp, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox3.Text = "Austenitic Steels" Then
                    Dim chart2(,) As Double = {{900, 0.4},
                {950, 0.4},
                {1000, 0.4},
                {1050, 0.4},
                {1100, 0.5},
                {1150, 0.7},
                {1200, 0.7},
                {1250, 0.7}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim temp As Double

                    temp = Val(TextBox14.Text)

                    If temp < 900 Then
                        y = 0.4
                    ElseIf temp > 1250 Then
                        y = 0.7
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

                                y = Inter(temp, up1, up2, Lo1, Lo2)

                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox3.Text = "Nickel Alloys UNS Nos. N06617 , N08800, N08810, N08825" Then
                    Dim chart2(,) As Double = {{900, 0.4},
                {950, 0.4},
                {1000, 0.4},
                {1050, 0.4},
                {1100, 0.5},
                {1150, 0.7},
                {1200, 0.5},
                {1250, 0.7}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim temp As Double

                    temp = Val(TextBox14.Text)

                    If temp < 900 Then
                        y = 0.4
                    ElseIf temp > 1250 Then
                        y = 0.7
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

                                y = Inter(temp, up1, up2, Lo1, Lo2)

                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox3.Text = "Cast Iron" Then
                    y = 0
                End If

                If ComboBox3.Text = "Nonferrous" Then
                    y = 0
                End If
            End If

            If RadioButton24.Checked = True Then
                If ComboBox3.Text = "Ferritic Steels" Then
                    Dim chart2(,) As Double = {{482, 0.4},
                {510, 0.5},
                {538, 0.7},
                {566, 0.7},
                {593, 0.7},
                {621, 0.7},
                {649, 0.7},
                {677, 0.7}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim temp As Double

                    temp = Val(TextBox14.Text)

                    If temp < 482 Then
                        y = 0.4
                    ElseIf temp > 677 Then
                        y = 0.7
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

                                y = Inter(temp, up1, up2, Lo1, Lo2)

                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox3.Text = "Austenitic Steels" Then
                    Dim chart2(,) As Double = {{482, 0.4},
                {510, 0.4},
                {538, 0.4},
                {566, 0.4},
                {593, 0.5},
                {621, 0.7},
                {649, 0.7},
                {677, 0.7}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim temp As Double

                    temp = Val(TextBox14.Text)

                    If temp < 482 Then
                        y = 0.4
                    ElseIf temp > 677 Then
                        y = 0.7
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

                                y = Inter(temp, up1, up2, Lo1, Lo2)

                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox3.Text = "Nickel Alloys UNS Nos. N06617 , N08800, N08810, N08825" Then
                    Dim chart2(,) As Double = {{482, 0.4},
                {510, 0.4},
                {538, 0.4},
                {566, 0.4},
                {593, 0.4},
                {621, 0.4},
                {649, 0.5},
                {677, 0.7}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim temp As Double

                    temp = Val(TextBox14.Text)

                    If temp < 482 Then
                        y = 0.4
                    ElseIf temp > 677 Then
                        y = 0.7
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

                                y = Inter(temp, up1, up2, Lo1, Lo2)

                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox3.Text = "Cast Iron" Then
                    y = 0
                End If

                If ComboBox3.Text = "Nonferrous" Then
                    y = 0
                End If
            End If


            tpipe = ((p * R1) / ((s * ejo) + (p * y))) + ca
            TextBox8.Text = tpipe
        End If

        If ComboBox4.Text = "AST (Shell Course)" Then
            If RadioButton21.Checked = True Then
                ttanksi = (((4.9 * d) * (h - 0.3) * gsi) / s) + ca
                TextBox8.Text = ttanksi
            ElseIf RadioButton22.Checked = True Then
                ttankus = (((2.6 * d) * (h - 1) * gsi) / s) + ca
                TextBox8.Text = ttankus
            End If
        End If

        If ComboBox4.Text = "Other Components" Then
            t1 = ((p * R1) / ((s * ejo) - (0.6 * p))) + ca
            t2 = ((p * R1) / ((2 * s * ejo) + (0.4 * p))) + ca
            t3 = ((p * R1) / ((2 * s * ejo) - (0.2 * p))) + ca

            If t1 > t2 And t1 > t3 Then
                TextBox8.Text = t1
            ElseIf t2 > t1 And t2 > t3 Then
                TextBox8.Text = t2
            ElseIf t3 > t1 And t3 > t2 Then
                TextBox8.Text = t3

            ElseIf t1 = t2 And t1 = t3 Then
                TextBox8.Text = t1

            ElseIf t1 >= t2 And t1 >= t3 Then
                TextBox8.Text = t1
            ElseIf t2 >= t1 And t2 >= t3 Then
                TextBox8.Text = t2
            ElseIf t3 >= t2 And t3 >= t1 Then
                TextBox8.Text = t3
            End If
        End If

        Dim trde As Double

        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim agetk As Double

        If ComboBox141.Text = "A" Or ComboBox141.Text = "B" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

            trde = Val(TextBox2.Text)
        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

            trde = Val(TextBox1.Text)
        End If

        If datelastforgood > 0 Then
            agetk = datelastforgood
        Else
            agetk = datestartforpoor
        End If

        Dim datecoat As Double

        Dim agecoat As Double

        datecoat = DateDiff(DateInterval.DayOfYear, DateTimePicker11.Value, DateTimePicker13.Value)

        agecoat = datecoat / 365

        Dim coatadj As Double

        If agetk < agecoat Then

            If ComboBox120.Text = "No Coating or Poor – No coating or primer only" Then
                coatadj = "0"
            End If

            If ComboBox120.Text = "Medium – Single coat epoxy" Then
                If 5 < agecoat Then
                    If 5 < (agecoat - agetk) Then
                        coatadj = 5 - 5
                    End If
                    If 5 > (agecoat - agetk) Then
                        coatadj = 5 - (agecoat - agetk)
                    End If
                    If 5 = (agecoat - agetk) Then
                        coatadj = "0"
                    End If
                End If
                If 5 > agecoat Then
                    If 5 < (agecoat - agetk) Then
                        coatadj = agecoat - 5
                    End If
                    If 5 > (agecoat - agetk) Then
                        coatadj = agecoat - (agecoat - agetk)
                    End If
                    If 5 = (agecoat - agetk) Then
                        coatadj = agecoat - 5
                    End If
                End If
            End If

            If ComboBox120.Text = "High – Multi coat epoxy or filled epoxy" Then
                If 15 < agecoat Then
                    If 15 < (agecoat - agetk) Then
                        coatadj = 15 - 15
                    End If
                    If 15 > (agecoat - agetk) Then
                        coatadj = 15 - (agecoat - agetk)
                    End If
                    If 15 = (agecoat - agetk) Then
                        coatadj = "0"
                    End If
                End If
                If 15 > agecoat Then
                    If 15 < (agecoat - agetk) Then
                        coatadj = agecoat - 15
                    End If
                    If 15 > (agecoat - agetk) Then
                        coatadj = agecoat - (agecoat - agetk)
                    End If
                    If 15 = (agecoat - agetk) Then
                        coatadj = agecoat - 15
                    End If
                End If
            End If
        End If

        If agetk >= agecoat Then
            If ComboBox120.Text = "No Coating or Poor – No coating or primer only" Then
                coatadj = "0"
            End If

            If ComboBox120.Text = "Medium – Single coat epoxy" Then
                If 5 < agecoat Then
                    coatadj = "5"
                Else
                    coatadj = agecoat
                End If
            End If

            If ComboBox120.Text = "High – Multi coat epoxy or filled epoxy" Then
                If 15 < agecoat Then
                    coatadj = "15"
                Else
                    coatadj = agecoat
                End If
            End If
        End If

        Dim age As Double

        age = agetk - coatadj

        Dim Art As Double

        Art = (cr * age) / trde

        Dim fs As Double

        Dim ys As Double
        Dim ts As Double
        Dim wje As Double

        ys = Val(TextBox13.Text)
        ts = Val(TextBox14.Text)
        wje = Val(TextBox8.Text)

        fs = ((ys + ts) / 2) * wje * 1.1

        'SRP method 1
        Dim tmin As Double
        Dim tc As Double
        Dim trdi As Double

        tmin = Val(TextBox8.Text)
        tc = Val(TextBox85.Text)
        trdi = Val(TextBox49.Text)

        Dim srp As Double

        If CheckBox1.Checked = False Then
            srp = ((s * wje) / fs) * (tmin / trdi)

        ElseIf CheckBox1.Checked = True Then
            srp = ((s * wje) / fs) * (tc / trdi)

        End If

        'SRP method 2
        If ComboBox135.Text = "Cylinder" Then
            srp = (p * d) / (2 * fs * trdi)

        ElseIf ComboBox135.Text = "Sphere" Then
            srp = (p * d) / (4 * fs * trdi)

        ElseIf ComboBox135.Text = "Head" Then
            srp = (p * d) / (1.13 * fs * trdi)

        End If

        'prior probability
        Dim prp1 As Double
        Dim prp2 As Double
        Dim prp3 As Double

        If ComboBox137.Text = "Low" Then
            prp1 = 0.5
            prp2 = 0.3
            prp3 = 0.2
        ElseIf ComboBox137.Text = "Medium" Then
            prp1 = 0.7
            prp2 = 0.2
            prp3 = 0.1
        ElseIf ComboBox137.Text = "High" Then
            prp1 = 0.8
            prp2 = 0.15
            prp3 = 0.05
        End If

        'inspection eff factor
        Dim cop1a As Double = 0.9
        Dim cop1b As Double = 0.7
        Dim cop1c As Double = 0.5
        Dim cop1d As Double = 0.4
        Dim cop1e As Double = 0.33

        Dim cop2a As Double = 0.09
        Dim cop2b As Double = 0.2
        Dim cop2c As Double = 0.3
        Dim cop2d As Double = 0.33
        Dim cop2e As Double = 0.33

        Dim cop3a As Double = 0.01
        Dim cop3b As Double = 0.1
        Dim cop3c As Double = 0.2
        Dim cop3d As Double = 0.27
        Dim cop3e As Double = 0.33

        Dim na As Double
        Dim nb As Double
        Dim nc As Double
        Dim nd As Double
        Dim ne As Double

        If ComboBox136.Text = "A" Then
            na = NumericUpDown13.Value
            nb = 0
            nc = 0
            nd = 0
            ne = 0
        ElseIf ComboBox136.Text = "B" Then
            na = 0
            nb = NumericUpDown13.Value
            nc = 0
            nd = 0
            ne = 0
        ElseIf ComboBox136.Text = "C" Then
            na = 0
            nb = 0
            nc = NumericUpDown13.Value
            nd = 0
            ne = 0
        ElseIf ComboBox136.Text = "D" Then
            na = 0
            nb = 0
            nc = 0
            nd = NumericUpDown13.Value
            ne = 0
        ElseIf ComboBox136.Text = "E" Then
            na = 0
            nb = 0
            nc = 0
            nd = 0
            ne = NumericUpDown13.Value
        End If

        Dim i1 As Double
        Dim i2 As Double
        Dim i3 As Double

        i1 = prp1 * (cop1a ^ na) * (cop1b ^ nb) * (cop1c ^ nc) * (cop1d ^ nd) * (cop1e ^ ne)
        i2 = prp2 * (cop2a ^ na) * (cop2b ^ nb) * (cop2c ^ nc) * (cop2d ^ nd) * (cop2e ^ ne)
        i3 = prp3 * (cop3a ^ na) * (cop3b ^ nb) * (cop3c ^ nc) * (cop3d ^ nd) * (cop3e ^ ne)

        Dim po1 As Double
        Dim po2 As Double
        Dim po3 As Double

        po1 = i1 / (i1 + i2 + i3)
        po2 = i2 / (i1 + i2 + i3)
        po3 = i3 / (i1 + i2 + i3)

        'reliability indices

        Dim ds1 As Double = 1
        Dim covt As Double = 0.2
        Dim covsf As Double = 0.2
        Dim covp As Double = 0.05
        Dim ds2 As Double = 2
        Dim ds3 As Double = 4

        Dim reliabilityindices1 As Double
        Dim reliabilityindices2 As Double
        Dim reliabilityindices3 As Double

        reliabilityindices1 = (1 - (ds1 * Art) - srp) / ((((ds1 ^ 2) * (Art ^ 2) * (covt ^ 2)) + (((1 - (ds1 * Art)) ^ 2) * (covsf ^ 2)) + ((srp ^ 2) * (covp ^ 2))) ^ 0.5)
        reliabilityindices2 = (1 - (ds2 * Art) - srp) / ((((ds2 ^ 2) * (Art ^ 2) * (covt ^ 2)) + (((1 - (ds2 * Art)) ^ 2) * (covsf ^ 2)) + ((srp ^ 2) * (covp ^ 2))) ^ 0.5)
        reliabilityindices3 = (1 - (ds3 * Art) - srp) / ((((ds3 ^ 2) * (Art ^ 2) * (covt ^ 2)) + (((1 - (ds3 * Art)) ^ 2) * (covsf ^ 2)) + ((srp ^ 2) * (covp ^ 2))) ^ 0.5)

        'final df
        Dim norm4 As Double
        Dim norm5 As Double
        Dim norm6 As Double
        Dim x1 As Double
        Dim x2 As Double
        Dim x3 As Double

        x1 = -1 * reliabilityindices1
        x2 = -1 * reliabilityindices2
        x3 = -1 * reliabilityindices3

        norm4 = Cumnorm(x1)
        norm5 = Cumnorm(x2)
        norm6 = Cumnorm(x3)

        Dim df As Double

        df = ((po1 * norm4) + (po2 * norm5) + (po3 * norm6)) / 0.000156

        Label360.Text = df
    End Sub

    'External Chloride Stress Corrosion Cracking – Austenitic Component Coding perhitungan --------------------------------
    Private Sub susceptibilityexternalsccaustenitic()
        If Units.ComboBox1.Text = "SI" Then
            If TextBox5.Text < 49 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 49 And ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 49 And ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 49 And ComboBox29.Text = "Severe" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Severe" Then
                TextBox39.Text = "HIGH SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Severe" Then
                TextBox39.Text = "MEDIUM SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Severe" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If TextBox5.Text < 120 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 120 And ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 120 And ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 120 And ComboBox29.Text = "Severe" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Severe" Then
                TextBox39.Text = "HIGH SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Severe" Then
                TextBox39.Text = "MEDIUM SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Temperate" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Arid" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Severe" Then
                TextBox39.Text = "NONE SUSCEPTIBILITY"
            End If
        End If

    End Sub

    Public Sub externalsccausteniticdf()
        Dim severityindex As Double

        If TextBox39.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        ElseIf TextBox39.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 1
        ElseIf TextBox39.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf TextBox39.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 50
        End If

        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim crack As Double

        If ComboBox127.Text = "A" Or ComboBox127.Text = "B" AndAlso ComboBox30.Text = "No" Or ComboBox31.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            crack = datelastforgood
        Else
            crack = datestartforpoor
        End If


        Dim datecoat As Double

        Dim coat As Double

        datecoat = DateDiff(DateInterval.DayOfYear, DateTimePicker4.Value, DateTimePicker13.Value)

        coat = datecoat / 365

        Dim coatadj As Double

        Dim coatadj1 As Double
        Dim coatadj2 As Double

        If coat < 5 Then
            coatadj1 = coat
        ElseIf coat > 5 Then
            coatadj1 = 5
        End If

        If coat < 15 Then
            coatadj2 = coat
        ElseIf coat > 15 Then
            coatadj2 = 15
        End If

        If crack >= coat And ComboBox128.Text = "No Coating or Poor Coating Quality" Then
            coatadj = "0"
        ElseIf crack >= coat And ComboBox128.Text = "Medium Coating Quality" Then
            coatadj = coatadj1
        ElseIf crack >= coat And ComboBox128.Text = "High Coating Quality" Then
            coatadj = coatadj2
        End If

        Dim min1 As Double
        Dim min2 As Double
        Dim min3 As Double
        Dim min4 As Double

        If coat < 15 Then
            min3 = coat
        ElseIf coat > 15 Then
            min3 = 15
        End If

        If (coat - crack) < 15 Then
            min4 = (coat - crack)
        ElseIf (coat - crack) > 15 Then
            min4 = 15
        End If

        If coat < 5 Then
            min1 = coat
        ElseIf coat > 5 Then
            min1 = 5
        End If

        If (coat - crack) < 5 Then
            min2 = (coat - crack)
        ElseIf (coat - crack) > 5 Then
            min2 = 5
        End If

        If crack < coat And ComboBox128.Text = "No Coating or Poor Coating Quality" Then
            coatadj = "0"
        ElseIf crack < coat And ComboBox128.Text = "Medium Coating Quality" Then
            coatadj = min1 - min2
        ElseIf crack < coat And ComboBox128.Text = "High Coating Quality" Then
            coatadj = min3 - min4
        End If

        Dim age As Double

        age = crack - coatadj

        Dim basedf As Double

        If NumericUpDown10.Value = 1 And ComboBox127.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "A" And severityindex = “50” Then
            basedf = "3"

        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "B" And severityindex = “50” Then
            basedf = "5"

        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "C" And severityindex = “50” Then
            basedf = "17"

        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "D" And severityindex = “50” Then
            basedf = "40"

        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown10.Value = 1 And ComboBox127.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "B" And severityindex = “50” Then
            basedf = "2"

        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "C" And severityindex = “50” Then
            basedf = "10"

        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "D" And severityindex = “50” Then
            basedf = "30"

        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown10.Value = 2 And ComboBox127.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "C" And severityindex = “50” Then
            basedf = "5"

        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "D" And severityindex = “50” Then
            basedf = "20"

        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown10.Value = 3 And ComboBox127.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "C" And severityindex = “50” Then
            basedf = "2"

        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "D" And severityindex = “50” Then
            basedf = "10"

        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown10.Value = 4 And ComboBox127.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "C" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "D" And severityindex = “50” Then
            basedf = "5"

        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown10.Value = 5 And ComboBox127.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "C" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "D" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown10.Value = 6 And ComboBox127.Text = "E" And severityindex = “50” Then
            basedf = “50”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim df As Double

        If age > 1 Then
            maxage1 = age
            df = basedf * (maxage1 ^ 1.1)
            Label362.Text = df
        ElseIf age <= 1 Then
            maxage2 = 1
            df = basedf * (maxage2 ^ 1.1)
            Label362.Text = df
        End If

    End Sub

    'External Chloride Stress Corrosion Cracking Under Insulation – Austenitic Component Coding perhitungan ---------------
    Private Sub pipingcomplexity()
        If ComboBox131.Text = "Below Average" Then
            If TextBox40.Text = "LOW SUSCEPTIBILITY" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "MEDIUM SUSCEPTIBILITY" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "HIGH SUSCEPTIBILITY" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"
            End If
        ElseIf ComboBox131.Text = "Average" Then

        ElseIf ComboBox131.Text = "Above Average" Then
            If TextBox40.Text = "NONE SUSCEPTIBILITY" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "LOW SUSCEPTIBILITY" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "MEDIUM SUSCEPTIBILITY" Then
                TextBox40.Text = "HIGH SUSCEPTIBILITY"
            End If
        End If
    End Sub

    Private Sub insulationcondition()
        If ComboBox132.Text = "Below Average" Then
            If TextBox40.Text = "LOW SUSCEPTIBILITY" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "MEDIUM SUSCEPTIBILITY" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "HIGH SUSCEPTIBILITY" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"
            End If
        ElseIf ComboBox132.Text = "Average" Then

        ElseIf ComboBox132.Text = "Above Average" Then
            If TextBox40.Text = "NONE SUSCEPTIBILITY" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "LOW SUSCEPTIBILITY" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf TextBox40.Text = "MEDIUM SUSCEPTIBILITY" Then
                TextBox40.Text = "HIGH SUSCEPTIBILITY"
            End If
        End If
    End Sub

    Private Sub susceptibilityexternalscccuiaustenitic()
        If Units.ComboBox1.Text = "SI" Then
            If TextBox5.Text < 49 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 49 And ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 49 And ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 49 And ComboBox29.Text = "Severe" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 49 AndAlso TextBox5.Text < 93 AndAlso ComboBox29.Text = "Severe" Then
                TextBox40.Text = "HIGH SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 93 AndAlso TextBox5.Text < 149 AndAlso ComboBox29.Text = "Severe" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 149 And ComboBox29.Text = "Severe" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If TextBox5.Text < 120 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 120 And ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 120 And ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text < 120 And ComboBox29.Text = "Severe" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 120 AndAlso TextBox5.Text < 200 AndAlso ComboBox29.Text = "Severe" Then
                TextBox40.Text = "HIGH SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "LOW SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 200 AndAlso TextBox5.Text < 300 AndAlso ComboBox29.Text = "Severe" Then
                TextBox40.Text = "MEDIUM SUSCEPTIBILITY"

            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Marine / Cooling Tower Drift Area" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Temperate" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Arid" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            ElseIf TextBox5.Text >= 300 And ComboBox29.Text = "Severe" Then
                TextBox40.Text = "NONE SUSCEPTIBILITY"
            End If
        End If

    End Sub

    Public Sub externalscccuiausteniticdf()
        Dim severityindex As Double

        If TextBox40.Text = "NONE SUSCEPTIBILITY" Then
            severityindex = 0
        ElseIf TextBox40.Text = "LOW SUSCEPTIBILITY" Then
            severityindex = 1
        ElseIf TextBox40.Text = "MEDIUM SUSCEPTIBILITY" Then
            severityindex = 10
        ElseIf TextBox40.Text = "HIGH SUSCEPTIBILITY" Then
            severityindex = 50
        End If

        Dim date1 As Double
        Dim date2 As Double

        Dim datelastforgood As Double
        Dim datestartforpoor As Double

        Dim crack As Double

        If ComboBox130.Text = "A" Or ComboBox130.Text = "B" AndAlso ComboBox32.Text = "No" Or ComboBox129.Text = "Yes" Then

            date1 = DateDiff(DateInterval.DayOfYear, DateTimePicker1.Value, DateTimePicker13.Value)
            datelastforgood = date1 / 365

        Else

            date2 = DateDiff(DateInterval.DayOfYear, DateTimePicker2.Value, DateTimePicker13.Value)
            datestartforpoor = date2 / 365

        End If

        If datelastforgood > 0 Then
            crack = datelastforgood
        Else
            crack = datestartforpoor
        End If


        Dim datecoat As Double

        Dim coat As Double

        datecoat = DateDiff(DateInterval.DayOfYear, DateTimePicker4.Value, DateTimePicker13.Value)

        coat = datecoat / 365

        Dim coatadj As Double

        Dim coatadj1 As Double
        Dim coatadj2 As Double

        If coat < 5 Then
            coatadj1 = coat
        ElseIf coat > 5 Then
            coatadj1 = 5
        End If

        If coat < 15 Then
            coatadj2 = coat
        ElseIf coat > 15 Then
            coatadj2 = 15
        End If

        If crack >= coat And ComboBox128.Text = "No Coating or Poor Coating Quality" Then
            coatadj = "0"
        ElseIf crack >= coat And ComboBox128.Text = "Medium Coating Quality" Then
            coatadj = coatadj1
        ElseIf crack >= coat And ComboBox128.Text = "High Coating Quality" Then
            coatadj = coatadj2
        End If

        Dim min1 As Double
        Dim min2 As Double
        Dim min3 As Double
        Dim min4 As Double

        If coat < 15 Then
            min3 = coat
        ElseIf coat > 15 Then
            min3 = 15
        End If

        If (coat - crack) < 15 Then
            min4 = (coat - crack)
        ElseIf (coat - crack) > 15 Then
            min4 = 15
        End If

        If coat < 5 Then
            min1 = coat
        ElseIf coat > 5 Then
            min1 = 5
        End If

        If (coat - crack) < 5 Then
            min2 = (coat - crack)
        ElseIf (coat - crack) > 5 Then
            min2 = 5
        End If

        If crack < coat And ComboBox128.Text = "No Coating or Poor Coating Quality" Then
            coatadj = "0"
        ElseIf crack < coat And ComboBox128.Text = "Medium Coating Quality" Then
            coatadj = min1 - min2
        ElseIf crack < coat And ComboBox128.Text = "High Coating Quality" Then
            coatadj = min3 - min4
        End If

        Dim age As Double

        age = crack - coatadj

        Dim basedf As Double

        If NumericUpDown11.Value = 1 And ComboBox130.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "A" And severityindex = “50” Then
            basedf = "3"

        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "B" And severityindex = “50” Then
            basedf = "5"

        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "C" And severityindex = "10" Then
            basedf = "3"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "C" And severityindex = “50” Then
            basedf = "17"

        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "D" And severityindex = "10" Then
            basedf = "8"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "D" And severityindex = “50” Then
            basedf = "40"

        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown11.Value = 1 And ComboBox130.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "B" And severityindex = “50” Then
            basedf = "2"

        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "C" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "C" And severityindex = “50” Then
            basedf = "10"

        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "D" And severityindex = "10" Then
            basedf = "6"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "D" And severityindex = “50” Then
            basedf = "30"

        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown11.Value = 2 And ComboBox130.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "C" And severityindex = “50” Then
            basedf = "5"

        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "D" And severityindex = "10" Then
            basedf = "4"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "D" And severityindex = “50” Then
            basedf = "20"

        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown11.Value = 3 And ComboBox130.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "C" And severityindex = “50” Then
            basedf = "2"

        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "D" And severityindex = "10" Then
            basedf = "2"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "D" And severityindex = “50” Then
            basedf = "10"

        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown11.Value = 4 And ComboBox130.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "C" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "D" And severityindex = “50” Then
            basedf = "5"

        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown11.Value = 5 And ComboBox130.Text = "E" And severityindex = “50” Then
            basedf = “50”


        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "A" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "A" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "A" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "A" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "B" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "B" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "B" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "B" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "C" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "C" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "C" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "C" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "D" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "D" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "D" And severityindex = "10" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "D" And severityindex = “50” Then
            basedf = "1"

        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "E" And severityindex = "0" Then
            basedf = "0"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "E" And severityindex = "1" Then
            basedf = "1"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "E" And severityindex = "10" Then
            basedf = "10"
        ElseIf NumericUpDown11.Value = 6 And ComboBox130.Text = "E" And severityindex = “50” Then
            basedf = “50”
        End If

        Dim maxage1 As Double
        Dim maxage2 As Double
        Dim df As Double

        If age > 1 Then
            maxage1 = age
            df = basedf * (maxage1 ^ 1.1)
            Label363.Text = df
        ElseIf age <= 1 Then
            maxage2 = 1
            df = basedf * (maxage2 ^ 1.1)
            Label363.Text = df
        End If

    End Sub

    'High Temperature Hydrogen Attack Coding perhitungan ------------------------------------------------------------------
    Public Sub hthadf()
        If TextBox45.Text = "Damage Observed" Then
            Label376.Text = 5000
        ElseIf TextBox45.Text = "High Susceptibility" Then
            Label376.Text = 5000
        ElseIf TextBox45.Text = "Medium Susceptibility" Then
            Label376.Text = 2000
        ElseIf TextBox45.Text = "Low Susceptibility" Then
            Label376.Text = 100
        ElseIf TextBox45.Text = "No Susceptibility" Then
            Label376.Text = 0
        End If
    End Sub

    'Brittle Fracture Coding perhitungan ----------------------------------------------------------------------------------
    Public Sub brittledf()

        Dim reftemp As Double

        If Units.ComboBox1.Text = "SI" Then
            If ComboBox1.Text = "Carbon Steel" Then
                If ComboBox157.Text = "A" Then
                    Dim chart(,) As Double = {{200, 42},
                            {210, 38},
                            {220, 36},
                            {230, 33},
                            {240, 31},
                            {260, 27},
                            {280, 24},
                            {300, 22},
                            {320, 19},
                            {340, 17},
                            {360, 15}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart.GetUpperBound(0) - 1
                            If chart(r, 0) <= 210 AndAlso chart(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart(r, c)
                                    UpNum2(c) = chart(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart.GetUpperBound(0) - 1
                            If chart(r, 0) <= 340 AndAlso chart(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart(r, c)
                                    UpNum2(c) = chart(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart.GetUpperBound(0) - 1
                        If chart(r, 0) <= mys AndAlso chart(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart(r, c)
                                UpNum2(c) = chart(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "B" Then
                    Dim chart10(,) As Double = {{200, 21},
                        {210, 17},
                        {220, 15},
                        {230, 1.25},
                        {240, 10},
                        {260, 6},
                        {280, 3},
                        {300, 1},
                        {320, -2},
                        {340, -4},
                        {360, -6}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart10.GetUpperBound(0) - 1
                            If chart10(r, 0) <= 210 AndAlso chart10(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart10(r, c)
                                    UpNum2(c) = chart10(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart10.GetUpperBound(0) - 1
                            If chart10(r, 0) <= 340 AndAlso chart10(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart10(r, c)
                                    UpNum2(c) = chart10(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart10.GetUpperBound(0) - 1
                        If chart10(r, 0) <= mys AndAlso chart10(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart10(r, c)
                                UpNum2(c) = chart10(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "C" Then
                    Dim chart11(,) As Double = {{200, 0},
                        {210, -4},
                        {220, -7},
                        {230, -9},
                        {240, -11},
                        {260, -15},
                        {280, -18},
                        {300, -21},
                        {320, -23},
                        {340, -25},
                        {360, -27}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 210 AndAlso chart11(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 340 AndAlso chart11(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= mys AndAlso chart11(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "D" Then
                    Dim chart12(,) As Double = {{200, -15},
                        {210, -18},
                        {220, -21},
                        {230, -23},
                        {240, -26},
                        {260, -29},
                        {280, -32},
                        {300, -35},
                        {320, -37},
                        {340, -39},
                        {360, -41}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart12.GetUpperBound(0) - 1
                            If chart12(r, 0) <= 210 AndAlso chart12(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart12(r, c)
                                    UpNum2(c) = chart12(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart12.GetUpperBound(0) - 1
                            If chart12(r, 0) <= 340 AndAlso chart12(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart12(r, c)
                                    UpNum2(c) = chart12(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart12.GetUpperBound(0) - 1
                        If chart12(r, 0) <= mys AndAlso chart12(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart12(r, c)
                                UpNum2(c) = chart12(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

            ElseIf ComboBox1.Text = "Low Alloy Steel" Then
                If ComboBox157.Text = "A" Then
                    Dim chart13(,) As Double = {{200, 55},
                        {210, 50},
                        {220, 46},
                        {230, 43},
                        {240, 40},
                        {250, 38},
                        {260, 36},
                        {270, 34},
                        {280, 32},
                        {290, 31},
                        {300, 30},
                        {310, 28},
                        {320, 27},
                        {330, 26},
                        {340, 25},
                        {360, 23},
                        {380, 21},
                        {400, 19},
                        {420, 18},
                        {440, 16},
                        {460, 15},
                        {480, 14},
                        {500, 13},
                        {520, 12},
                        {540, 11},
                        {560, 10}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart13.GetUpperBound(0) - 1
                            If chart13(r, 0) <= 210 AndAlso chart13(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart13(r, c)
                                    UpNum2(c) = chart13(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart13.GetUpperBound(0) - 1
                            If chart13(r, 0) <= 540 AndAlso chart13(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart13(r, c)
                                    UpNum2(c) = chart13(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart13.GetUpperBound(0) - 1
                        If chart13(r, 0) <= mys AndAlso chart13(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart13(r, c)
                                UpNum2(c) = chart13(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "B" Then
                    Dim chart14(,) As Double = {{200, 33},
                        {210, 29},
                        {220, 25},
                        {230, 22},
                        {240, 19},
                        {250, 17},
                        {260, 15},
                        {270, 13},
                        {280, 11},
                        {290, 10},
                        {300, 8},
                        {310, 7},
                        {320, 6},
                        {330, 5},
                        {340, 4},
                        {360, 2},
                        {380, 0},
                        {400, -2},
                        {420, -3},
                        {440, -5},
                        {460, -6},
                        {480, -7},
                        {500, -8},
                        {520, -9},
                        {540, -10},
                        {560, -11}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart14.GetUpperBound(0) - 1
                            If chart14(r, 0) <= 210 AndAlso chart14(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart14(r, c)
                                    UpNum2(c) = chart14(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart14.GetUpperBound(0) - 1
                            If chart14(r, 0) <= 540 AndAlso chart14(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart14(r, c)
                                    UpNum2(c) = chart14(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart14.GetUpperBound(0) - 1
                        If chart14(r, 0) <= mys AndAlso chart14(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart14(r, c)
                                UpNum2(c) = chart14(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "C" Then
                    Dim chart15(,) As Double = {{200, 12},
                        {210, 8},
                        {220, 4},
                        {230, 1},
                        {240, -2},
                        {250, -4},
                        {260, -6},
                        {270, -8},
                        {280, -10},
                        {290, -11},
                        {300, -13},
                        {310, -14},
                        {320, -15},
                        {330, -16},
                        {340, -17},
                        {360, -19},
                        {380, -21},
                        {400, -23},
                        {420, -24},
                        {440, -26},
                        {460, -27},
                        {480, -28},
                        {500, -29},
                        {520, -30},
                        {540, -31},
                        {560, -32}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart15.GetUpperBound(0) - 1
                            If chart15(r, 0) <= 210 AndAlso chart15(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart15(r, c)
                                    UpNum2(c) = chart15(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart15.GetUpperBound(0) - 1
                            If chart15(r, 0) <= 540 AndAlso chart15(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart15(r, c)
                                    UpNum2(c) = chart15(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart15.GetUpperBound(0) - 1
                        If chart15(r, 0) <= mys AndAlso chart15(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart15(r, c)
                                UpNum2(c) = chart15(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "D" Then
                    Dim chart16(,) As Double = {{200, -2},
                        {210, -7},
                        {220, -11},
                        {230, -14},
                        {240, -16},
                        {250, -19},
                        {260, -21},
                        {270, -23},
                        {280, -24},
                        {290, -26},
                        {300, -27},
                        {310, -28},
                        {320, -30},
                        {330, -31},
                        {340, -32},
                        {360, -34},
                        {380, -36},
                        {400, -37},
                        {420, -39},
                        {440, -40},
                        {460, -42},
                        {480, -43},
                        {500, -44},
                        {520, -45},
                        {540, -46},
                        {560, -47}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart16.GetUpperBound(0) - 1
                            If chart16(r, 0) <= 210 AndAlso chart16(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart16(r, c)
                                    UpNum2(c) = chart16(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart16.GetUpperBound(0) - 1
                            If chart16(r, 0) <= 540 AndAlso chart16(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart16(r, c)
                                    UpNum2(c) = chart16(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart16.GetUpperBound(0) - 1
                        If chart16(r, 0) <= mys AndAlso chart16(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart16(r, c)
                                UpNum2(c) = chart16(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If ComboBox1.Text = "Carbon Steel" Then
                If ComboBox157.Text = "A" Then
                    Dim chart2(,) As Double = {{30, 104},
                        {32, 97},
                        {34, 91},
                        {36, 86},
                        {38, 81},
                        {40, 78},
                        {42, 74},
                        {44, 71},
                        {46, 68},
                        {48, 66},
                        {50, 63}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 32 AndAlso chart2(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 48 AndAlso chart2(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= mys AndAlso chart2(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "B" Then
                    Dim chart3(,) As Double = {{30, 66},
                        {32, 59},
                        {34, 53},
                        {36, 48},
                        {38, 43},
                        {40, 40},
                        {42, 36},
                        {44, 33},
                        {46, 30},
                        {48, 28},
                        {50, 25}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                            If chart3(r, 0) <= 32 AndAlso chart3(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart3(r, c)
                                    UpNum2(c) = chart3(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                            If chart3(r, 0) <= 48 AndAlso chart3(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart3(r, c)
                                    UpNum2(c) = chart3(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                        If chart3(r, 0) <= mys AndAlso chart3(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart3(r, c)
                                UpNum2(c) = chart3(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "C" Then
                    Dim chart4(,) As Double = {{30, 28},
                        {32, 21},
                        {34, 15},
                        {36, 10},
                        {38, 5},
                        {40, 2},
                        {42, -2},
                        {44, -5},
                        {46, -8},
                        {48, -10},
                        {50, -13}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                            If chart4(r, 0) <= 32 AndAlso chart4(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart4(r, c)
                                    UpNum2(c) = chart4(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                            If chart4(r, 0) <= 48 AndAlso chart4(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart4(r, c)
                                    UpNum2(c) = chart4(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                        If chart4(r, 0) <= mys AndAlso chart4(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart4(r, c)
                                UpNum2(c) = chart4(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "D" Then
                    Dim chart5(,) As Double = {{30, 2},
                        {32, -5},
                        {34, -11},
                        {36, -16},
                        {38, -21},
                        {40, -24},
                        {42, -28},
                        {44, -31},
                        {46, -34},
                        {48, -36},
                        {50, -39}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                            If chart5(r, 0) <= 32 AndAlso chart5(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart5(r, c)
                                    UpNum2(c) = chart5(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                            If chart5(r, 0) <= 48 AndAlso chart5(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart5(r, c)
                                    UpNum2(c) = chart5(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                        If chart5(r, 0) <= mys AndAlso chart5(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart5(r, c)
                                UpNum2(c) = chart5(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

            ElseIf ComboBox1.Text = "Low Alloy Steel" Then
                If ComboBox157.Text = "A" Then
                    Dim chart6(,) As Double = {{30, 124},
                        {32, 115},
                        {34, 107},
                        {36, 101},
                        {38, 96},
                        {40, 92},
                        {42, 88},
                        {44, 85},
                        {46, 81},
                        {48, 79},
                        {50, 76},
                        {52, 73},
                        {54, 71},
                        {56, 69},
                        {58, 67},
                        {60, 65},
                        {62, 63},
                        {64, 62},
                        {66, 60},
                        {68, 58},
                        {70, 57},
                        {72, 56},
                        {74, 54},
                        {76, 53},
                        {78, 52},
                        {90, 51}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                            If chart6(r, 0) <= 32 AndAlso chart6(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart6(r, c)
                                    UpNum2(c) = chart6(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                            If chart6(r, 0) <= 78 AndAlso chart6(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart6(r, c)
                                    UpNum2(c) = chart6(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                        If chart6(r, 0) <= mys AndAlso chart6(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart6(r, c)
                                UpNum2(c) = chart6(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "B" Then
                    Dim chart7(,) As Double = {{30, 86},
                        {32, 77},
                        {34, 69},
                        {36, 63},
                        {38, 58},
                        {40, 54},
                        {42, 50},
                        {44, 47},
                        {46, 43},
                        {48, 41},
                        {50, 38},
                        {52, 35},
                        {54, 33},
                        {56, 31},
                        {58, 29},
                        {60, 27},
                        {62, 25},
                        {64, 24},
                        {66, 22},
                        {68, 20},
                        {70, 19},
                        {72, 18},
                        {74, 16},
                        {76, 15},
                        {78, 14},
                        {90, 13}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart7.GetUpperBound(0) - 1
                            If chart7(r, 0) <= 32 AndAlso chart7(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart7(r, c)
                                    UpNum2(c) = chart7(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart7.GetUpperBound(0) - 1
                            If chart7(r, 0) <= 78 AndAlso chart7(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart7(r, c)
                                    UpNum2(c) = chart7(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart7.GetUpperBound(0) - 1
                        If chart7(r, 0) <= mys AndAlso chart7(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart7(r, c)
                                UpNum2(c) = chart7(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "C" Then
                    Dim chart8(,) As Double = {{30, 48},
                        {32, 39},
                        {34, 31},
                        {36, 25},
                        {38, 20},
                        {40, 16},
                        {42, 12},
                        {44, 9},
                        {46, 5},
                        {48, 3},
                        {50, 0},
                        {52, -3},
                        {54, -5},
                        {56, -7},
                        {58, -9},
                        {60, -11},
                        {62, -13},
                        {64, -14},
                        {66, -16},
                        {68, -18},
                        {70, -19},
                        {72, -20},
                        {74, -22},
                        {76, -23},
                        {78, -24},
                        {90, -25}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart8.GetUpperBound(0) - 1
                            If chart8(r, 0) <= 32 AndAlso chart8(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart8(r, c)
                                    UpNum2(c) = chart8(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart8.GetUpperBound(0) - 1
                            If chart8(r, 0) <= 78 AndAlso chart8(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart8(r, c)
                                    UpNum2(c) = chart8(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart8.GetUpperBound(0) - 1
                        If chart8(r, 0) <= mys AndAlso chart8(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart8(r, c)
                                UpNum2(c) = chart8(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox157.Text = "D" Then
                    Dim chart9(,) As Double = {{30, 22},
                        {32, 13},
                        {34, 5},
                        {36, -1},
                        {38, -6},
                        {40, -10},
                        {42, -14},
                        {44, -17},
                        {46, -21},
                        {48, -23},
                        {50, -26},
                        {52, -29},
                        {54, -31},
                        {56, -33},
                        {58, -35},
                        {60, -37},
                        {62, -39},
                        {64, -40},
                        {66, -42},
                        {68, -44},
                        {70, -45},
                        {72, -46},
                        {74, -48},
                        {76, -49},
                        {78, -50},
                        {90, -51}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart9.GetUpperBound(0) - 1
                            If chart9(r, 0) <= 32 AndAlso chart9(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart9(r, c)
                                    UpNum2(c) = chart9(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart9.GetUpperBound(0) - 1
                            If chart9(r, 0) <= 78 AndAlso chart9(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart9(r, c)
                                    UpNum2(c) = chart9(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart9.GetUpperBound(0) - 1
                        If chart9(r, 0) <= mys AndAlso chart9(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart9(r, c)
                                UpNum2(c) = chart9(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If
            End If
        End If

        Dim cetkurangreftemp As Double

        cetkurangreftemp = Val(TextBox90.Text) - reftemp

        Dim basedf As Double

        If Units.ComboBox1.Text = "SI" Then
            If ComboBox158.Text = "Component Not Subject to PWHT" Then
                Dim chart111(,) As Double = {{-56, 4, 61, 579, 1436, 2336, 3160, 3883, 4509, 5000},
                    {-44, 3, 46, 474, 1239, 2080, 2873, 3581, 4203, 4746},
                    {-33, 2, 30, 350, 988, 1740, 2479, 3160, 3769, 4310},
                    {-22, 2, 16, 220, 697, 1317, 1969, 2596, 3176, 3703},
                    {-11, 1.2, 7, 109, 405, 850, 1366, 1897, 2415, 2903},
                    {-0, 0.9, 3, 39, 175, 424, 759, 1142, 1545, 1950},
                    {11, 0.1, 1.3, 10, 49, 143, 296, 500, 741, 1008},
                    {22, 0, 0.7, 2, 9, 29, 69, 133, 224, 338},
                    {33, 0, 0, 1, 2, 4, 9, 19, 36, 60},
                    {44, 0, 0, 0, 0.8, 1.1, 2, 2, 4, 6},
                    {56, 0, 0, 0, 0, 0, 0, 0.9, 1.1, 1.2}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = cetkurangreftemp
                th = Val(TextBox1.Text)

                If th < 6.4 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                basedf = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 12.7 AndAlso th <= 25.4 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 25.4
                                Lot = 12.7
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 25.4 AndAlso th <= 38.1 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 38.1
                                Lot = 25.4
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 38.1 AndAlso th <= 50.8 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 50.8
                                Lot = 38.1
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 50.8 AndAlso th <= 63.5 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 63.5
                                Lot = 50.8
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 63.5 AndAlso th <= 76.2 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 76.2
                                Lot = 63.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 76.2 AndAlso th <= 88.9 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 88.9
                                Lot = 76.2
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 88.9 AndAlso th <= 101.6 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                basedf = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -56 Then
                    If th < 6.4 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 25.4
                                    Lot = 12.7
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 38.1
                                    Lot = 25.4
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 50.8
                                    Lot = 38.1
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 63.5
                                    Lot = 50.8
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 76.2
                                    Lot = 63.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 88.9
                                    Lot = 76.2
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 56 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 25.4
                                    Lot = 12.7
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 38.1
                                    Lot = 25.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 50.8
                                    Lot = 38.1
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 63.5
                                    Lot = 50.8
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 76.2
                                    Lot = 63.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 88.9
                                    Lot = 76.2
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox158.Text = "Component Subject to PWHT" Then
                Dim chart1111(,) As Double = {{-56, 0, 1.3, 9, 46, 133, 277, 472, 704, 962},
                    {-44, 0, 1.2, 7, 34, 102, 219, 382, 582, 810},
                    {-33, 0, 1.1, 5, 22, 68, 153, 277, 436, 623},
                    {-22, 0, 0.9, 3, 12, 38, 90, 171, 281, 416},
                    {-11, 0, 0.4, 2, 5, 17, 41, 83, 144, 224},
                    {-0, 0, 0, 1.1, 2, 6, 14, 29, 53, 88},
                    {11, 0, 0, 0.6, 1.2, 2, 4, 7, 13, 23},
                    {22, 0, 0, 0, 0.5, 1.1, 1.3, 2, 3, 4},
                    {33, 0, 0, 0, 0, 0, 0.5, 0.9, 1.1, 1.3},
                    {44, 0, 0, 0, 0, 0, 0, 0, 0, 0.2},
                    {56, 0, 0, 0, 0, 0, 0, 0, 0, 0}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = cetkurangreftemp
                th = Val(TextBox1.Text)

                If th < 6.4 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                basedf = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 12.7 AndAlso th <= 25.4 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 25.4
                                Lot = 12.7
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 25.4 AndAlso th <= 38.1 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 38.1
                                Lot = 25.4
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 38.1 AndAlso th <= 50.8 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 50.8
                                Lot = 38.1
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 50.8 AndAlso th <= 63.5 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 63.5
                                Lot = 50.8
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 63.5 AndAlso th <= 76.2 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 76.2
                                Lot = 63.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 76.2 AndAlso th <= 88.9 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 88.9
                                Lot = 76.2
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 88.9 AndAlso th <= 101.6 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                basedf = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -56 Then
                    If th < 6.4 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 25.4
                                    Lot = 12.7
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 38.1
                                    Lot = 25.4
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 50.8
                                    Lot = 38.1
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 63.5
                                    Lot = 50.8
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 76.2
                                    Lot = 63.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 88.9
                                    Lot = 76.2
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 56 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 25.4
                                    Lot = 12.7
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 38.1
                                    Lot = 25.4
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 50.8
                                    Lot = 38.1
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 63.5
                                    Lot = 50.8
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 76.2
                                    Lot = 63.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 88.9
                                    Lot = 76.2
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then

            If ComboBox158.Text = "Component Subject to PWHT" Then
                Dim chart11(,) As Double = {{-100, 0, 1.3, 9, 46, 133, 277, 472, 704, 962},
                    {-80, 0, 1.2, 7, 34, 102, 219, 382, 582, 810},
                    {-60, 0, 1.1, 5, 22, 68, 153, 277, 436, 623},
                    {-40, 0, 0.9, 3, 12, 38, 90, 171, 281, 416},
                    {-20, 0, 0.4, 2, 5, 17, 41, 83, 144, 224},
                    {0, 0, 0, 1.1, 2, 6, 14, 29, 53, 88},
                    {20, 0, 0, 0.6, 1.2, 2, 4, 7, 13, 23},
                    {40, 0, 0, 0, 0.5, 1.1, 1.3, 2, 3, 4},
                    {60, 0, 0, 0, 0, 0, 0.5, 0.9, 1.1, 1.3},
                    {80, 0, 0, 0, 0, 0, 0, 0, 0, 2},
                    {100, 0, 0, 0, 0, 0, 0, 0, 0, 0}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = cetkurangreftemp
                th = Val(TextBox1.Text)

                If th < 0.25 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                basedf = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 0.5 AndAlso th <= 1.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.0
                                Lot = 0.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.0 AndAlso th <= 1.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.5
                                Lot = 1.0
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.5 AndAlso th <= 2.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.0
                                Lot = 1.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.0 AndAlso th <= 2.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.5
                                Lot = 2.0
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.5 AndAlso th <= 3.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.0
                                Lot = 2.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.0 AndAlso th <= 3.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.5
                                Lot = 3.0
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.5 AndAlso th <= 4.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                basedf = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.0
                                    Lot = 0.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.5
                                    Lot = 1.0
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.0
                                    Lot = 1.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.5
                                    Lot = 2.0
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.0
                                    Lot = 2.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.5
                                    Lot = 3.0
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.0
                                    Lot = 0.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.5
                                    Lot = 1.0
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.0
                                    Lot = 1.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.5
                                    Lot = 2.0
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.0
                                    Lot = 2.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.5
                                    Lot = 3.0
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox158.Text = "Component Subject to PWHT" Then
                Dim chart11(,) As Double = {{-100, 0, 1.3, 9, 46, 133, 277, 472, 704, 962},
                    {-80, 0, 1.2, 7, 34, 102, 219, 382, 582, 810},
                    {-60, 0, 1.1, 5, 22, 68, 153, 277, 436, 623},
                    {-40, 0, 0.9, 3, 12, 38, 90, 171, 281, 416},
                    {-20, 0, 0.4, 2, 5, 17, 41, 83, 144, 224},
                    {0, 0, 0, 1.1, 2, 6, 14, 29, 53, 88},
                    {20, 0, 0, 0.6, 1.2, 2, 4, 7, 13, 23},
                    {40, 0, 0, 0, 0.5, 1.1, 1.3, 2, 3, 4},
                    {60, 0, 0, 0, 0, 0, 0.5, 0.9, 1.1, 1.3},
                    {80, 0, 0, 0, 0, 0, 0, 0, 0, 2},
                    {100, 0, 0, 0, 0, 0, 0, 0, 0, 0}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = cetkurangreftemp
                th = Val(TextBox1.Text)

                If th < 0.25 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                basedf = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 0.5 AndAlso th <= 1.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.0
                                Lot = 0.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.0 AndAlso th <= 1.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.5
                                Lot = 1.0
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.5 AndAlso th <= 2.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.0
                                Lot = 1.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.0 AndAlso th <= 2.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.5
                                Lot = 2.0
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.5 AndAlso th <= 3.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.0
                                Lot = 2.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.0 AndAlso th <= 3.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.5
                                Lot = 3.0
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.5 AndAlso th <= 4.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                basedf = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                basedf = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.0
                                    Lot = 0.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.5
                                    Lot = 1.0
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.0
                                    Lot = 1.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.5
                                    Lot = 2.0
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.0
                                    Lot = 2.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.5
                                    Lot = 3.0
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.0
                                    Lot = 0.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.5
                                    Lot = 1.0
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.0
                                    Lot = 1.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.5
                                    Lot = 2.0
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.0
                                    Lot = 2.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.5
                                    Lot = 3.0
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    basedf = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End If

        If ComboBox158.Text = "Component Subject to PWHT" And Units.ComboBox1.Text = "SI" And TextBox1.Text = 101.6 And cetkurangreftemp = 56 Then
            basedf = "0"
        End If

        If basedf < 0 Then
            basedf = 0
        End If

        Dim fse As Double

        fse = Val(ComboBox159.Text)

        Dim df As Double
        df = basedf * fse

        Label364.Text = df
    End Sub

    'Low Alloy Steel Embrittlement Coding perhitungan ---------------------------------------------------------------------
    Private Sub fatt()

        Dim fatt As Double

        If ComboBox171.Text = "Engineering analysis or Impact Test" Then
            fatt = Val(TextBox92.Text)
        ElseIf ComboBox171.Text = "Step Cooling Embrittlement (SCE) test" Then
            fatt = 0.67 * (Math.Log10(Val(TextBox94.Text) - 0.91) * Val(TextBox93.Text))
        ElseIf ComboBox171.Text = "J-factor correlation, X-bar factor" Then
            If Label402.Enabled = False Then
                'J-factor
                fatt = (Val(TextBox95.Text) + Val(TextBox96.Text)) * (Val(TextBox97.Text) + Val(TextBox98.Text)) * 10000

            ElseIf Label400.Enabled = False Then
                'X-bar
                fatt = (Val(TextBox99.Text) + Val(TextBox100.Text) + Val(TextBox101.Text) + Val(TextBox102.Text)) * 100
            End If
        ElseIf ComboBox171.Text = "Conservative Assumption based on year of fabrication" Then
            If Units.ComboBox1.Text = "SI" Then
                If ComboBox172.Text = "1 generation equipment (1965 to 1972)" Then
                    fatt = 177
                ElseIf ComboBox172.Text = "2 generation equipment (1973 to 1980)" Then
                    fatt = 149
                ElseIf ComboBox172.Text = "3 generation equipment (1981 to 1987)" Then
                    fatt = 121
                ElseIf ComboBox172.Text = "4 generation equipment (after to 1998)" Then
                    fatt = 66
                End If
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                If ComboBox172.Text = "1 generation equipment (1965 to 1972)" Then
                    fatt = 350
                ElseIf ComboBox172.Text = "2 generation equipment (1973 to 1980)" Then
                    fatt = 300
                ElseIf ComboBox172.Text = "3 generation equipment (1981 to 1987)" Then
                    fatt = 250
                ElseIf ComboBox172.Text = "4 generation equipment (after to 1998)" Then
                    fatt = 150
                End If
            End If
        End If

        TextBox103.Text = fatt
    End Sub

    Public Sub lowalloydf()

        Dim reftemp As Double

        If Units.ComboBox1.Text = "SI" Then
            If ComboBox1.Text = "Carbon Steel" Then
                If ComboBox170.Text = "A" Then
                    Dim chart(,) As Double = {{200, 42},
                        {210, 38},
                        {220, 36},
                        {230, 33},
                        {240, 31},
                        {260, 27},
                        {280, 24},
                        {300, 22},
                        {320, 19},
                        {340, 17},
                        {360, 15}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart.GetUpperBound(0) - 1
                            If chart(r, 0) <= 210 AndAlso chart(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart(r, c)
                                    UpNum2(c) = chart(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart.GetUpperBound(0) - 1
                            If chart(r, 0) <= 340 AndAlso chart(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart(r, c)
                                    UpNum2(c) = chart(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart.GetUpperBound(0) - 1
                        If chart(r, 0) <= mys AndAlso chart(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart(r, c)
                                UpNum2(c) = chart(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "B" Then
                    Dim chart10(,) As Double = {{200, 21},
                        {210, 17},
                        {220, 15},
                        {230, 1.25},
                        {240, 10},
                        {260, 6},
                        {280, 3},
                        {300, 1},
                        {320, -2},
                        {340, -4},
                        {360, -6}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart10.GetUpperBound(0) - 1
                            If chart10(r, 0) <= 210 AndAlso chart10(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart10(r, c)
                                    UpNum2(c) = chart10(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart10.GetUpperBound(0) - 1
                            If chart10(r, 0) <= 340 AndAlso chart10(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart10(r, c)
                                    UpNum2(c) = chart10(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart10.GetUpperBound(0) - 1
                        If chart10(r, 0) <= mys AndAlso chart10(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart10(r, c)
                                UpNum2(c) = chart10(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "C" Then
                    Dim chart11(,) As Double = {{200, 0},
                        {210, -4},
                        {220, -7},
                        {230, -9},
                        {240, -11},
                        {260, -15},
                        {280, -18},
                        {300, -21},
                        {320, -23},
                        {340, -25},
                        {360, -27}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 210 AndAlso chart11(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 340 AndAlso chart11(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= mys AndAlso chart11(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "D" Then
                    Dim chart12(,) As Double = {{200, -15},
                        {210, -18},
                        {220, -21},
                        {230, -23},
                        {240, -26},
                        {260, -29},
                        {280, -32},
                        {300, -35},
                        {320, -37},
                        {340, -39},
                        {360, -41}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart12.GetUpperBound(0) - 1
                            If chart12(r, 0) <= 210 AndAlso chart12(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart12(r, c)
                                    UpNum2(c) = chart12(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 360 Then

                        For r As Integer = 0 To chart12.GetUpperBound(0) - 1
                            If chart12(r, 0) <= 340 AndAlso chart12(r + 1, 0) >= 360 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart12(r, c)
                                    UpNum2(c) = chart12(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 360
                                    lot = 340

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart12.GetUpperBound(0) - 1
                        If chart12(r, 0) <= mys AndAlso chart12(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart12(r, c)
                                UpNum2(c) = chart12(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

            ElseIf ComboBox1.Text = "Low Alloy Steel" Then
                If ComboBox170.Text = "A" Then
                    Dim chart13(,) As Double = {{200, 55},
                        {210, 50},
                        {220, 46},
                        {230, 43},
                        {240, 40},
                        {250, 38},
                        {260, 36},
                        {270, 34},
                        {280, 32},
                        {290, 31},
                        {300, 30},
                        {310, 28},
                        {320, 27},
                        {330, 26},
                        {340, 25},
                        {360, 23},
                        {380, 21},
                        {400, 19},
                        {420, 18},
                        {440, 16},
                        {460, 15},
                        {480, 14},
                        {500, 13},
                        {520, 12},
                        {540, 11},
                        {560, 10}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart13.GetUpperBound(0) - 1
                            If chart13(r, 0) <= 210 AndAlso chart13(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart13(r, c)
                                    UpNum2(c) = chart13(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart13.GetUpperBound(0) - 1
                            If chart13(r, 0) <= 540 AndAlso chart13(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart13(r, c)
                                    UpNum2(c) = chart13(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart13.GetUpperBound(0) - 1
                        If chart13(r, 0) <= mys AndAlso chart13(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart13(r, c)
                                UpNum2(c) = chart13(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "B" Then
                    Dim chart14(,) As Double = {{200, 33},
                        {210, 29},
                        {220, 25},
                        {230, 22},
                        {240, 19},
                        {250, 17},
                        {260, 15},
                        {270, 13},
                        {280, 11},
                        {290, 10},
                        {300, 8},
                        {310, 7},
                        {320, 6},
                        {330, 5},
                        {340, 4},
                        {360, 2},
                        {380, 0},
                        {400, -2},
                        {420, -3},
                        {440, -5},
                        {460, -6},
                        {480, -7},
                        {500, -8},
                        {520, -9},
                        {540, -10},
                        {560, -11}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart14.GetUpperBound(0) - 1
                            If chart14(r, 0) <= 210 AndAlso chart14(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart14(r, c)
                                    UpNum2(c) = chart14(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart14.GetUpperBound(0) - 1
                            If chart14(r, 0) <= 540 AndAlso chart14(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart14(r, c)
                                    UpNum2(c) = chart14(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart14.GetUpperBound(0) - 1
                        If chart14(r, 0) <= mys AndAlso chart14(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart14(r, c)
                                UpNum2(c) = chart14(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "C" Then
                    Dim chart15(,) As Double = {{200, 12},
                        {210, 8},
                        {220, 4},
                        {230, 1},
                        {240, -2},
                        {250, -4},
                        {260, -6},
                        {270, -8},
                        {280, -10},
                        {290, -11},
                        {300, -13},
                        {310, -14},
                        {320, -15},
                        {330, -16},
                        {340, -17},
                        {360, -19},
                        {380, -21},
                        {400, -23},
                        {420, -24},
                        {440, -26},
                        {460, -27},
                        {480, -28},
                        {500, -29},
                        {520, -30},
                        {540, -31},
                        {560, -32}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart15.GetUpperBound(0) - 1
                            If chart15(r, 0) <= 210 AndAlso chart15(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart15(r, c)
                                    UpNum2(c) = chart15(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart15.GetUpperBound(0) - 1
                            If chart15(r, 0) <= 540 AndAlso chart15(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart15(r, c)
                                    UpNum2(c) = chart15(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart15.GetUpperBound(0) - 1
                        If chart15(r, 0) <= mys AndAlso chart15(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart15(r, c)
                                UpNum2(c) = chart15(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "D" Then
                    Dim chart16(,) As Double = {{200, -2},
                        {210, -7},
                        {220, -11},
                        {230, -14},
                        {240, -16},
                        {250, -19},
                        {260, -21},
                        {270, -23},
                        {280, -24},
                        {290, -26},
                        {300, -27},
                        {310, -28},
                        {320, -30},
                        {330, -31},
                        {340, -32},
                        {360, -34},
                        {380, -36},
                        {400, -37},
                        {420, -39},
                        {440, -40},
                        {460, -42},
                        {480, -43},
                        {500, -44},
                        {520, -45},
                        {540, -46},
                        {560, -47}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 200 Then
                        For r As Integer = 0 To chart16.GetUpperBound(0) - 1
                            If chart16(r, 0) <= 210 AndAlso chart16(r + 1, 0) >= 200 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart16(r, c)
                                    UpNum2(c) = chart16(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 210
                                    lot = 200

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 560 Then

                        For r As Integer = 0 To chart16.GetUpperBound(0) - 1
                            If chart16(r, 0) <= 540 AndAlso chart16(r + 1, 0) >= 560 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart16(r, c)
                                    UpNum2(c) = chart16(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 560
                                    lot = 540

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart16.GetUpperBound(0) - 1
                        If chart16(r, 0) <= mys AndAlso chart16(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart16(r, c)
                                UpNum2(c) = chart16(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If ComboBox1.Text = "Carbon Steel" Then
                If ComboBox170.Text = "A" Then
                    Dim chart2(,) As Double = {{30, 104},
                        {32, 97},
                        {34, 91},
                        {36, 86},
                        {38, 81},
                        {40, 78},
                        {42, 74},
                        {44, 71},
                        {46, 68},
                        {48, 66},
                        {50, 63}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 32 AndAlso chart2(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 48 AndAlso chart2(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= mys AndAlso chart2(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "B" Then
                    Dim chart3(,) As Double = {{30, 66},
                        {32, 59},
                        {34, 53},
                        {36, 48},
                        {38, 43},
                        {40, 40},
                        {42, 36},
                        {44, 33},
                        {46, 30},
                        {48, 28},
                        {50, 25}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                            If chart3(r, 0) <= 32 AndAlso chart3(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart3(r, c)
                                    UpNum2(c) = chart3(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                            If chart3(r, 0) <= 48 AndAlso chart3(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart3(r, c)
                                    UpNum2(c) = chart3(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                        If chart3(r, 0) <= mys AndAlso chart3(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart3(r, c)
                                UpNum2(c) = chart3(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "C" Then
                    Dim chart4(,) As Double = {{30, 28},
                        {32, 21},
                        {34, 15},
                        {36, 10},
                        {38, 5},
                        {40, 2},
                        {42, -2},
                        {44, -5},
                        {46, -8},
                        {48, -10},
                        {50, -13}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                            If chart4(r, 0) <= 32 AndAlso chart4(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart4(r, c)
                                    UpNum2(c) = chart4(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                            If chart4(r, 0) <= 48 AndAlso chart4(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart4(r, c)
                                    UpNum2(c) = chart4(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                        If chart4(r, 0) <= mys AndAlso chart4(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart4(r, c)
                                UpNum2(c) = chart4(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "D" Then
                    Dim chart5(,) As Double = {{30, 2},
                        {32, -5},
                        {34, -11},
                        {36, -16},
                        {38, -21},
                        {40, -24},
                        {42, -28},
                        {44, -31},
                        {46, -34},
                        {48, -36},
                        {50, -39}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                            If chart5(r, 0) <= 32 AndAlso chart5(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart5(r, c)
                                    UpNum2(c) = chart5(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 50 Then

                        For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                            If chart5(r, 0) <= 48 AndAlso chart5(r + 1, 0) >= 50 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart5(r, c)
                                    UpNum2(c) = chart5(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 50
                                    lot = 48

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                        If chart5(r, 0) <= mys AndAlso chart5(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart5(r, c)
                                UpNum2(c) = chart5(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

            ElseIf ComboBox1.Text = "Low Alloy Steel" Then
                If ComboBox170.Text = "A" Then
                    Dim chart6(,) As Double = {{30, 124},
                        {32, 115},
                        {34, 107},
                        {36, 101},
                        {38, 96},
                        {40, 92},
                        {42, 88},
                        {44, 85},
                        {46, 81},
                        {48, 79},
                        {50, 76},
                        {52, 73},
                        {54, 71},
                        {56, 69},
                        {58, 67},
                        {60, 65},
                        {62, 63},
                        {64, 62},
                        {66, 60},
                        {68, 58},
                        {70, 57},
                        {72, 56},
                        {74, 54},
                        {76, 53},
                        {78, 52},
                        {90, 51}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                            If chart6(r, 0) <= 32 AndAlso chart6(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart6(r, c)
                                    UpNum2(c) = chart6(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                            If chart6(r, 0) <= 78 AndAlso chart6(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart6(r, c)
                                    UpNum2(c) = chart6(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                        If chart6(r, 0) <= mys AndAlso chart6(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart6(r, c)
                                UpNum2(c) = chart6(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "B" Then
                    Dim chart7(,) As Double = {{30, 86},
                        {32, 77},
                        {34, 69},
                        {36, 63},
                        {38, 58},
                        {40, 54},
                        {42, 50},
                        {44, 47},
                        {46, 43},
                        {48, 41},
                        {50, 38},
                        {52, 35},
                        {54, 33},
                        {56, 31},
                        {58, 29},
                        {60, 27},
                        {62, 25},
                        {64, 24},
                        {66, 22},
                        {68, 20},
                        {70, 19},
                        {72, 18},
                        {74, 16},
                        {76, 15},
                        {78, 14},
                        {90, 13}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart7.GetUpperBound(0) - 1
                            If chart7(r, 0) <= 32 AndAlso chart7(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart7(r, c)
                                    UpNum2(c) = chart7(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart7.GetUpperBound(0) - 1
                            If chart7(r, 0) <= 78 AndAlso chart7(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart7(r, c)
                                    UpNum2(c) = chart7(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart7.GetUpperBound(0) - 1
                        If chart7(r, 0) <= mys AndAlso chart7(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart7(r, c)
                                UpNum2(c) = chart7(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "C" Then
                    Dim chart8(,) As Double = {{30, 48},
                        {32, 39},
                        {34, 31},
                        {36, 25},
                        {38, 20},
                        {40, 16},
                        {42, 12},
                        {44, 9},
                        {46, 5},
                        {48, 3},
                        {50, 0},
                        {52, -3},
                        {54, -5},
                        {56, -7},
                        {58, -9},
                        {60, -11},
                        {62, -13},
                        {64, -14},
                        {66, -16},
                        {68, -18},
                        {70, -19},
                        {72, -20},
                        {74, -22},
                        {76, -23},
                        {78, -24},
                        {90, -25}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart8.GetUpperBound(0) - 1
                            If chart8(r, 0) <= 32 AndAlso chart8(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart8(r, c)
                                    UpNum2(c) = chart8(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart8.GetUpperBound(0) - 1
                            If chart8(r, 0) <= 78 AndAlso chart8(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart8(r, c)
                                    UpNum2(c) = chart8(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart8.GetUpperBound(0) - 1
                        If chart8(r, 0) <= mys AndAlso chart8(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart8(r, c)
                                UpNum2(c) = chart8(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If

                If ComboBox170.Text = "D" Then
                    Dim chart9(,) As Double = {{30, 22},
                        {32, 13},
                        {34, 5},
                        {36, -1},
                        {38, -6},
                        {40, -10},
                        {42, -14},
                        {44, -17},
                        {46, -21},
                        {48, -23},
                        {50, -26},
                        {52, -29},
                        {54, -31},
                        {56, -33},
                        {58, -35},
                        {60, -37},
                        {62, -39},
                        {64, -40},
                        {66, -42},
                        {68, -44},
                        {70, -45},
                        {72, -46},
                        {74, -48},
                        {76, -49},
                        {78, -50},
                        {90, -51}}

                    Dim UpNum2(1) As Double
                    Dim LowNum2(1) As Double

                    Dim Lo1 As Double
                    Dim Lo2 As Double
                    Dim up1 As Double
                    Dim up2 As Double

                    Dim mys As Double
                    Dim upt As Double
                    Dim lot As Double
                    Dim x As Double
                    Dim z As Double


                    mys = Val(TextBox13.Text)

                    If mys < 30 Then
                        For r As Integer = 0 To chart9.GetUpperBound(0) - 1
                            If chart9(r, 0) <= 32 AndAlso chart9(r + 1, 0) >= 30 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart9(r, c)
                                    UpNum2(c) = chart9(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 32
                                    lot = 30

                                    reftemp = Inter3b(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    ElseIf mys > 80 Then

                        For r As Integer = 0 To chart9.GetUpperBound(0) - 1
                            If chart9(r, 0) <= 78 AndAlso chart9(r + 1, 0) >= 80 Then
                                For c As Integer = 0 To 1
                                    LowNum2(c) = chart9(r, c)
                                    UpNum2(c) = chart9(r + 1, c)
                                    x = LowNum2(1)
                                    z = UpNum2(1)
                                    upt = 80
                                    lot = 78

                                    reftemp = Inter3c(upt, lot, mys, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If

                    For r As Integer = 0 To chart9.GetUpperBound(0) - 1
                        If chart9(r, 0) <= mys AndAlso chart9(r + 1, 0) >= mys Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart9(r, c)
                                UpNum2(c) = chart9(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                reftemp = Inter(mys, up1, up2, Lo1, Lo2)


                            Next
                            Exit For
                        End If
                    Next
                End If
            End If
        End If

        Dim reftemptambahfatt As Single

        reftemptambahfatt = reftemp + Val(TextBox103.Text)

        Dim tmptkurangreftemptambahfatt As Single

        tmptkurangreftemptambahfatt = Val(TextBox91.Text) - reftemptambahfatt

        Dim df As Double

        If Units.ComboBox1.Text = "SI" Then
            If ComboBox173.Text = "Component Not Subject to PWHT" Then
                Dim chart111(,) As Double = {{-56, 4, 61, 579, 1436, 2336, 3160, 3883, 4509, 5000},
                    {-44, 3, 46, 474, 1239, 2080, 2873, 3581, 4203, 4746},
                    {-33, 2, 30, 350, 988, 1740, 2479, 3160, 3769, 4310},
                    {-22, 2, 16, 220, 697, 1317, 1969, 2596, 3176, 3703},
                    {-11, 1.2, 7, 109, 405, 850, 1366, 1897, 2415, 2903},
                    {-0, 0.9, 3, 39, 175, 424, 759, 1142, 1545, 1950},
                    {11, 0.1, 1.3, 10, 49, 143, 296, 500, 741, 1008},
                    {22, 0, 0.7, 2, 9, 29, 69, 133, 224, 338},
                    {33, 0, 0, 1, 2, 4, 9, 19, 36, 60},
                    {44, 0, 0, 0, 0.8, 1.1, 2, 2, 4, 6},
                    {56, 0, 0, 0, 0, 0, 0, 0.9, 1.1, 1.2}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = tmptkurangreftemptambahfatt
                th = Val(TextBox1.Text)

                If th < 6.4 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                df = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 12.7 AndAlso th <= 25.4 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 25.4
                                Lot = 12.7
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 25.4 AndAlso th <= 38.1 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 38.1
                                Lot = 25.4
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 38.1 AndAlso th <= 50.8 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 50.8
                                Lot = 38.1
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 50.8 AndAlso th <= 63.5 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 63.5
                                Lot = 50.8
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 63.5 AndAlso th <= 76.2 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 76.2
                                Lot = 63.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 76.2 AndAlso th <= 88.9 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 88.9
                                Lot = 76.2
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 88.9 AndAlso th <= 101.6 Then
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                        If chart111(r, 0) <= temp AndAlso chart111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart111(r, c)
                                UpNum2(c) = chart111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                df = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -56 Then
                    If th < 6.4 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 25.4
                                    Lot = 12.7
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 38.1
                                    Lot = 25.4
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 50.8
                                    Lot = 38.1
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 63.5
                                    Lot = 50.8
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 76.2
                                    Lot = 63.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 88.9
                                    Lot = 76.2
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= -44 AndAlso chart111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 56 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 25.4
                                    Lot = 12.7
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 38.1
                                    Lot = 25.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 50.8
                                    Lot = 38.1
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 63.5
                                    Lot = 50.8
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 76.2
                                    Lot = 63.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 88.9
                                    Lot = 76.2
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart111.GetUpperBound(0) - 1
                            If chart111(r, 0) <= 44 AndAlso chart111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart111(r, c)
                                    UpNum2(c) = chart111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox173.Text = "Component Subject to PWHT" Then
                Dim chart1111(,) As Double = {{-56, 0, 1.3, 9, 46, 133, 277, 472, 704, 962},
                    {-44, 0, 1.2, 7, 34, 102, 219, 382, 582, 810},
                    {-33, 0, 1.1, 5, 22, 68, 153, 277, 436, 623},
                    {-22, 0, 0.9, 3, 12, 38, 90, 171, 281, 416},
                    {-11, 0, 0.4, 2, 5, 17, 41, 83, 144, 224},
                    {-0, 0, 0, 1.1, 2, 6, 14, 29, 53, 88},
                    {11, 0, 0, 0.6, 1.2, 2, 4, 7, 13, 23},
                    {22, 0, 0, 0, 0.5, 1.1, 1.3, 2, 3, 4},
                    {33, 0, 0, 0, 0, 0, 0.5, 0.9, 1.1, 1.3},
                    {44, 0, 0, 0, 0, 0, 0, 0, 0, 0.2},
                    {56, 0, 0, 0, 0, 0, 0, 0, 0, 0}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = tmptkurangreftemptambahfatt
                th = Val(TextBox1.Text)

                If th < 6.4 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                df = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 12.7
                                Lot = 6.4
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 12.7 AndAlso th <= 25.4 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 25.4
                                Lot = 12.7
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 25.4 AndAlso th <= 38.1 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 38.1
                                Lot = 25.4
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 38.1 AndAlso th <= 50.8 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 50.8
                                Lot = 38.1
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 50.8 AndAlso th <= 63.5 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 63.5
                                Lot = 50.8
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 63.5 AndAlso th <= 76.2 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 76.2
                                Lot = 63.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 76.2 AndAlso th <= 88.9 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 88.9
                                Lot = 76.2
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 88.9 AndAlso th <= 101.6 Then
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                        If chart1111(r, 0) <= temp AndAlso chart1111(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart1111(r, c)
                                UpNum2(c) = chart1111(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 101.6
                                Lot = 88.9
                                df = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -56 Then
                    If th < 6.4 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 25.4
                                    Lot = 12.7
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 38.1
                                    Lot = 25.4
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 50.8
                                    Lot = 38.1
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 63.5
                                    Lot = 50.8
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 76.2
                                    Lot = 63.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 88.9
                                    Lot = 76.2
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= -44 AndAlso chart1111(r + 1, 0) >= -56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -44
                                    lota = -56
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -44
                                    lotb = -56
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 56 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 6.4 AndAlso th <= 12.7 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 12.7
                                    Lot = 6.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 12.7 AndAlso th <= 25.4 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 25.4
                                    Lot = 12.7
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 25.4 AndAlso th <= 38.1 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 38.1
                                    Lot = 25.4
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 38.1 AndAlso th <= 50.8 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 50.8
                                    Lot = 38.1
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 50.8 AndAlso th <= 63.5 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 63.5
                                    Lot = 50.8
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 63.5 AndAlso th <= 76.2 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 76.2
                                    Lot = 63.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 76.2 AndAlso th <= 88.9 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 88.9
                                    Lot = 76.2
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 88.9 AndAlso th <= 101.6 Then
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1111.GetUpperBound(0) - 1
                            If chart1111(r, 0) <= 44 AndAlso chart1111(r + 1, 0) >= 56 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1111(r, c)
                                    UpNum2(c) = chart1111(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 56
                                    lotx = 44
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 56
                                    lotz = 44
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 101.6
                                    Lot = 88.9
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If ComboBox173.Text = "Component Not Subject to PWHT" Then
                Dim chart1(,) As Double = {{-100, 4, 61, 579, 1436, 2336, 3160, 3883, 4509, 5000},
                    {-80, 3, 46, 474, 1239, 2080, 2873, 3581, 4203, 4746},
                    {-60, 2, 30, 350, 988, 1740, 2479, 3160, 3769, 4310},
                    {-40, 2, 16, 220, 697, 1317, 1969, 2596, 3176, 3703},
                    {-20, 1.2, 7, 109, 405, 850, 1366, 1897, 2415, 2903},
                    {0, 0.9, 3, 39, 175, 424, 759, 1142, 1545, 1950},
                    {20, 0.1, 1.3, 10, 49, 143, 296, 500, 741, 1008},
                    {40, 0, 0.7, 2, 9, 29, 69, 133, 224, 338},
                    {60, 0, 0, 1, 2, 4, 9, 19, 36, 60},
                    {80, 0, 0, 0, 0.8, 1.1, 2, 2, 4, 6},
                    {100, 0, 0, 0, 0, 0, 0, 0.9, 1.1, 1.2}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = tmptkurangreftemptambahfatt
                th = Val(TextBox1.Text)

                If th < 0.25 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                df = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 0.5 AndAlso th <= 1.0 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.0
                                Lot = 0.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.0 AndAlso th <= 1.5 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.5
                                Lot = 1.0
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.5 AndAlso th <= 2.0 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.0
                                Lot = 1.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.0 AndAlso th <= 2.5 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.5
                                Lot = 2.0
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.5 AndAlso th <= 3.0 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.0
                                Lot = 2.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.0 AndAlso th <= 3.5 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.5
                                Lot = 3.0
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.5 AndAlso th <= 4.0 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= temp AndAlso chart1(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                df = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.0
                                    Lot = 0.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.5
                                    Lot = 1.0
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.0
                                    Lot = 1.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.5
                                    Lot = 2.0
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.0
                                    Lot = 2.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.5
                                    Lot = 3.0
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= -80 AndAlso chart1(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.0
                                    Lot = 0.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.5
                                    Lot = 1.0
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.0
                                    Lot = 1.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.5
                                    Lot = 2.0
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.0
                                    Lot = 2.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.5
                                    Lot = 3.0
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 80 AndAlso chart1(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox173.Text = "Component Subject to PWHT" Then
                Dim chart11(,) As Double = {{-100, 0, 1.3, 9, 46, 133, 277, 472, 704, 962},
                    {-80, 0, 1.2, 7, 34, 102, 219, 382, 582, 810},
                    {-60, 0, 1.1, 5, 22, 68, 153, 277, 436, 623},
                    {-40, 0, 0.9, 3, 12, 38, 90, 171, 281, 416},
                    {-20, 0, 0.4, 2, 5, 17, 41, 83, 144, 224},
                    {0, 0, 0, 1.1, 2, 6, 14, 29, 53, 88},
                    {20, 0, 0, 0.6, 1.2, 2, 4, 7, 13, 23},
                    {40, 0, 0, 0, 0.5, 1.1, 1.3, 2, 3, 4},
                    {60, 0, 0, 0, 0, 0, 0.5, 0.9, 1.1, 1.3},
                    {80, 0, 0, 0, 0, 0, 0, 0, 0, 2},
                    {100, 0, 0, 0, 0, 0, 0, 0, 0, 0}}

                Dim UpNum2(9) As Double
                Dim LowNum2(9) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double
                Dim Lo3 As Double
                Dim Lo4 As Double
                Dim up3 As Double
                Dim up4 As Double
                Dim x As Double
                Dim z As Double
                Dim temp As Double
                Dim th As Double
                Dim upt As Double
                Dim Lot As Double
                Dim uptx As Double
                Dim uptz As Double
                Dim lotz As Double
                Dim lotx As Double
                Dim xx As Double
                Dim xz As Double
                Dim zx As Double
                Dim zz As Double
                Dim upta As Double
                Dim uptb As Double
                Dim lota As Double
                Dim lotb As Double
                Dim xa As Double
                Dim xb As Double
                Dim za As Double
                Dim zb As Double


                temp = tmptkurangreftemptambahfatt
                th = Val(TextBox1.Text)

                If th < 0.25 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                df = Inter3b1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next

                ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)

                                z = Interx1(temp, up3, up4, Lo3, Lo4)

                                upt = 0.5
                                Lot = 0.25
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf th > 0.5 AndAlso th <= 1.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.0
                                Lot = 0.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.0 AndAlso th <= 1.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 1.5
                                Lot = 1.0
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 1.5 AndAlso th <= 2.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(4)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(4)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(5)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(5)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.0
                                Lot = 1.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.0 AndAlso th <= 2.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 6
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(5)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(5)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(6)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(6)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 2.5
                                Lot = 2.0
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 2.5 AndAlso th <= 3.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 7
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(6)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(6)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(7)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(7)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.0
                                Lot = 2.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.0 AndAlso th <= 3.5 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 8
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(7)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(7)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(8)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(8)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 3.5
                                Lot = 3.0
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                ElseIf th > 3.5 AndAlso th <= 4.0 Then
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                df = Inter3a1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                        If chart11(r, 0) <= temp AndAlso chart11(r + 1, 0) >= temp Then
                            For c As Integer = 0 To 9
                                LowNum2(c) = chart11(r, c)
                                UpNum2(c) = chart11(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(8)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(8)

                                x = Interx(temp, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(9)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(9)
                                z = Interx1(temp, up3, up4, Lo3, Lo4)
                                upt = 4.0
                                Lot = 3.5
                                df = Inter3c1(upt, Lot, th, x, z)
                            Next
                        End If
                    Next
                End If

                If temp < -100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.0
                                    Lot = 0.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 1.5
                                    Lot = 1.0
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(4)
                                    xa = LowNum2(4)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(5)
                                    xb = LowNum2(5)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.0
                                    Lot = 1.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(5)
                                    xa = LowNum2(5)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(6)
                                    xb = LowNum2(6)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 2.5
                                    Lot = 2.0
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(6)
                                    xa = LowNum2(6)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(7)
                                    xb = LowNum2(7)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.0
                                    Lot = 2.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(7)
                                    xa = LowNum2(7)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(8)
                                    xb = LowNum2(8)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 3.5
                                    Lot = 3.0
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th > 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3a1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= -80 AndAlso chart11(r + 1, 0) >= -100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    za = UpNum2(8)
                                    xa = LowNum2(8)
                                    upta = -80
                                    lota = -100
                                    x = Inter3da(upta, lota, temp, xa, za)

                                    zb = UpNum2(9)
                                    xb = LowNum2(9)
                                    uptb = -80
                                    lotb = -100
                                    z = Inter3db(uptb, lotb, temp, xb, zb)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If temp > 100 Then
                    If th < 0.25 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.25 AndAlso th <= 0.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 0.5
                                    Lot = 0.25
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 0.5 AndAlso th <= 1.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.0
                                    Lot = 0.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.0 AndAlso th <= 1.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 1.5
                                    Lot = 1.0
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 1.5 AndAlso th <= 2.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(4)
                                    xx = LowNum2(4)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(5)
                                    xz = LowNum2(5)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.0
                                    Lot = 1.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.0 AndAlso th <= 2.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(5)
                                    xx = LowNum2(5)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(6)
                                    xz = LowNum2(6)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 2.5
                                    Lot = 2.0
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 2.5 AndAlso th <= 3.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(6)
                                    xx = LowNum2(6)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(7)
                                    xz = LowNum2(7)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.0
                                    Lot = 2.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.0 AndAlso th <= 3.5 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(7)
                                    xx = LowNum2(7)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(8)
                                    xz = LowNum2(8)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 3.5
                                    Lot = 3.0
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf th >= 3.5 AndAlso th <= 4.0 Then
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3b1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart11.GetUpperBound(0) - 1
                            If chart11(r, 0) <= 80 AndAlso chart11(r + 1, 0) >= 100 Then
                                For c As Integer = 0 To 9
                                    LowNum2(c) = chart11(r, c)
                                    UpNum2(c) = chart11(r + 1, c)

                                    zx = UpNum2(8)
                                    xx = LowNum2(8)
                                    uptx = 100
                                    lotx = 80
                                    x = Inter3ex(uptx, lotx, temp, xx, zx)

                                    zz = UpNum2(9)
                                    xz = LowNum2(9)
                                    uptz = 100
                                    lotz = 80
                                    z = Inter3ez(uptz, lotz, temp, xz, zz)
                                    upt = 4.0
                                    Lot = 3.5
                                    df = Inter3c1(upt, Lot, th, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End If

        If ComboBox173.Text = "Component Subject to PWHT" And Units.ComboBox1.Text = "SI" And TextBox91.Text = 101.6 And tmptkurangreftemptambahfatt = 56 Then
            df = "0"
        End If

        If df < 0 Then
            df = 0
        End If

    End Sub

    '885°F Embrittlement Coding perhitungan -------------------------------------------------------------------------------
    Public Sub delapandelapanlimafembrittlementdf()
        Dim tref As Double

        If Units.ComboBox1.Text = "SI" Then
            tref = 28
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            tref = 80
        End If

        Dim tminkurangtref As Double

        tminkurangtref = Val(TextBox46.Text) - tref

        Dim df As Double

        If Units.ComboBox1.Text = "SI" Then
            If tminkurangtref >= 57 Then
                df = 0
            End If

            Dim chart2(,) As Double = {{-56, 1381},
                {-44, 1216},
                {-33, 1022},
                {-22, 806},
                {-11, 581},
                {-0, 371},
                {11, 200},
                {22, 87},
                {33, 30},
                {44, 8},
                {56, 2}}

            Dim UpNum2(1) As Double
            Dim LowNum2(1) As Double

            Dim Lo1 As Double
            Dim Lo2 As Double
            Dim up1 As Double
            Dim up2 As Double

            Dim temp As Double
            Dim upt As Double
            Dim lot As Double
            Dim x As Double
            Dim z As Double

            temp = tminkurangtref

            If temp < -56 Then
                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= -44 AndAlso chart2(r + 1, 0) >= -56 Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)
                            x = LowNum2(1)
                            z = UpNum2(1)
                            upt = -44
                            lot = -56

                            df = Inter3b(upt, lot, temp, x, z)

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

                        df = Inter(temp, up1, up2, Lo1, Lo2)


                    Next
                    Exit For
                End If
            Next

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If tminkurangtref >= 101 Then
                df = 0
            End If


            Dim chart2(,) As Double = {{-100, 1381},
                {-80, 1216},
                {-60, 1022},
                {-40, 806},
                {-20, 581},
                {-0, 371},
                {20, 200},
                {40, 87},
                {60, 30},
                {80, 8},
                {100, 2}}

            Dim UpNum2(1) As Double
            Dim LowNum2(1) As Double

            Dim Lo1 As Double
            Dim Lo2 As Double
            Dim up1 As Double
            Dim up2 As Double

            Dim temp As Double
            Dim upt As Double
            Dim lot As Double
            Dim x As Double
            Dim z As Double

            temp = tminkurangtref

            If temp < -100 Then
                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                    If chart2(r, 0) <= -80 AndAlso chart2(r + 1, 0) >= -100 Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart2(r, c)
                            UpNum2(c) = chart2(r + 1, c)
                            x = LowNum2(1)
                            z = UpNum2(1)
                            upt = -80
                            lot = -100

                            df = Inter3b(upt, lot, temp, x, z)

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

                        df = Inter(temp, up1, up2, Lo1, Lo2)


                    Next
                    Exit For
                End If
            Next

        End If

        Label366.Text = df
    End Sub

    'Sigma Phase Coding perhitungan ---------------------------------------------------------------------------------------
    Public Sub sigmaphasedf()

        Dim df As Double

        If Units.ComboBox1.Text = "SI" Then
            If ComboBox166.Text = "(>1%, <5%)" Then
                Dim chart2(,) As Double = {{-46, 1.1},
                    {-18, 1.0},
                    {10, 0.9},
                    {38, 0.6},
                    {66, 0.3},
                    {93, 0.1},
                    {204, 0},
                    {316, 0},
                    {427, 0},
                    {538, 0},
                    {649, 0}}

                Dim UpNum2(1) As Double
                Dim LowNum2(1) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double

                Dim temp As Double
                Dim upt As Double
                Dim lot As Double
                Dim x As Double
                Dim z As Double

                temp = Val(TextBox47.Text)

                If temp < -46 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= -18 AndAlso chart2(r + 1, 0) >= -46 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = -18
                                lot = -46

                                df = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next

                ElseIf temp > 649 Then

                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= 538 AndAlso chart2(r + 1, 0) >= 649 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 649
                                lot = 538

                                df = Inter3c(upt, lot, temp, x, z)

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

                            df = Inter(temp, up1, up2, Lo1, Lo2)

                        Next
                        Exit For
                    End If
                Next

                If TextBox47.Text = 204 Then
                    df = 0
                End If


            ElseIf ComboBox166.Text = "(≥5%, <10%)" Then
                Dim chart3(,) As Double = {{-46, 34},
                {-18, 20},
                {10, 11},
                {38, 7},
                {66, 5},
                {93, 3},
                {204, 1.3},
                {316, 0.9},
                {427, 0.2},
                {538, 0},
                {649, 0}}

                Dim UpNum2(1) As Double
                Dim LowNum2(1) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double

                Dim temp As Double
                Dim upt As Double
                Dim lot As Double
                Dim x As Double
                Dim z As Double

                temp = Val(TextBox47.Text)

                If temp < -46 Then
                    For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                        If chart3(r, 0) <= -18 AndAlso chart3(r + 1, 0) >= -46 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart3(r, c)
                                UpNum2(c) = chart3(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = -18
                                lot = -46

                                df = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next

                ElseIf temp > 649 Then

                    For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                        If chart3(r, 0) <= 538 AndAlso chart3(r + 1, 0) >= 649 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart3(r, c)
                                UpNum2(c) = chart3(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 649
                                lot = 538

                                df = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                    If chart3(r, 0) <= temp AndAlso chart3(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart3(r, c)
                            UpNum2(c) = chart3(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            df = Inter(temp, up1, up2, Lo1, Lo2)

                        Next
                        Exit For
                    End If
                Next

                If TextBox47.Text = 538 Then
                    df = 0
                End If

            ElseIf ComboBox166.Text = "(≥10%)" Then
                Dim chart5(,) As Double = {{-46, 4196},
                {-18, 4196},
                {10, 4196},
                {38, 4196},
                {66, 3871},
                {93, 3202},
                {204, 1333},
                {316, 481},
                {427, 160},
                {538, 53},
                {649, 18}}

                Dim UpNum2(1) As Double
                Dim LowNum2(1) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double

                Dim temp As Double
                Dim upt As Double
                Dim lot As Double
                Dim x As Double
                Dim z As Double

                temp = Val(TextBox47.Text)

                If temp < -46 Then
                    For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                        If chart5(r, 0) <= -18 AndAlso chart5(r + 1, 0) >= -46 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart5(r, c)
                                UpNum2(c) = chart5(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = -18
                                lot = -46

                                df = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next

                ElseIf temp > 649 Then

                    For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                        If chart5(r, 0) <= 538 AndAlso chart5(r + 1, 0) >= 649 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart5(r, c)
                                UpNum2(c) = chart5(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 649
                                lot = 538

                                df = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart5.GetUpperBound(0) - 1
                    If chart5(r, 0) <= temp AndAlso chart5(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart5(r, c)
                            UpNum2(c) = chart5(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            df = Inter(temp, up1, up2, Lo1, Lo2)

                        Next
                        Exit For
                    End If
                Next
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If ComboBox166.Text = "(>1%, <5%)" Then
                Dim chart3(,) As Double = {{-50, 1.1},
                {0, 1},
                {50, 0.9},
                {100, 0.6},
                {150, 0.3},
                {200, 0.1},
                {400, 0},
                {600, 0},
                {800, 0},
                {1000, 0},
                {1200, 0}}

                Dim UpNum2(1) As Double
                Dim LowNum2(1) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double

                Dim temp As Double
                Dim upt As Double
                Dim lot As Double
                Dim x As Double
                Dim z As Double

                temp = Val(TextBox47.Text)

                If temp < -50 Then
                    For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                        If chart3(r, 0) <= -0 AndAlso chart3(r + 1, 0) >= -50 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart3(r, c)
                                UpNum2(c) = chart3(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = -0
                                lot = -50

                                df = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next

                ElseIf temp > 1200 Then

                    For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                        If chart3(r, 0) <= 1000 AndAlso chart3(r + 1, 0) >= 1200 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart3(r, c)
                                UpNum2(c) = chart3(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 1200
                                lot = 1000

                                df = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart3.GetUpperBound(0) - 1
                    If chart3(r, 0) <= temp AndAlso chart3(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart3(r, c)
                            UpNum2(c) = chart3(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            df = Inter(temp, up1, up2, Lo1, Lo2)

                        Next
                        Exit For
                    End If
                Next

                If TextBox47.Text = 400 Then
                    df = 0
                End If


            ElseIf ComboBox166.Text = "(≥5%, <10%)" Then
                Dim chart4(,) As Double = {{-50, 34},
                    {0, 20},
                    {50, 11},
                    {100, 7},
                    {150, 5},
                    {200, 3},
                    {400, 1.3},
                    {600, 0.9},
                    {800, 0.2},
                    {1000, 0},
                    {1200, 0}}

                Dim UpNum2(1) As Double
                Dim LowNum2(1) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double

                Dim temp As Double
                Dim upt As Double
                Dim lot As Double
                Dim x As Double
                Dim z As Double

                temp = Val(TextBox47.Text)

                If temp < -50 Then
                    For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                        If chart4(r, 0) <= -0 AndAlso chart4(r + 1, 0) >= -50 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart4(r, c)
                                UpNum2(c) = chart4(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = -0
                                lot = -50

                                df = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next

                ElseIf temp > 1200 Then

                    For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                        If chart4(r, 0) <= 1000 AndAlso chart4(r + 1, 0) >= 1200 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart4(r, c)
                                UpNum2(c) = chart4(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 1200
                                lot = 1000

                                df = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart4.GetUpperBound(0) - 1
                    If chart4(r, 0) <= temp AndAlso chart4(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart4(r, c)
                            UpNum2(c) = chart4(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            df = Inter(temp, up1, up2, Lo1, Lo2)

                        Next
                        Exit For
                    End If
                Next

                If TextBox47.Text = 1000 Then
                    df = 0
                End If

            ElseIf ComboBox166.Text = "(≥10%)" Then
            Dim chart6(,) As Double = {{-50, 4196},
                {0, 4196},
                {50, 4196},
                {100, 4196},
                {150, 3871},
                {200, 3202},
                {400, 1333},
                {600, 481},
                {800, 160},
                {1000, 53},
                {1200, 18}}

                Dim UpNum2(1) As Double
                Dim LowNum2(1) As Double

                Dim Lo1 As Double
                Dim Lo2 As Double
                Dim up1 As Double
                Dim up2 As Double

                Dim temp As Double
                Dim upt As Double
                Dim lot As Double
                Dim x As Double
                Dim z As Double

                temp = Val(TextBox47.Text)

                If temp < -50 Then
                    For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                        If chart6(r, 0) <= -0 AndAlso chart6(r + 1, 0) >= -50 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart6(r, c)
                                UpNum2(c) = chart6(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = -0
                                lot = -50

                                df = Inter3b(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next

                ElseIf temp > 1200 Then

                    For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                        If chart6(r, 0) <= 1000 AndAlso chart6(r + 1, 0) >= 1200 Then
                            For c As Integer = 0 To 1
                                LowNum2(c) = chart6(r, c)
                                UpNum2(c) = chart6(r + 1, c)
                                x = LowNum2(1)
                                z = UpNum2(1)
                                upt = 1200
                                lot = 1000

                                df = Inter3c(upt, lot, temp, x, z)

                            Next
                            Exit For
                        End If
                    Next
                End If

                For r As Integer = 0 To chart6.GetUpperBound(0) - 1
                    If chart6(r, 0) <= temp AndAlso chart6(r + 1, 0) >= temp Then
                        For c As Integer = 0 To 1
                            LowNum2(c) = chart6(r, c)
                            UpNum2(c) = chart6(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            df = Inter(temp, up1, up2, Lo1, Lo2)

                        Next
                        Exit For
                    End If
                Next
            End If
        End If

        Label367.Text = df
    End Sub

    'Mechanical Fatigue Coding perhitungan --------------------------------------------------------------------------------
    Public Sub mechanicalfatiguedf()
        Dim step1 As Double
        Dim step2 As Double
        Dim step3 As Double
        Dim step4 As Double

        If ComboBox65.Text = "None" Then
            step1 = 1
        ElseIf ComboBox65.Text = "One" Then
            step1 = 50
        ElseIf ComboBox65.Text = "Greater than one" Then
            step1 = 500
        End If

        If ComboBox66.Text = "Minor" Then
            step2 = 1
        ElseIf ComboBox66.Text = "Moderate" Then
            step2 = 50
        ElseIf ComboBox66.Text = "Severe" Then
            step2 = 500
        End If

        If ComboBox67.Text = "Shaking less than 2 weeks" Then
            step3 = 1
        ElseIf ComboBox67.Text = "Shaking between 2 and 13 weeks" Then
            step3 = 0.2
        ElseIf ComboBox67.Text = "Shaking between 13 and 52 weeks" Then
            step3 = 0.02
        End If

        If ComboBox68.Text = "Reciprocating Machinery" Then
            step4 = 50
        ElseIf ComboBox68.Text = "PRV Chatter" Then
            step4 = 25
        ElseIf ComboBox68.Text = "Valve with high pressure drop" Then
            step4 = 10
        ElseIf ComboBox68.Text = "None" Then
            step4 = 1
        End If

        Dim step2_3 As Double

        step2_3 = step2 * step3

        Dim basedf As Double

        If step1 > step2_3 AndAlso step1 > step4 Then
            basedf = step1
        End If
        If step2_3 > step1 AndAlso step2_3 > step4 Then
            basedf = step2_3
        End If
        If step4 > step1 AndAlso step4 > step2_3 Then
            basedf = step4
        End If
        If step1 = 1 AndAlso step2_3 = 1 AndAlso step4 = 1 Then
            basedf = 1
        End If

        Dim adj1 As Double
        Dim adj2 As Double
        Dim adj3 As Double
        Dim adj4 As Double
        Dim adj5 As Double

        If ComboBox69.Text = "Modification based on complete engineering analysis" Then
            adj1 = 0.002
        ElseIf ComboBox69.Text = "Modification based on experience" Then
            adj1 = 0.2
        ElseIf ComboBox69.Text = "No modifications" Then
            adj1 = 2
        End If

        If ComboBox70.Text = "0 to 5 total pipe fittings" Then
            adj2 = 0.5
        ElseIf ComboBox70.Text = "6 to 10 total pipe fittings" Then
            adj2 = 1
        ElseIf ComboBox70.Text = "Greater than 10 total pipe fittings" Then
            adj2 = 2
        End If

        If ComboBox71.Text = "Missing or damaged supports, improper support" Then
            adj3 = 2
        ElseIf ComboBox71.Text = "Broken gussets, gussets welded directly to the pipe" Then
            adj3 = 2
        ElseIf ComboBox71.Text = "Good Condition" Then
            adj3 = 1
        End If

        If ComboBox72.Text = "Threaded, socketweld, saddle on" Then
            adj4 = 2
        ElseIf ComboBox72.Text = "Saddle in fittings" Then
            adj4 = 1
        ElseIf ComboBox72.Text = "Piping tee, Weldolets" Then
            adj4 = 0.2
        ElseIf ComboBox72.Text = "Sweepolets" Then
            adj4 = 0.02
        End If

        If ComboBox73.Text = "All branches less than or equal to 2 NPS" Then
            adj5 = 1
        ElseIf ComboBox73.Text = "Any branch greater than 2 NPS" Then
            adj5 = 0.02
        End If

        Dim dfmechanicalfatigue As Double

        dfmechanicalfatigue = basedf * adj1 * adj2 * adj3 * adj4 * adj5

        Label378.Text = dfmechanicalfatigue

    End Sub


End Class
