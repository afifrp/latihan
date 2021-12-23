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

    Private Sub ComboBox64_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox64.SelectedIndexChanged
        If ComboBox62.Text = "Pressure Vessel" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "Compressors" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "Heat Exchanger" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "AirFin Heat Exchanger Header Boxes" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "Pipes & Tubes" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "Pumps" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "Atmospheric Storage Tank - Shell Courses" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "Atmospheric Storage Tank - Bottom Plates" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Hide()
        ElseIf ComboBox62.Text = "Pressure Relief Devices" Then
            Panelgeneraldata.Hide()
            PanelgeneraldataHE.Hide()
            PanelgeneraldataPRD.Show()
        ElseIf ComboBox62.Text = "Heat Exchanger Tube Bundles" Then
            Panelgeneraldata.Show()
            PanelgeneraldataHE.Show()
            PanelgeneraldataPRD.Hide()
        End If
    End Sub

    Private Sub ComboBox62_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox62.SelectedIndexChanged

        ComboBox63.Items.Clear()


        If ComboBox62.Text = "Pressure Vessel" Then
            ComboBox63.Items.Add("KODRUM")
            ComboBox63.Items.Add("DRUM")
            ComboBox63.Items.Add("REACTOR")
            ComboBox63.Items.Add("COLTOP")
            ComboBox63.Items.Add("COLMID")
            ComboBox63.Items.Add("COLBTM")
        ElseIf ComboBox62.Text = "Compressors" Then
            ComboBox63.Items.Add("COMPC")
            ComboBox63.Items.Add("COMPR")
        ElseIf ComboBox62.Text = "Heat Exchanger" Then
            ComboBox63.Items.Add("HEXSS")
            ComboBox63.Items.Add("HEXTS")
            ComboBox63.Items.Add("HEXTUBE")
        ElseIf ComboBox62.Text = "AirFin Heat Exchanger Header Boxes" Then
            ComboBox63.Items.Add("FINFAN")
            ComboBox63.Items.Add("FILTER")
        ElseIf ComboBox62.Text = "Pipes & Tubes" Then
            ComboBox63.Items.Add("PIPE-1")
            ComboBox63.Items.Add("PIPE-3")
            ComboBox63.Items.Add("PIPE-4")
            ComboBox63.Items.Add("PIPE-6")
            ComboBox63.Items.Add("PIPE-8")
            ComboBox63.Items.Add("PIPE-10")
            ComboBox63.Items.Add("PIPE-13")
            ComboBox63.Items.Add("PIPE-16")
            ComboBox63.Items.Add("PIPEGT16")
        ElseIf ComboBox62.Text = "Pumps" Then
            ComboBox63.Items.Add("PUMP3S")
            ComboBox63.Items.Add("PUMP1S")
            ComboBox63.Items.Add("PUMPR")
        ElseIf ComboBox62.Text = "Atmospheric Storage Tank - Shell Courses" Then
            ComboBox63.Items.Add("COURSES-1-10")
        ElseIf ComboBox62.Text = "Atmospheric Storage Tank - Bottom Plates" Then
            ComboBox63.Items.Add("TANKBOTTOM")
        ElseIf ComboBox62.Text = "Pressure Relief Devices" Then
            ComboBox63.Items.Add("Pressure Relief Devices")
        ElseIf ComboBox62.Text = "Heat Exchanger Tube Bundles" Then
            ComboBox63.Items.Add("Heat Exchanger Tube Bundles")
        End If


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
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            CheckBox6.Enabled = False
        Else
            CheckBox6.Enabled = True
        End If

    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            CheckBox5.Enabled = False
        Else
            CheckBox5.Enabled = True
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
        Else
            CheckBox8.Enabled = True
            CheckBox9.Enabled = True
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            CheckBox7.Enabled = False
            CheckBox9.Enabled = False
        Else
            CheckBox7.Enabled = True
            CheckBox9.Enabled = True
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
        Else
            CheckBox7.Enabled = True
            CheckBox8.Enabled = True
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            CheckBox11.Enabled = False
        Else
            CheckBox11.Enabled = True
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            CheckBox10.Enabled = False
        Else
            CheckBox10.Enabled = True
        End If
    End Sub

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

    Private Sub CheckBox33_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox33.Checked = True Then
            CheckBox34.Enabled = False
            Panel44.Visible = True
        Else
            CheckBox34.Enabled = True
            Panel44.Visible = False
        End If
    End Sub

    Private Sub CheckBox34_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox34.Checked = True Then
            CheckBox33.Enabled = False
            ComboBox8.Text = "NONE SUSCEPTIBILITY"
        Else
            CheckBox33.Enabled = True
            ComboBox8.Text = ""
        End If
    End Sub

    'SCC Sulfide Stress Cracking DF
    Private Sub CheckBox50_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox50.CheckedChanged
        If CheckBox50.Checked = True Then
            CheckBox51.Enabled = False
            CheckBox52.Enabled = False
            CheckBox53.Enabled = False
            CheckBox54.Enabled = False
        Else
            CheckBox51.Enabled = True
            CheckBox52.Enabled = True
            CheckBox53.Enabled = True
            CheckBox54.Enabled = True

        End If
    End Sub

    Private Sub CheckBox51_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox51.CheckedChanged
        If CheckBox51.Checked = True Then
            CheckBox50.Enabled = False
            CheckBox52.Enabled = False
            CheckBox53.Enabled = False
            CheckBox54.Enabled = False
        Else
            CheckBox50.Enabled = True
            CheckBox52.Enabled = True
            CheckBox53.Enabled = True
            CheckBox54.Enabled = True

        End If
    End Sub

    Private Sub CheckBox52_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox52.CheckedChanged
        If CheckBox52.Checked = True Then
            CheckBox51.Enabled = False
            CheckBox50.Enabled = False
            CheckBox53.Enabled = False
            CheckBox54.Enabled = False
        Else
            CheckBox51.Enabled = True
            CheckBox50.Enabled = True
            CheckBox53.Enabled = True
            CheckBox54.Enabled = True

        End If
    End Sub

    Private Sub CheckBox53_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox53.CheckedChanged
        If CheckBox53.Checked = True Then
            CheckBox51.Enabled = False
            CheckBox52.Enabled = False
            CheckBox50.Enabled = False
            CheckBox54.Enabled = False
        Else
            CheckBox51.Enabled = True
            CheckBox52.Enabled = True
            CheckBox50.Enabled = True
            CheckBox54.Enabled = True

        End If
    End Sub

    Private Sub CheckBox54_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox54.CheckedChanged
        If CheckBox54.Checked = True Then
            CheckBox51.Enabled = False
            CheckBox52.Enabled = False
            CheckBox53.Enabled = False
            CheckBox50.Enabled = False
        Else
            CheckBox51.Enabled = True
            CheckBox52.Enabled = True
            CheckBox53.Enabled = True
            CheckBox50.Enabled = True

        End If
    End Sub

    Private Sub CheckBox55_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox55.CheckedChanged
        If CheckBox55.Checked = True Then
            CheckBox56.Enabled = False
            CheckBox57.Enabled = False
            CheckBox58.Enabled = False

        Else
            CheckBox56.Enabled = True
            CheckBox57.Enabled = True
            CheckBox58.Enabled = True

        End If
    End Sub

    Private Sub CheckBox56_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox56.CheckedChanged
        If CheckBox56.Checked = True Then
            CheckBox55.Enabled = False
            CheckBox57.Enabled = False
            CheckBox58.Enabled = False

        Else
            CheckBox55.Enabled = True
            CheckBox57.Enabled = True
            CheckBox58.Enabled = True

        End If
    End Sub

    Private Sub CheckBox57_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox57.CheckedChanged
        If CheckBox57.Checked = True Then
            CheckBox56.Enabled = False
            CheckBox55.Enabled = False
            CheckBox58.Enabled = False

        Else
            CheckBox56.Enabled = True
            CheckBox55.Enabled = True
            CheckBox58.Enabled = True

        End If
    End Sub

    Private Sub CheckBox58_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox58.CheckedChanged
        If CheckBox58.Checked = True Then
            CheckBox56.Enabled = False
            CheckBox57.Enabled = False
            CheckBox55.Enabled = False

        Else
            CheckBox56.Enabled = True
            CheckBox57.Enabled = True
            CheckBox55.Enabled = True

        End If
    End Sub

    Private Sub CheckBox59_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox59.CheckedChanged
        If CheckBox59.Checked = True Then
            CheckBox60.Enabled = False
            CheckBox61.Enabled = True
            CheckBox62.Enabled = True
            CheckBox63.Enabled = True
        Else
            CheckBox60.Enabled = True
            CheckBox61.Enabled = False
            CheckBox62.Enabled = False
            CheckBox63.Enabled = False
        End If
    End Sub

    Private Sub CheckBox60_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox60.CheckedChanged
        If CheckBox60.Checked = True Then
            CheckBox59.Enabled = False
            CheckBox61.Enabled = True
            CheckBox62.Enabled = True
            CheckBox63.Enabled = True
        Else
            CheckBox59.Enabled = True
            CheckBox61.Enabled = False
            CheckBox62.Enabled = False
            CheckBox63.Enabled = False
        End If
    End Sub

    Private Sub CheckBox61_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox61.CheckedChanged
        If CheckBox61.Checked = True Then
            CheckBox62.Enabled = False
            CheckBox63.Enabled = False
        Else
            CheckBox62.Enabled = True
            CheckBox63.Enabled = True
        End If
    End Sub

    Private Sub CheckBox62_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox62.CheckedChanged
        If CheckBox62.Checked = True Then
            CheckBox61.Enabled = False
            CheckBox63.Enabled = False
        Else
            CheckBox61.Enabled = True
            CheckBox63.Enabled = True
        End If
    End Sub

    Private Sub CheckBox63_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox63.CheckedChanged
        If CheckBox63.Checked = True Then
            CheckBox61.Enabled = False
            CheckBox63.Enabled = False
        Else
            CheckBox61.Enabled = True
            CheckBox63.Enabled = True
        End If
    End Sub

    Private Sub CheckBox64_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox64.CheckedChanged
        If CheckBox64.Checked = True Then
            CheckBox65.Enabled = False
            ComboBox10.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox65.Enabled = True
            ComboBox10.Text = ""
        End If
    End Sub

    Private Sub CheckBox65_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox65.CheckedChanged
        If CheckBox65.Checked = True Then
            CheckBox64.Enabled = False

        Else
            CheckBox64.Enabled = True
        End If
    End Sub

    'SCC-HIC/SOHIC-H2S DF
    Private Sub CheckBox66_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox66.CheckedChanged
        If CheckBox66.Checked = True Then
            CheckBox67.Enabled = False
            CheckBox68.Enabled = False
        Else
            CheckBox67.Enabled = True
            CheckBox68.Enabled = True
        End If
    End Sub

    Private Sub CheckBox67_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox67.CheckedChanged
        If CheckBox67.Checked = True Then
            CheckBox66.Enabled = False
            CheckBox68.Enabled = False
        Else
            CheckBox66.Enabled = True
            CheckBox68.Enabled = True
        End If
    End Sub

    Private Sub CheckBox68_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox68.CheckedChanged
        If CheckBox68.Checked = True Then
            CheckBox67.Enabled = False
            CheckBox66.Enabled = False
        Else
            CheckBox67.Enabled = True
            CheckBox66.Enabled = True
        End If
    End Sub

    Private Sub CheckBox69_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox69.CheckedChanged
        If CheckBox69.Checked = True Then
            CheckBox70.Enabled = False
            ComboBox11.Text = "HIGH SUSCEPTIBILITY"
        Else
            CheckBox70.Enabled = True
            ComboBox11.Text = ""
        End If
    End Sub

    'Scc Damage Factor – Alkaline Carbonate Stress Corrosion Cracking
    Private Sub CheckBox77_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox77.CheckedChanged
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

    Private Sub CheckBox78_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox78.CheckedChanged
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

    Private Sub CheckBox79_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox79.CheckedChanged
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

    Private Sub CheckBox80_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox80.CheckedChanged
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

    Private Sub CheckBox81_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox81.CheckedChanged
        If CheckBox81.Checked = True Then
            CheckBox82.Checked = False
        Else
            CheckBox82.Checked = True
        End If
    End Sub

    Private Sub CheckBox82_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox82.CheckedChanged
        If CheckBox82.Checked = True Then
            CheckBox81.Checked = False
        Else
            CheckBox81.Checked = True
        End If
    End Sub

    Private Sub CheckBox83_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox83.CheckedChanged
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

    Private Sub CheckBox84_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox84.CheckedChanged
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

    Private Sub CheckBox98_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox98.CheckedChanged
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

    Private Sub CheckBox99_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox99.CheckedChanged
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

    Private Sub CheckBox100_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox100.CheckedChanged
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

    Private Sub CheckBox101_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox101.CheckedChanged
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

    Private Sub CheckBox102_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox102.CheckedChanged
        If CheckBox102.Checked = True Then
            CheckBox103.Checked = False
        Else
            CheckBox103.Checked = True
        End If
    End Sub

    Private Sub CheckBox103_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox103.CheckedChanged
        If CheckBox103.Checked = True Then
            CheckBox102.Checked = False
        Else
            CheckBox102.Checked = True
        End If
    End Sub

    Private Sub CheckBox104_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox104.CheckedChanged
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

    Private Sub CheckBox105_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox105.CheckedChanged
        If CheckBox105.Checked = True Then
            CheckBox104.Checked = False
        Else
            CheckBox104.Checked = True
        End If
    End Sub

    Private Sub CheckBox106_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox106.CheckedChanged
        If CheckBox106.Checked = True Then
            CheckBox107.Checked = False
        Else
            CheckBox107.Checked = True
        End If
    End Sub

    Private Sub CheckBox107_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox107.CheckedChanged
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
            CheckBox117.Visible = True
            CheckBox118.Visible = True
            CheckBox119.Visible = True
            CheckBox120.Visible = True
            CheckBox121.Visible = True
        Else
            CheckBox116.Checked = True
            Label163.Visible = False
            Label164.Visible = False
            CheckBox117.Visible = False
            CheckBox118.Visible = False
            CheckBox119.Visible = False
            CheckBox120.Visible = False
            CheckBox121.Visible = False

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

    Private Sub CheckBox117_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox117.CheckedChanged
        If CheckBox117.Checked = True Then
            CheckBox118.Checked = False
        Else
            CheckBox118.Checked = True
        End If
    End Sub

    Private Sub CheckBox118_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox118.CheckedChanged
        If CheckBox118.Checked = True Then
            CheckBox117.Checked = False
        Else
            CheckBox117.Checked = True
        End If
    End Sub

    Private Sub CheckBox119_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox119.CheckedChanged
        If CheckBox119.Checked = True Then
            CheckBox120.Checked = False
            CheckBox121.Checked = False
        Else
            CheckBox120.Checked = True
            CheckBox121.Checked = True
        End If
    End Sub

    Private Sub CheckBox120_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox120.CheckedChanged
        If CheckBox120.Checked = True Then
            CheckBox119.Checked = False
            CheckBox121.Checked = False
        Else
            CheckBox119.Checked = True
            CheckBox121.Checked = True
        End If
    End Sub

    Private Sub CheckBox121_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox121.CheckedChanged
        If CheckBox121.Checked = True Then
            CheckBox120.Checked = False
            CheckBox119.Checked = False
        Else
            CheckBox120.Checked = True
            CheckBox119.Checked = True
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

    Private Sub ComboBox63_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox63.SelectedIndexChanged
        If ComboBox63.Text = "COMPC" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0
            totalfms.Text = 0.00003
        ElseIf ComboBox63.Text = "COMPR" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "HEXSS" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "HEXTS" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-1" Then
            small.Text = 0.000028
            medium.Text = 0
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-2" Then
            small.Text = 0.000028
            medium.Text = 0
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-4" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-6" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0
            rupture.Text = 0.0000026
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-8" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-10" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-12" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPE-16" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PIPEGT16" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PUMP2S" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PUMP1S" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "PUMPR" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "TANKBOTTOM" Then
            small.Text = 0.00072
            medium.Text = 0
            large.Text = 0
            rupture.Text = 0.000002
            totalfms.Text = 0.00072
        ElseIf ComboBox63.Text = "COURSES-1-10" Then
            small.Text = 0.00007
            medium.Text = 0.000025
            large.Text = 0.000005
            rupture.Text = 0.0000001
            totalfms.Text = 0.0001
        ElseIf ComboBox63.Text = "KODRUM" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "COLBTM" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "FINFAN" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "FILTER" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "DRUM" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "REACTOR" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "COLTOP" Then
            small.Text = 0.000008
            medium.Text = 0.00002
            large.Text = 0.000002
            rupture.Text = 0.0000006
            totalfms.Text = 0.0000306
        ElseIf ComboBox63.Text = "COLMID" Then
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
End Class
