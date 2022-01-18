Public Class Corrosion_Rate

    'interpolasi
    Private Function Inter(ByVal ph As Double, ByVal Up1 As Double, ByVal Up2 As Double, ByVal Lo1 As Double, ByVal Lo2 As Double) As Double
        Inter = ((ph - Lo1) * (Up2 - Lo2) / (Up1 - Lo1)) + Lo2
    End Function

    Private Function Inter1(ByVal ph As Double, ByVal Up3 As Double, ByVal Up4 As Double, ByVal Lo3 As Double, ByVal Lo4 As Double) As Double
        Inter1 = ((ph - Lo3) * (Up4 - Lo4) / (Up3 - Lo3)) + Lo4
    End Function

    Private Function Inter3a(ByVal upt As Double, ByVal Lot As Double, ByVal temp As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3a = ((temp - Lot) * (z - x) / (upt - Lot)) + x
    End Function

    'extrapolation lower

    Private Function Inter3b(ByVal upt As Double, ByVal Lot As Double, ByVal temp As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3b = ((temp - upt) * (x - z) / (Lot - upt)) + z
    End Function

    'extrapolation upper
    Private Function Inter3c(ByVal upt As Double, ByVal Lot As Double, ByVal temp As Double, ByVal x As Double, ByVal z As Double) As Double
        Inter3c = ((temp - Lot) * (z - x) / (upt - Lot)) + x
    End Function

    Private Function Inter3da(ByVal upta As Double, ByVal Lota As Double, ByVal ph As Double, ByVal xa As Double, ByVal za As Double) As Double
        Inter3da = ((ph - upta) * (xa - za) / (Lota - upta)) + za
    End Function

    Private Function Inter3db(ByVal uptb As Double, ByVal Lotb As Double, ByVal ph As Double, ByVal xb As Double, ByVal zb As Double) As Double
        Inter3db = ((ph - uptb) * (xb - zb) / (Lotb - uptb)) + zb
    End Function

    Private Function Inter3ex(ByVal uptx As Double, ByVal Lotx As Double, ByVal ph As Double, ByVal xx As Double, ByVal zx As Double) As Double
        Inter3ex = ((ph - Lotx) * (zx - xx) / (uptx - Lotx)) + xx
    End Function

    Private Function Inter3ez(ByVal uptz As Double, ByVal Lotz As Double, ByVal ph As Double, ByVal xz As Double, ByVal zz As Double) As Double
        Inter3ez = ((ph - Lotz) * (zz - xz) / (uptz - Lotz)) + xz
    End Function

    Private Function Inter4(ByVal clcon As Double, ByVal Up1 As Double, ByVal Up2 As Double, ByVal Lo1 As Double, ByVal Lo2 As Double) As Double
        Inter4 = ((clcon - Lo1) * (Up2 - Lo2) / (Up1 - Lo1)) + Lo2
    End Function

    Private Function Inter5(ByVal clcon As Double, ByVal Up3 As Double, ByVal Up4 As Double, ByVal Lo3 As Double, ByVal Lo4 As Double) As Double
        Inter5 = ((clcon - Lo3) * (Up4 - Lo4) / (Up3 - Lo3)) + Lo4
    End Function

    '==================================================================================================================================

    Private Sub checkedfalse()
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False
        Panel15.Visible = False
        Panel16.Visible = False
        Panel17.Visible = False
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        CheckBox6.Enabled = True
        CheckBox7.Enabled = True
        CheckBox8.Enabled = True
        CheckBox9.Enabled = True
        CheckBox10.Enabled = True
        CheckBox11.Enabled = True
        CheckBox12.Enabled = True
        CheckBox13.Enabled = True
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            Panel5.Visible = True
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = True
        Else
            Call checkedfalse()
        End If

        If Units.ComboBox1.Text = "SI" Then
            ComboBox6.Text = ""
            ComboBox6.Items.Clear()
            ComboBox6.Items.Add("< 49")
            ComboBox6.Items.Add("49 - 104")
            ComboBox6.Items.Add("> 104")
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            ComboBox6.Text = ""
            ComboBox6.Items.Clear()
            ComboBox6.Items.Add("< 120")
            ComboBox6.Items.Add("120 - 220")
            ComboBox6.Items.Add("> 220")
        End If
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = True
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = True
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = True
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = True
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If

        If Units.ComboBox1.Text = "SI" Then
            ComboBox14.Text = ""
            ComboBox14.Items.Clear()
            ComboBox14.Items.Add("< 49")
            ComboBox14.Items.Add("49 - 104")
            ComboBox14.Items.Add("> 104")
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            ComboBox14.Text = ""
            ComboBox14.Items.Clear()
            ComboBox14.Items.Add("< 120")
            ComboBox14.Items.Add("120 - 220")
            ComboBox14.Items.Add("> 220")
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = True
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = True
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = True
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = True
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = True
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = True
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = True
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = True
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = True
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = True
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = True
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = True
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If

        If Units.ComboBox1.Text = "SI" Then
            ComboBox30.Text = ""
            ComboBox30.Items.Clear()
            ComboBox30.Items.Add("< 10")
            ComboBox30.Items.Add("≥ 10")
        ElseIf Units.ComboBox1.Text = "US Customary" Then
            ComboBox30.Text = ""
            ComboBox30.Items.Clear()
            ComboBox30.Items.Add("< 3.05")
            ComboBox30.Items.Add("≥ 3.05")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = True
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = True
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = True
            Panel16.Visible = False
            Panel17.Visible = False
            CheckBox1.Enabled = True
            CheckBox2.Enabled = False
            CheckBox3.Enabled = True
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = True
            Panel17.Visible = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = True
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Panel5.Visible = False
            Panel6.Visible = False
            Panel7.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
            Panel10.Visible = False
            Panel11.Visible = False
            Panel12.Visible = False
            Panel13.Visible = False
            Panel14.Visible = False
            Panel15.Visible = False
            Panel16.Visible = False
            Panel17.Visible = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            Call checkedfalse()
        End If
    End Sub

    Private Sub panggillocalorgeneral()
        If Home.RadioButton1.Checked = True Then
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = False
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            CheckBox6.Enabled = True
            CheckBox7.Enabled = True
            CheckBox8.Enabled = False
            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox11.Enabled = True
            CheckBox12.Enabled = True
            CheckBox13.Enabled = True
        ElseIf Home.RadioButton2.Checked = True Then
            CheckBox1.Enabled = False
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            CheckBox5.Enabled = False
            CheckBox6.Enabled = True
            CheckBox7.Enabled = True
            CheckBox8.Enabled = True
            CheckBox9.Enabled = True
            CheckBox10.Enabled = True
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
        Else
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            CheckBox6.Enabled = False
            CheckBox7.Enabled = False
            CheckBox8.Enabled = False
            CheckBox9.Enabled = False
            CheckBox10.Enabled = False
            CheckBox11.Enabled = False
            CheckBox12.Enabled = False
            CheckBox13.Enabled = False
            MsgBox("Select thining type local/general !")
        End If
    End Sub

    Private Sub Corrosion_Rate_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call panggillocalorgeneral()
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        If CheckBox14.Checked = True Then
            If Units.ComboBox1.Text = "SI" Then
                TextBox1.ReadOnly = True
                TextBox1.Text = 0.13
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                TextBox1.ReadOnly = True
                TextBox1.Text = 5
            End If
        Else
            TextBox1.ReadOnly = False
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = True Then
            If Units.ComboBox1.Text = "SI" Then
                TextBox2.ReadOnly = True
                TextBox2.Text = 0.05
            ElseIf Units.ComboBox1.Text = "US Customary" Then
                TextBox2.ReadOnly = True
                TextBox2.Text = 2
            End If
        Else
            TextBox2.ReadOnly = False
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox16.SelectedIndexChanged
        If ComboBox16.Text = "No Coating" Then
            ComboBox17.Enabled = False
            ComboBox18.Enabled = False
            ComboBox19.Enabled = False
        Else
            ComboBox17.Enabled = True
            ComboBox18.Enabled = True
            ComboBox19.Enabled = True
        End If
    End Sub

    Private Sub ComboBox20_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox20.SelectedIndexChanged
        If ComboBox20.Text = "Carbon Steel" Then
            Label27.Visible = True
            TextBox5.Visible = True
            Label29.Visible = False
            ComboBox21.Visible = False
        ElseIf ComboBox20.Text = "Type 304, 316, 321, 347 Series Stainless Steels" Then
            Label27.Visible = True
            TextBox5.Visible = True
            Label29.Visible = False
            ComboBox21.Visible = False
        ElseIf ComboBox20.Text = "Alloy 825" Then
            Label27.Visible = False
            TextBox5.Visible = False
            Label29.Visible = False
            ComboBox21.Visible = False
        ElseIf ComboBox20.Text = "Alloy 20" Then
            Label27.Visible = False
            TextBox5.Visible = False
            Label29.Visible = False
            ComboBox21.Visible = False
        ElseIf ComboBox20.Text = "Alloy 625" Then
            Label27.Visible = False
            TextBox5.Visible = False
            Label29.Visible = False
            ComboBox21.Visible = False
        ElseIf ComboBox20.Text = "Alloy C-276" Then
            Label27.Visible = False
            TextBox5.Visible = False
            Label29.Visible = False
            ComboBox21.Visible = False
        ElseIf ComboBox20.Text = "Alloy B-2" Then
            Label27.Visible = False
            TextBox5.Visible = False
            Label29.Visible = True
            ComboBox21.Visible = True
        ElseIf ComboBox20.Text = "Alloy 400" Then
            Label27.Visible = False
            TextBox5.Visible = False
            Label29.Visible = True
            ComboBox21.Visible = True

        Else
            Label27.Visible = False
            TextBox5.Visible = False
            Label29.Visible = False
            ComboBox21.Visible = False
        End If
    End Sub

    Private Sub ComboBox24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox24.SelectedIndexChanged
        If ComboBox24.Text = "12Cr Steels" Then
            Label35.Visible = False
            ComboBox25.Visible = False
        ElseIf ComboBox24.Text = "Type 304, 304L, 316, 316L, 321, 347 Stainless Steel" Then
            Label35.Visible = False
            ComboBox25.Visible = False
        Else
            Label35.Visible = True
            ComboBox25.Visible = True
        End If
    End Sub

    Private Sub ComboBox28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox28.SelectedIndexChanged
        If ComboBox28.Text = "≤ 80" Then
            TextBox15.Visible = True
            ComboBox29.Visible = False
            ComboBox29.Text = ""
        ElseIf ComboBox28.Text = "> 80" Then
            TextBox15.Visible = False
            ComboBox29.Visible = True
            TextBox15.Text = ""
        End If
    End Sub

    Private Sub ComboBox27_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox27.SelectedIndexChanged
        If ComboBox28.Text = "Carbon Steel" Then
            Label43.Visible = True
            ComboBox30.Visible = True
            Label47.Visible = False
            ComboBox31.Visible = False
            ComboBox31.Text = ""
        ElseIf ComboBox28.Text = "Alloy 400" Then
            Label43.Visible = False
            ComboBox30.Visible = False
            Label47.Visible = True
            ComboBox31.Visible = True
            ComboBox30.Text = ""
        End If
    End Sub

    Private Sub ComboBox33_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox33.SelectedIndexChanged
        If ComboBox33.Text = "MEA" Then
            ComboBox34.Text = ""
            ComboBox34.Items.Clear()
            ComboBox34.Items.Add("≤ 20")
            ComboBox34.Items.Add("21-25")
            ComboBox34.Items.Add("> 25")
        End If
        If ComboBox33.Text = "DEA" Then
            ComboBox34.Text = ""
            ComboBox34.Items.Clear()
            ComboBox34.Items.Add("≤ 30")
            ComboBox34.Items.Add("31-40")
            ComboBox34.Items.Add("> 40")
        End If
        If ComboBox33.Text = "MDEA" Then
            ComboBox34.Text = ""
            ComboBox34.Items.Clear()
            ComboBox34.Items.Add("≤ 50")

        End If
    End Sub

    'cr ast bottom
    Public Sub astbottom()
        Dim resistivy As Double

        If ComboBox1.Text = "< 500" Then
            resistivy = 1.5
        End If
        If ComboBox1.Text = "500 – 1000" Then
            resistivy = 1.25
        End If
        If ComboBox1.Text = "1000 – 2000" Then
            resistivy = 1.0
        End If
        If ComboBox1.Text = "2000 – 10000" Then
            resistivy = 0.83
        End If
        If ComboBox1.Text = "> 10000" Then
            resistivy = 0.66
        End If
        If ComboBox1.Text = "AST with RPB (Release Prevention Barrier)" Then
            resistivy = 1.0
        End If

        Dim astpadtype As Double

        If ComboBox2.Text = "Soil With High Salt" Then
            astpadtype = 1.5
        End If
        If ComboBox2.Text = "Crushed Limestone" Then
            astpadtype = 1.4
        End If
        If ComboBox2.Text = "Native Soil" Then
            astpadtype = 1.3
        End If
        If ComboBox2.Text = "Construction Grade Sand" Then
            astpadtype = 1.15
        End If
        If ComboBox2.Text = "Continuous Asphalt" Then
            astpadtype = 1.0
        End If
        If ComboBox2.Text = "Continuous Concrete" Then
            astpadtype = 1.0
        End If
        If ComboBox2.Text = "Oil Sand" Then
            astpadtype = 0.7
        End If
        If ComboBox2.Text = "High Resistivity Low Chloride Sand" Then
            astpadtype = 0.7
        End If

        Dim astdrainagetype As Double

        If ComboBox3.Text = "One Third Frequently Underwater" Then
            astdrainagetype = 3
        End If
        If ComboBox3.Text = "Storm Water Collects At AST Base" Then
            astdrainagetype = 2
        End If
        If ComboBox3.Text = "Storm Water Does Not Collect At AST Base" Then
            astdrainagetype = 1
        End If

        Dim cathodic As Double

        If ComboBox4.Text = "None" Then
            cathodic = 1.0
        End If
        If ComboBox4.Text = "Yes Not Per API 651" Then
            cathodic = 0.66
        End If
        If ComboBox4.Text = "Yes Per API 651" Then
            cathodic = 0.33
        End If

        Dim astbottomtype As Double

        If ComboBox5.Text = "RPB Not Per API 650" Then
            astbottomtype = 1.4
        End If
        If ComboBox5.Text = "RPB Per API 650" Then
            astbottomtype = 1.0
        End If
        If ComboBox5.Text = "Single Bottom" Then
            astbottomtype = 1.0
        End If

        Dim soilsidetemp As Double

        If ComboBox6.Text = "Temp ≤ 24" Then
            soilsidetemp = 1.0
        ElseIf ComboBox6.Text = "24 < Temp ≤ 66" Then
            soilsidetemp = 1.1
        ElseIf ComboBox6.Text = "66 < Temp ≤ 93" Then
            soilsidetemp = 1.3
        ElseIf ComboBox6.Text = "93 < Temp ≤ 121" Then
            soilsidetemp = 1.4
        ElseIf ComboBox6.Text = "> 121" Then
            soilsidetemp = 1.0

        ElseIf ComboBox6.Text = "Temp ≤ 75" Then
            soilsidetemp = 1.0
        ElseIf ComboBox6.Text = "75 < Temp ≤ 150" Then
            soilsidetemp = 1.1
        ElseIf ComboBox6.Text = "150 < Temp ≤ 200" Then
            soilsidetemp = 1.3
        ElseIf ComboBox6.Text = "200 < Temp ≤ 250" Then
            soilsidetemp = 1.4
        ElseIf ComboBox6.Text = "> 250" Then
            soilsidetemp = 1.0
        End If

        Dim modifiedsoilside As Double

        modifiedsoilside = Val(TextBox1.Text) * resistivy * astpadtype * astdrainagetype * cathodic * astbottomtype * soilsidetemp

        Dim productside As Double

        If ComboBox7.Text = "Wet" Then
            productside = 2.5
        End If
        If ComboBox7.Text = "Dry" Then
            productside = 1.0
        End If

        Dim productsidetemp As Double

        If ComboBox8.Text = "Temp ≤ 24" Then
            productsidetemp = 1.0
        ElseIf ComboBox8.Text = "24 < Temp ≤ 66" Then
            productsidetemp = 1.1
        ElseIf ComboBox8.Text = "66 < Temp ≤ 93" Then
            productsidetemp = 1.3
        ElseIf ComboBox8.Text = "93 < Temp ≤ 121" Then
            productsidetemp = 1.4
        ElseIf ComboBox8.Text = "> 121" Then
            productsidetemp = 1.0

        ElseIf ComboBox8.Text = "Temp ≤ 75" Then
            productsidetemp = 1.0
        ElseIf ComboBox8.Text = "75 < Temp ≤ 150" Then
            productsidetemp = 1.1
        ElseIf ComboBox8.Text = "150 < Temp ≤ 200" Then
            productsidetemp = 1.3
        ElseIf ComboBox8.Text = "200 < Temp ≤ 250" Then
            productsidetemp = 1.4
        ElseIf ComboBox8.Text = "> 250" Then
            productsidetemp = 1.0
        End If

        Dim steamcoil As Double

        If ComboBox9.Text = "No" Then
            steamcoil = 1.0
        ElseIf ComboBox9.Text = "Yes" Then
            steamcoil = 1.15
        End If

        Dim waterdrawoff As Double

        If ComboBox10.Text = "No" Then
            waterdrawoff = 1.0
        ElseIf ComboBox10.Text = "Yes" Then
            waterdrawoff = 0.7
        End If

        Dim modifiedproductside As Double

        modifiedproductside = Val(TextBox2.Text) * productside * productsidetemp * steamcoil * waterdrawoff

        Dim cr As Double

        If ComboBox11.Text = "Widespread" Then
            cr = modifiedsoilside + modifiedproductside
        ElseIf ComboBox11.Text = "Localized" Then
            If modifiedsoilside > modifiedproductside Then
                cr = modifiedsoilside
            ElseIf modifiedproductside > modifiedsoilside Then
                cr = modifiedproductside
            End If
        End If

    End Sub

    'cr CO2 corrosion

    'cr soil side corrosion
    Public Sub soilside()
        Dim primarysoil As Double

        If Units.ComboBox1.Text = "SI" Then
            If ComboBox12.Text = "Sand" Then
                primarysoil = 0.03
            ElseIf ComboBox12.Text = "Silt" Then
                primarysoil = 0.13
            ElseIf ComboBox12.Text = "Clay" Then
                primarysoil = 0.25
            End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then
            If ComboBox12.Text = "Sand" Then
                primarysoil = 1
            ElseIf ComboBox12.Text = "Silt" Then
                primarysoil = 5
            ElseIf ComboBox12.Text = "Clay" Then
                primarysoil = 10
            End If
        End If

        Dim resistivy As Double

        If ComboBox13.Text = "<500" Then
            resistivy = 1.5
        End If
        If ComboBox13.Text = "500-1000" Then
            resistivy = 1.25
        End If
        If ComboBox13.Text = "1000-2000" Then
            resistivy = 1.0
        End If
        If ComboBox13.Text = "2000-10000" Then
            resistivy = 0.83
        End If
        If ComboBox13.Text = ">10000" Then
            resistivy = 0.6
        End If

        Dim temp As Double

        If ComboBox14.Text = "< 49" Then
            temp = "1.00"
        End If
        If ComboBox14.Text = "49 - 104" Then
            temp = "2.00"
        End If
        If ComboBox14.Text = "> 104" Then
            temp = "1.00"
        End If
        If ComboBox14.Text = "< 120" Then
            temp = "1.00"
        End If
        If ComboBox14.Text = "120 - 220" Then
            temp = "2.00"
        End If
        If ComboBox14.Text = "> 220" Then
            temp = "1.00"
        End If

        Dim cathodic As Double

        If ComboBox15.Text = "No CP on structure (or CP exists but is not regularly tested per NACE RP0169) and CP on an adjacent structure could cause stray current corrosion" Then
            cathodic = 10
        ElseIf ComboBox15.Text = "No cathodic protection" Then
            cathodic = 1.0
        ElseIf ComboBox15.Text = "Cathodic protection exists, but is not tested each year or part of the structure is not in accordance with any NACE RP0169 criteria" Then
            cathodic = 0.8
        ElseIf ComboBox15.Text = "Cathodic protection is tested annually and is in accordance with NACE RP0169 'On' potential criteria over entire structure" Then
            cathodic = 0.4
        ElseIf ComboBox15.Text = "Cathodic protection is tested annually and is in accordance with NACE RP0169 polarized or 'instant-off' potential criteria over entire structure" Then
            cathodic = 0.05
        End If

        Dim coatingtype As Double

        If ComboBox16.Text = "Fusion Bonded Epoxy" Then
            coatingtype = 1.0
        ElseIf ComboBox16.Text = "Liquid Epoxy" Then
            coatingtype = 1.0
        ElseIf ComboBox16.Text = "Asphalt Enamel" Then
            coatingtype = 1.0
        ElseIf ComboBox16.Text = "Asphalt Mastic" Then
            coatingtype = 1.0
        ElseIf ComboBox16.Text = "Coal Tar Enamel" Then
            coatingtype = 1.0
        ElseIf ComboBox16.Text = "Extruded Polyethylene with mastic or rubber" Then
            coatingtype = 1.0
        ElseIf ComboBox16.Text = "Mill Applied PE Tape with mastic" Then
            coatingtype = 1.5
        ElseIf ComboBox16.Text = "Field Applied PE Tape with mastic " Then
            coatingtype = 2.0
        ElseIf ComboBox16.Text = "Three-Layer PE Or PP" Then
            coatingtype = 1.0
        End If

        Dim agecoating As Double

        If ComboBox16.Text = "Liquid Epoxy" And ComboBox17.Text = "Yes" Then
            agecoating = 1.1
        ElseIf ComboBox16.Text = "Asphalt Enamel" And ComboBox17.Text = "Yes" Then
            agecoating = 1.1
        ElseIf ComboBox16.Text = "Asphalt Mastic" And ComboBox17.Text = "Yes" Then
            agecoating = 1.1
        ElseIf ComboBox16.Text = "Coal Tar Enamel" And ComboBox17.Text = "Yes" Then
            agecoating = 1.2
        ElseIf ComboBox16.Text = "Extruded Polyethylene with mastic or rubber" And ComboBox17.Text = "Yes" Then
            agecoating = 1.2
        ElseIf ComboBox16.Text = "Mill Applied PE Tape With mastic" And ComboBox17.Text = "Yes" Then
            agecoating = 1.2
        ElseIf ComboBox16.Text = "Field Applied PE Tape with mastic " And ComboBox17.Text = "Yes" Then
            agecoating = 2.0
        ElseIf ComboBox16.Text = "Three-Layer PE or PP" And ComboBox17.Text = "Yes" Then
            agecoating = 1.2
        End If
        If ComboBox16.Text = "Fusion Bonded Epoxy" And ComboBox17.Text = "No" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Liquid Epoxy" And ComboBox17.Text = "No" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Asphalt Enamel" And ComboBox17.Text = "No" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Asphalt Mastic" And ComboBox17.Text = "No" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Coal Tar Enamel" And ComboBox17.Text = "No" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Extruded Polyethylene With mastic Or rubber" And ComboBox17.Text = "Yes" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Mill Applied PE Tape with mastic" And ComboBox17.Text = "No" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Field Applied PE Tape With mastic " And ComboBox17.Text = "No" Then
            agecoating = 1.0
        ElseIf ComboBox16.Text = "Three-Layer PE or PP" And ComboBox17.Text = "No" Then
            agecoating = 1.0
        End If

        Dim maxrate As Double

        If ComboBox16.Text = "Fusion Bonded Epoxy" And ComboBox18.Text = "Yes" Then
            maxrate = 1.5
        ElseIf ComboBox16.Text = "Liquid Epoxy" And ComboBox18.Text = "Yes" Then
            maxrate = 1.5
        ElseIf ComboBox16.Text = "Asphalt Enamel" And ComboBox18.Text = "Yes" Then
            maxrate = 1.5
        ElseIf ComboBox16.Text = "Asphalt Mastic" And ComboBox18.Text = "Yes" Then
            maxrate = 1.5
        ElseIf ComboBox16.Text = "Coal Tar Enamel" And ComboBox18.Text = "Yes" Then
            maxrate = 2.0
        ElseIf ComboBox16.Text = "Extruded Polyethylene with mastic or rubber" And ComboBox18.Text = "Yes" Then
            maxrate = 3.0
        ElseIf ComboBox16.Text = "Mill Applied PE Tape With mastic" And ComboBox18.Text = "Yes" Then
            maxrate = 3.0
        ElseIf ComboBox16.Text = "Field Applied PE Tape with mastic " And ComboBox18.Text = "Yes" Then
            maxrate = 3.0
        ElseIf ComboBox16.Text = "Three-Layer PE Or PP" And ComboBox18.Text = "Yes" Then
            maxrate = 2.0
        End If

        If ComboBox16.Text = "Fusion Bonded Epoxy" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Liquid Epoxy" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Asphalt Enamel" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Asphalt Mastic" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Coal Tar Enamel" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Extruded Polyethylene with mastic or rubber" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Mill Applied PE Tape with mastic" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Field Applied PE Tape with mastic " And ComboBox18.Text = "No" Then
            maxrate = 1.0
        ElseIf ComboBox16.Text = "Three-Layer PE or PP" And ComboBox18.Text = "No" Then
            maxrate = 1.0
        End If

        Dim coatingmaintenance As Double

        If ComboBox16.Text = "Fusion Bonded Epoxy" And ComboBox19.Text = "Coating Maintenance is Rare or None" Then
            coatingmaintenance = 1.1
        ElseIf ComboBox16.Text = "Liquid Epoxy" And ComboBox19.Text = "Coating Maintenance is Rare or None" Then
            coatingmaintenance = 1.5
        ElseIf ComboBox16.Text = "Asphalt Enamel" And ComboBox19.Text = "Coating Maintenance Is Rare Or None" Then
            coatingmaintenance = 1.5
        ElseIf ComboBox16.Text = "Asphalt Mastic" And ComboBox19.Text = "Coating Maintenance is Rare or None" Then
            coatingmaintenance = 1.5
        ElseIf ComboBox16.Text = "Coal Tar Enamel" And ComboBox19.Text = "Coating Maintenance is Rare or None" Then
            coatingmaintenance = 1.5
        ElseIf ComboBox16.Text = "Extruded Polyethylene with mastic or rubber" And ComboBox19.Text = "Coating Maintenance is Rare or None" Then
            coatingmaintenance = 1.5
        ElseIf ComboBox16.Text = "Mill Applied PE Tape With mastic" And ComboBox19.Text = "Coating Maintenance Is Rare Or None" Then
            coatingmaintenance = 1.5
        ElseIf ComboBox16.Text = "Field Applied PE Tape with mastic " And ComboBox19.Text = "Coating Maintenance is Rare or None" Then
            coatingmaintenance = 1.5
        ElseIf ComboBox16.Text = "Three-Layer PE Or PP" And ComboBox19.Text = "Coating Maintenance Is Rare Or None" Then
            coatingmaintenance = 1.2
        End If
        If ComboBox16.Text = "Fusion Bonded Epoxy" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Liquid Epoxy" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Asphalt Enamel" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Asphalt Mastic" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Coal Tar Enamel" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Extruded Polyethylene with mastic or rubber" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Mill Applied PE Tape With mastic" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Field Applied PE Tape with mastic " And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        ElseIf ComboBox16.Text = "Three-Layer PE Or PP" And ComboBox19.Text = "Else" Then
            coatingmaintenance = 1.0
        End If

        Dim adjfactor As Double

        If ComboBox16.Text = "No Coating" Then
            adjfactor = 1.0
        Else
            adjfactor = coatingtype * agecoating * maxrate * coatingmaintenance
        End If

        Dim cr As Double
        cr = primarysoil * resistivy * temp * cathodic * adjfactor
    End Sub

    'cr cooling water


    'cr acid sour water


    'cr high temperature oxidation 
    'coding panjang ---------------------------------------

    'cr amine corrosion
    'coding panjang ----------------------------------------

    'cr sour water corrosion
    Public Sub sourwater()
        Dim h2s As Double
        Dim nh3 As Double

        h2s = Val(TextBox16.Text)
        nh3 = Val(TextBox17.Text)

        Dim nh4hs As Double

        If h2s < (2 * nh3) Then
            nh4hs = 1.5 * h2s
        End If
        If h2s > (2 * nh3) Then
            nh4hs = 3.0 * h2s
        End If
        If h2s = (2 * nh3) Then
            nh4hs = 1.5 * h2s
        End If

        Dim baselinecr As Double

        If Units.ComboBox1.Text = "SI" Then
            Dim chart1(,) As Double = {{2, 0.08, 0.1, 0.13, 0.2, 0.28},
             {5, 0.15, 0.23, 0.3, 0.38, 0.46},
             {10, 0.51, 0.69, 0.89, 1.09, 1.27},
            {15, 1.14, 1.78, 2.54, 3.81, 5.08}}

            Dim UpNum2(5) As Double
            Dim LowNum2(5) As Double

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
            Dim vel As Double
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

            vel = Val(TextBox18.Text)


            If vel < 3.05 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 2
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(2)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(2)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 4.57
                            Lot = 3.05
                            baselinecr = Inter3b(upt, Lot, vel, x, z)
                        Next
                        Exit For
                    End If
                Next
            ElseIf vel >= 3.05 AndAlso vel <= 4.57 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 2
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(2)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(2)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 4.57
                            Lot = 3.05
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                        Exit For
                    End If
                Next
            ElseIf vel > 6.1 AndAlso vel <= 4.57 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 3
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(2)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(2)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(3)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(3)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 6.1
                            Lot = 4.57
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            ElseIf vel > 6.1 AndAlso vel <= 7.62 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 4
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(3)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(3)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(4)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(4)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 7.62
                            Lot = 6.1
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            ElseIf vel > 7.62 AndAlso vel <= 9.14 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 5
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(4)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(4)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(5)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(5)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 9.14
                            Lot = 7.62
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            Else
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 5
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(4)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(4)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(5)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(5)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 9.14
                            Lot = 7.62
                            baselinecr = Inter3c(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            End If

            If nh4hs < 2 Then
                If vel < 3.05 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(1)
                                xa = LowNum2(1)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(2)
                                xb = LowNum2(2)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 4.57
                                Lot = 3.05
                                baselinecr = Inter3b(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel >= 3.05 AndAlso vel <= 4.57 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(1)
                                xa = LowNum2(1)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(2)
                                xb = LowNum2(2)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 4.57
                                Lot = 3.05
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 4.57 AndAlso vel <= 6.1 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(2)
                                xa = LowNum2(2)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(3)
                                xb = LowNum2(3)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 6.1
                                Lot = 4.57
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 6.1 AndAlso vel <= 7.62 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(3)
                                xa = LowNum2(3)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(4)
                                xb = LowNum2(4)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 7.62
                                Lot = 6.1
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 7.62 AndAlso vel <= 9.14 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(4)
                                xa = LowNum2(4)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(5)
                                xb = LowNum2(5)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 9.14
                                Lot = 7.62
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(4)
                                xa = LowNum2(4)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(5)
                                xb = LowNum2(5)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 9.14
                                Lot = 7.62
                                baselinecr = Inter3c(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                End If
            End If

            If nh4hs > 15 Then
                If vel < 3.05 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(1)
                                xx = LowNum2(1)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(2)
                                xz = LowNum2(2)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 4.57
                                Lot = 3.05
                                baselinecr = Inter3b(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel >= 3.05 AndAlso vel <= 4.57 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(1)
                                xx = LowNum2(1)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(2)
                                xz = LowNum2(2)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 4.57
                                Lot = 3.05
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 4.57 AndAlso vel <= 6.1 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(2)
                                xx = LowNum2(2)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(3)
                                xz = LowNum2(3)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 6.1
                                Lot = 4.57
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 6.1 AndAlso vel <= 7.62 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(3)
                                xx = LowNum2(3)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(4)
                                xz = LowNum2(4)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 7.62
                                Lot = 6.1
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 7.62 AndAlso vel <= 9.14 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(4)
                                xx = LowNum2(4)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(5)
                                xz = LowNum2(5)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 9.14
                                Lot = 7.62
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(4)
                                xx = LowNum2(4)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(5)
                                xz = LowNum2(5)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 9.14
                                Lot = 7.62
                                baselinecr = Inter3c(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                End If
            End If
        End If

        If Units.ComboBox1.Text = "US Customary" Then
            Dim chart1(,) As Double = {{2, 3, 4, 5, 8, 11},
            {5, 6, 9, 12, 15, 18},
            {10, 20, 27, 35, 43, 50},
            {15, 45, 70, 100, 150, 200}}

            Dim UpNum2(5) As Double
            Dim LowNum2(5) As Double

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
            Dim vel As Double
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

            vel = Val(TextBox18.Text)


            If vel < 10 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 2
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(2)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(2)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 15
                            Lot = 10
                            baselinecr = Inter3b(upt, Lot, vel, x, z)
                        Next
                        Exit For
                    End If
                Next
            ElseIf vel >= 10 AndAlso vel <= 15 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 2
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(1)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(1)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(2)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(2)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 15
                            Lot = 10
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                        Exit For
                    End If
                Next
            ElseIf vel > 15 AndAlso vel <= 20 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 3
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(2)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(2)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(3)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(3)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 20
                            Lot = 15
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            ElseIf vel > 20 AndAlso vel <= 25 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 4
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(3)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(3)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(4)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(4)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 25
                            Lot = 20
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            ElseIf vel > 25 AndAlso vel <= 30 Then
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 5
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(4)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(4)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(5)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(5)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 30
                            Lot = 25
                            baselinecr = Inter3a(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            Else
                For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                    If chart1(r, 0) <= nh4hs AndAlso chart1(r + 1, 0) >= nh4hs Then
                        For c As Integer = 0 To 5
                            LowNum2(c) = chart1(r, c)
                            UpNum2(c) = chart1(r + 1, c)

                            up1 = UpNum2(0)
                            up2 = UpNum2(4)

                            Lo1 = LowNum2(0)
                            Lo2 = LowNum2(4)

                            x = Inter(nh4hs, up1, up2, Lo1, Lo2)

                            up3 = UpNum2(0)
                            up4 = UpNum2(5)

                            Lo3 = LowNum2(0)
                            Lo4 = LowNum2(5)
                            z = Inter1(nh4hs, up3, up4, Lo3, Lo4)
                            upt = 30
                            Lot = 25
                            baselinecr = Inter3c(upt, Lot, vel, x, z)
                        Next
                    End If
                Next
            End If

            If nh4hs < 2 Then
                If vel < 10 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(1)
                                xa = LowNum2(1)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(2)
                                xb = LowNum2(2)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 15
                                Lot = 10
                                baselinecr = Inter3b(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel >= 10 AndAlso vel <= 15 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(1)
                                xa = LowNum2(1)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(2)
                                xb = LowNum2(2)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 15
                                Lot = 10
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 15 AndAlso vel <= 20 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(2)
                                xa = LowNum2(2)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(3)
                                xb = LowNum2(3)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 20
                                Lot = 15
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 20 AndAlso vel <= 25 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(3)
                                xa = LowNum2(3)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(4)
                                xb = LowNum2(4)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 25
                                Lot = 20
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 25 AndAlso vel <= 30 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(4)
                                xa = LowNum2(4)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(5)
                                xb = LowNum2(5)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 30
                                Lot = 25
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 5 AndAlso chart1(r + 1, 0) >= 2 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                za = UpNum2(4)
                                xa = LowNum2(4)
                                upta = 5
                                lota = 2
                                x = Inter3da(upta, lota, nh4hs, xa, za)

                                zb = UpNum2(5)
                                xb = LowNum2(5)
                                uptb = 5
                                lotb = 2
                                z = Inter3db(uptb, lotb, nh4hs, xb, zb)
                                upt = 30
                                Lot = 25
                                baselinecr = Inter3c(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                End If
            End If

            If nh4hs > 15 Then
                If vel < 10 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(1)
                                xx = LowNum2(1)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(2)
                                xz = LowNum2(2)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 15
                                Lot = 10
                                baselinecr = Inter3b(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel >= 10 AndAlso vel <= 15 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(1)
                                xx = LowNum2(1)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(2)
                                xz = LowNum2(2)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 15
                                Lot = 10
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 15 AndAlso vel <= 20 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(2)
                                xx = LowNum2(2)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(3)
                                xz = LowNum2(3)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 20
                                Lot = 15
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 20 AndAlso vel <= 25 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(3)
                                xx = LowNum2(3)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(4)
                                xz = LowNum2(4)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 25
                                Lot = 20
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf vel > 25 AndAlso vel <= 30 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(4)
                                xx = LowNum2(4)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(5)
                                xz = LowNum2(5)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 30
                                Lot = 25
                                baselinecr = Inter3a(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= 10 AndAlso chart1(r + 1, 0) >= 15 Then
                            For c As Integer = 0 To 5
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                zx = UpNum2(4)
                                xx = LowNum2(4)
                                uptx = 15
                                lotx = 10
                                x = Inter3ex(uptx, lotx, nh4hs, xx, zx)

                                zz = UpNum2(5)
                                xz = LowNum2(5)
                                uptz = 15
                                lotz = 10
                                z = Inter3ez(uptz, lotz, nh4hs, xz, zz)
                                upt = 30
                                Lot = 25
                                baselinecr = Inter3c(upt, Lot, vel, x, z)
                            Next
                            Exit For
                        End If
                    Next
                End If
            End If
        End If

        Dim h2smole As Double
        Dim p As Double

        h2smole = Val(TextBox19.Text)
        p = Val(TextBox20.Text)

        Dim partialpressure As Double

        partialpressure = h2smole * p

        Dim bcr1 As Double
        Dim bcr2 As Double

        If Units.ComboBox1.Text = "US Customary" Then
            bcr1 = ((baselinecr / 25) * (partialpressure - 50) + baselinecr)
            bcr2 = ((baselinecr / 40) * (partialpressure - 50) + baselinecr)
        ElseIf Units.ComboBox1.Text = "SI" Then
            bcr1 = ((baselinecr / 173) * (partialpressure - 345) + baselinecr)
            bcr2 = ((baselinecr / 276) * (partialpressure - 345) + baselinecr)
        End If

        Dim cr As Double

        If Units.ComboBox1.Text = "US Customary" Then
            If partialpressure < 50 Then
                If bcr1 > 0 Then
                    cr = bcr1
                End If
                If bcr1 < 0 Then
                    cr = "0"
                End If
            End If

            If partialpressure >= 50 Then
                If bcr2 > 0 Then
                    cr = bcr2
                End If
                If bcr2 < 0 Then
                    cr = "0"
                End If
            End If
        ElseIf Units.ComboBox1.Text = "SI" Then
            If partialpressure < 345 Then
                If bcr1 > 0 Then
                    cr = bcr1
                End If
                If bcr1 < 0 Then
                    cr = "0"
                End If
            End If

            If partialpressure >= 345 Then
                If bcr2 > 0 Then
                    cr = bcr2
                End If
                If bcr2 < 0 Then
                    cr = "0"
                End If
            End If
        End If

    End Sub

    'cr hydrofluoric acid (HF) corrosion
    'coding panjang -----------------------------------------

    'cr sulfuric acid (H2SO4) corrosion
    'coding panjang -----------------------------------------

    'cr high temperature H2S/H2 corrosion
    'coding panjang------------------------------------------

    'cr high temperature sulfidic/naphthenic acid corrosion
    'coding panjang------------------------------------------

    'cr hydrochloric acid (HCl) corrosion
    Public Sub hcl()

        Dim cr As Double

        If Units.ComboBox1.Text = "SI" Then
            If ComboBox20.Text = "Carbon Steel" Then
                Dim chart1(,) As Double = {{0.5, 25.37, 25.37, 25.37, 25.37},
             {0.8, 22.86, 25.37, 25.37, 25.37},
             {1.25, 10.16, 25.37, 25.37, 25.37},
             {1.75, 5.08, 17.78, 25.37, 25.37},
             {2.25, 2.54, 7.62, 10.16, 14.22},
             {2.75, 1.52, 3.3, 5.08, 7.11},
             {3.25, 1.02, 1.78, 2.54, 3.56},
             {3.75, 0.76, 1.27, 2.29, 3.18},
             {4.25, 0.51, 1.02, 1.78, 2.54},
             {4.75, 0.25, 0.76, 1.27, 1.78},
             {5.25, 0.18, 0.51, 0.76, 1.02},
             {5.75, 0.1, 0.38, 0.51, 0.76},
             {6.25, 0.08, 0.25, 0.38, 0.51},
             {6.8, 0.05, 0.13, 0.18, 0.25}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim ph As Double
                Dim temp As Double
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



                ph = Val(TextBox5.Text)
                temp = Val(TextBox3.Text)

                If temp < 38 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 38 AndAlso temp <= 52 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 52 AndAlso temp <= 79 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 79
                                Lot = 52
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 79 AndAlso temp <= 93 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If

                If ph < 0.5 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If ph > 6.8 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

            End If

            If ComboBox20.Text = "Type 304, 316, 321, 347 Series Stainless Steels" Then
                Dim chart1(,) As Double = {{0.5, 22.86, 25.37, 25.37, 25.37},
                        {0.8, 12.7, 25.37, 25.37, 25.37},
                        {1.25, 7.62, 12.7, 17.78, 25.37},
                        {1.75, 3.81, 6.6, 10.16, 12.7},
                        {2.25, 2.03, 3.56, 5.08, 6.35},
                        {2.75, 1.27, 1.78, 2.54, 3.05},
                        {3.25, 0.76, 1.02, 1.27, 1.65},
                        {3.75, 0.51, 0.64, 0.76, 0.89},
                        {4.25, 0.25, 0.38, 0.51, 0.64},
                        {4.75, 0.13, 0.18, 0.25, 0.3},
                        {5.25, 0.1, 0.13, 0.15, 0.18},
                        {5.75, 0.08, 0.1, 0.13, 0.15},
                        {6.25, 0.05, 0.08, 0.1, 0.13},
                        {6.8, 0.03, 0.05, 0.08, 0.1}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim ph As Double
                Dim temp As Double
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



                ph = Val(TextBox5.Text)
                temp = Val(TextBox3.Text)

                If temp < 38 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 38 AndAlso temp <= 52 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 52 AndAlso temp <= 79 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 79
                                Lot = 52
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 79 AndAlso temp <= 93 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If

                If ph < 0.5 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If ph > 6.8 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy 825" Then
                Dim chart2(,) As Double = {{0.5, 0.03, 0.08, 1.02, 5.08},
             {0.75, 0.05, 0.13, 2.03, 10.16},
             {1.0, 0.25, 1.78, 7.62, 25.37}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 38 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 38 AndAlso temp <= 52 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 52 AndAlso temp <= 79 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 79
                                Lot = 52
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 79 AndAlso temp <= 93 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If

                If clcon < 0.5 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy 20" Then
                Dim chart2(,) As Double = {{0.5, 0.03, 0.08, 1.02, 5.08},
             {0.75, 0.05, 0.13, 2.03, 10.16},
             {1.0, 0.25, 1.78, 7.62, 25.37}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 38 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 38 AndAlso temp <= 52 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 52 AndAlso temp <= 79 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 79
                                Lot = 52
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 79 AndAlso temp <= 93 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If
                If clcon < 0.5 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy 625" Then
                Dim chart2(,) As Double = {{0.5, 0.03, 0.05, 0.38, 1.91},
             {0.75, 0.03, 0.13, 0.64, 3.18},
             {1.0, 0.05, 1.78, 5.08, 10.16}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 38 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 38 AndAlso temp <= 52 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 52 AndAlso temp <= 79 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 79
                                Lot = 52
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 79 AndAlso temp <= 93 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If
                If clcon < 0.5 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy C-276" Then
                Dim chart2(,) As Double = {{0.5, 0.03, 0.05, 0.2, 0.76},
             {0.75, 0.03, 0.05, 0.38, 1.91},
             {1.0, 0.05, 0.25, 1.52, 7.62}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 38 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 38 AndAlso temp <= 52 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 52
                                Lot = 38
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 52 AndAlso temp <= 79 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 79
                                Lot = 52
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 79 AndAlso temp <= 93 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 93
                                Lot = 79
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If
                If clcon < 0.5 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy B-2" Then
                Dim chart2(,) As Double = {{0.5, 0.03, 0.1, 0.03, 0.1, 0.05, 0.2, 0.1, 0.41},
             {0.75, 0.03, 0.1, 0.03, 0.1, 0.13, 0.51, 0.51, 2.03},
             {1.0, 0.05, 0.2, 0.13, 0.51, 0.25, 1.02, 0.64, 2.54}}

                Dim UpNum2(8) As Double
                Dim LowNum2(8) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If ComboBox21.Text = "No" Then

                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(3)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(3)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(5)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(5)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    End If
                    If clcon < 0.5 Then
                        If temp < 38 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 38 AndAlso temp <= 52 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 52 AndAlso temp <= 79 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(2)
                                        xa = LowNum2(2)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(3)
                                        xb = LowNum2(3)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 79
                                        Lot = 52
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 79 AndAlso temp <= 93 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If clcon > 1.0 Then
                        If temp < 38 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 38 AndAlso temp <= 52 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 52 AndAlso temp <= 79 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(2)
                                        xx = LowNum2(2)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(3)
                                        xz = LowNum2(3)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 79
                                        Lot = 52
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 79 AndAlso temp <= 93 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

                If ComboBox21.Text = "Yes" Then

                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(2)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(2)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(4)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(4)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(2)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(2)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(4)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(4)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(4)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(4)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(6)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(6)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(6)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(6)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(8)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(8)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(6)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(6)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(8)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(8)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    End If
                    If clcon < 0.5 Then
                        If temp < 38 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 38 AndAlso temp <= 52 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 52 AndAlso temp <= 79 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(2)
                                        xa = LowNum2(2)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(3)
                                        xb = LowNum2(3)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 79
                                        Lot = 52
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 79 AndAlso temp <= 93 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If clcon > 1.0 Then
                        If temp < 38 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 38 AndAlso temp <= 52 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 52 AndAlso temp <= 79 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(2)
                                        xx = LowNum2(2)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(3)
                                        xz = LowNum2(3)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 79
                                        Lot = 52
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 79 AndAlso temp <= 93 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

            End If

            If ComboBox20.Text = "Alloy 400" Then
                Dim chart2(,) As Double = {{0.5, 0.03, 0.1, 0.08, 0.3, 0.76, 3.05, 7.62, 25.37},
             {0.75, 0.05, 0.25, 0.13, 0.51, 2.03, 8.13, 20.32, 25.37},
             {1.0, 0.48, 1.02, 0.64, 2.54, 3.81, 15.24, 22.86, 25.37}}

                Dim UpNum2(8) As Double
                Dim LowNum2(8) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If ComboBox21.Text = "No" Then

                    If temp < 38 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 38 AndAlso temp <= 52 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 52
                                    Lot = 38
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 52 AndAlso temp <= 79 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(3)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(3)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(5)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(5)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 79
                                    Lot = 52
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    ElseIf temp > 79 AndAlso temp <= 93 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 93
                                    Lot = 79
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    End If
                    If clcon < 0.5 Then
                        If temp < 38 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 38 AndAlso temp <= 52 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 52 AndAlso temp <= 79 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(2)
                                        xa = LowNum2(2)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(3)
                                        xb = LowNum2(3)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 79
                                        Lot = 52
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 79 AndAlso temp <= 93 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If clcon > 1.0 Then
                        If temp < 38 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 38 AndAlso temp <= 52 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 52 AndAlso temp <= 79 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(2)
                                        xx = LowNum2(2)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(3)
                                        xz = LowNum2(3)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 79
                                        Lot = 52
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 79 AndAlso temp <= 93 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

                If ComboBox21.Text = "Yes" Then

                        If temp < 38 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        up1 = UpNum2(0)
                                        up2 = UpNum2(2)

                                        Lo1 = LowNum2(0)
                                        Lo2 = LowNum2(2)

                                        x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                        up3 = UpNum2(0)
                                        up4 = UpNum2(4)

                                        Lo3 = LowNum2(0)
                                        Lo4 = LowNum2(4)
                                        z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 38 AndAlso temp <= 52 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        up1 = UpNum2(0)
                                        up2 = UpNum2(2)

                                        Lo1 = LowNum2(0)
                                        Lo2 = LowNum2(2)

                                        x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                        up3 = UpNum2(0)
                                        up4 = UpNum2(4)

                                        Lo3 = LowNum2(0)
                                        Lo4 = LowNum2(4)
                                        z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                        upt = 52
                                        Lot = 38
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 52 AndAlso temp <= 79 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                    For c As Integer = 0 To 6
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        up1 = UpNum2(0)
                                        up2 = UpNum2(4)

                                        Lo1 = LowNum2(0)
                                        Lo2 = LowNum2(4)

                                        x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                        up3 = UpNum2(0)
                                        up4 = UpNum2(6)

                                        Lo3 = LowNum2(0)
                                        Lo4 = LowNum2(6)
                                        z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                        upt = 79
                                        Lot = 52
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                End If
                            Next
                        ElseIf temp > 79 AndAlso temp <= 93 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                    For c As Integer = 0 To 8
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        up1 = UpNum2(0)
                                        up2 = UpNum2(6)

                                        Lo1 = LowNum2(0)
                                        Lo2 = LowNum2(6)

                                        x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                        up3 = UpNum2(0)
                                        up4 = UpNum2(8)

                                        Lo3 = LowNum2(0)
                                        Lo4 = LowNum2(8)
                                        z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                    For c As Integer = 0 To 8
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        up1 = UpNum2(0)
                                        up2 = UpNum2(6)

                                        Lo1 = LowNum2(0)
                                        Lo2 = LowNum2(6)

                                        x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                        up3 = UpNum2(0)
                                        up4 = UpNum2(8)

                                        Lo3 = LowNum2(0)
                                        Lo4 = LowNum2(8)
                                        z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                        upt = 93
                                        Lot = 79
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                End If
                            Next
                        End If
                        If clcon < 0.5 Then
                            If temp < 38 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                        For c As Integer = 0 To 2
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            za = UpNum2(1)
                                            xa = LowNum2(1)
                                            upta = 0.75
                                            lota = 0.5
                                            x = Inter3da(upta, lota, clcon, xa, za)

                                            zb = UpNum2(2)
                                            xb = LowNum2(2)
                                            uptb = 0.75
                                            lotb = 0.5
                                            z = Inter3db(uptb, lotb, clcon, xb, zb)
                                            upt = 52
                                            Lot = 38
                                            cr = Inter3b(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            ElseIf temp >= 38 AndAlso temp <= 52 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                        For c As Integer = 0 To 2
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            za = UpNum2(1)
                                            xa = LowNum2(1)
                                            upta = 0.75
                                            lota = 0.5
                                            x = Inter3da(upta, lota, clcon, xa, za)

                                            zb = UpNum2(2)
                                            xb = LowNum2(2)
                                            uptb = 0.75
                                            lotb = 0.5
                                            z = Inter3db(uptb, lotb, clcon, xb, zb)
                                            upt = 52
                                            Lot = 38
                                            cr = Inter3a(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            ElseIf temp > 52 AndAlso temp <= 79 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                        For c As Integer = 0 To 3
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            za = UpNum2(2)
                                            xa = LowNum2(2)
                                            upta = 0.75
                                            lota = 0.5
                                            x = Inter3da(upta, lota, clcon, xa, za)

                                            zb = UpNum2(3)
                                            xb = LowNum2(3)
                                            uptb = 0.75
                                            lotb = 0.5
                                            z = Inter3db(uptb, lotb, clcon, xb, zb)
                                            upt = 79
                                            Lot = 52
                                            cr = Inter3a(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            ElseIf temp > 79 AndAlso temp <= 93 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                        For c As Integer = 0 To 4
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            za = UpNum2(3)
                                            xa = LowNum2(3)
                                            upta = 0.75
                                            lota = 0.5
                                            x = Inter3da(upta, lota, clcon, xa, za)

                                            zb = UpNum2(4)
                                            xb = LowNum2(4)
                                            uptb = 0.75
                                            lotb = 0.5
                                            z = Inter3db(uptb, lotb, clcon, xb, zb)
                                            upt = 93
                                            Lot = 79
                                            cr = Inter3a(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            Else
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                        For c As Integer = 0 To 4
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            za = UpNum2(3)
                                            xa = LowNum2(3)
                                            upta = 0.75
                                            lota = 0.5
                                            x = Inter3da(upta, lota, clcon, xa, za)

                                            zb = UpNum2(4)
                                            xb = LowNum2(4)
                                            uptb = 0.75
                                            lotb = 0.5
                                            z = Inter3db(uptb, lotb, clcon, xb, zb)
                                            upt = 93
                                            Lot = 79
                                            cr = Inter3c(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            End If
                        End If

                        If clcon > 1.0 Then
                            If temp < 38 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                        For c As Integer = 0 To 2
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            zx = UpNum2(1)
                                            xx = LowNum2(1)
                                            uptx = 1.0
                                            lotx = 0.75
                                            x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                            zz = UpNum2(2)
                                            xz = LowNum2(2)
                                            uptz = 1.0
                                            lotz = 0.75
                                            z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                            upt = 52
                                            Lot = 38
                                            cr = Inter3b(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            ElseIf temp >= 38 AndAlso temp <= 52 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                        For c As Integer = 0 To 2
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            zx = UpNum2(1)
                                            xx = LowNum2(1)
                                            uptx = 1.0
                                            lotx = 0.75
                                            x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                            zz = UpNum2(2)
                                            xz = LowNum2(2)
                                            uptz = 1.0
                                            lotz = 0.75
                                            z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                            upt = 52
                                            Lot = 38
                                            cr = Inter3a(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            ElseIf temp > 52 AndAlso temp <= 79 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                        For c As Integer = 0 To 3
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            zx = UpNum2(2)
                                            xx = LowNum2(2)
                                            uptx = 1.0
                                            lotx = 0.75
                                            x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                            zz = UpNum2(3)
                                            xz = LowNum2(3)
                                            uptz = 1.0
                                            lotz = 0.75
                                            z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                            upt = 79
                                            Lot = 52
                                            cr = Inter3a(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            ElseIf temp > 79 AndAlso temp <= 93 Then
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                        For c As Integer = 0 To 4
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            zx = UpNum2(3)
                                            xx = LowNum2(3)
                                            uptx = 1.0
                                            lotx = 0.75
                                            x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                            zz = UpNum2(4)
                                            xz = LowNum2(4)
                                            uptz = 1.0
                                            lotz = 0.75
                                            z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                            upt = 93
                                            Lot = 79
                                            cr = Inter3a(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            Else
                                For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                    If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                        For c As Integer = 0 To 4
                                            LowNum2(c) = chart2(r, c)
                                            UpNum2(c) = chart2(r + 1, c)

                                            zx = UpNum2(3)
                                            xx = LowNum2(3)
                                            uptx = 1.0
                                            lotx = 0.75
                                            x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                            zz = UpNum2(4)
                                            xz = LowNum2(4)
                                            uptz = 1.0
                                            lotz = 0.75
                                            z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                            upt = 93
                                            Lot = 79
                                            cr = Inter3c(upt, Lot, temp, x, z)
                                        Next
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If

        ElseIf Units.ComboBox1.Text = "US Customary" Then

            If ComboBox20.Text = "Carbon Steel" Then
                Dim chart1(,) As Double = {{0.5, 999, 999, 999, 999},
             {0.8, 900, 999, 999, 999},
             {1.25, 400, 999, 999, 999},
             {1.75, 200, 700, 999, 999},
             {2.25, 100, 300, 400, 560},
             {2.75, 60, 130, 200, 280},
             {3.25, 40, 70, 100, 140},
             {3.75, 30, 50, 90, 125},
             {4.25, 20, 40, 70, 100},
             {4.75, 10, 30, 50, 70},
             {5.25, 7, 20, 30, 40},
             {5.75, 4, 15, 20, 30},
             {6.25, 3, 10, 15, 20},
             {6.8, 2, 5, 7, 10}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim ph As Double
                Dim temp As Double
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


                ph = Val(TextBox5.Text)
                temp = Val(TextBox3.Text)

                If temp < 100 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 100 AndAlso temp <= 125 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 125 AndAlso temp <= 175 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 175
                                Lot = 125
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 175 AndAlso temp <= 200 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If






                If ph < 0.5 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If ph > 6.8 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)

                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Type 304, 316, 321, 347 Series Stainless Steels" Then
                Dim chart1(,) As Double = {{0.5, 900, 999, 999, 999},
                 {0.8, 500, 999, 999, 999},
                 {1.25, 300, 500, 700, 999},
                 {1.75, 150, 260, 400, 500},
                 {2.25, 80, 140, 200, 250},
                 {2.75, 50, 70, 100, 120},
                 {3.25, 30, 40, 50, 65},
                 {3.75, 20, 25, 30, 35},
                 {4.25, 10, 15, 20, 25},
                 {4.75, 5, 7, 10, 12},
                 {5.25, 4, 5, 6, 7},
                 {5.75, 3, 4, 5, 6},
                 {6.25, 2, 3, 4, 5},
                 {6.8, 1, 2, 3, 4}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim ph As Double
                Dim temp As Double
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


                ph = Val(TextBox5.Text)
                temp = Val(TextBox3.Text)

                If temp < 100 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 100 AndAlso temp <= 125 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 125 AndAlso temp <= 175 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 175
                                Lot = 125
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 175 AndAlso temp <= 200 Then
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next

                Else
                    For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                        If chart1(r, 0) <= ph AndAlso chart1(r + 1, 0) >= ph Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart1(r, c)
                                UpNum2(c) = chart1(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter(ph, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter1(ph, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If
                If ph < 0.5 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 0.8 AndAlso chart1(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.8
                                    lota = 0.5
                                    x = Inter3da(upta, lota, ph, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.8
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, ph, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If ph > 6.8 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart1.GetUpperBound(0) - 1
                            If chart1(r, 0) <= 6.25 AndAlso chart1(r + 1, 0) >= 6.8 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart1(r, c)
                                    UpNum2(c) = chart1(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 6.8
                                    lotx = 6.25
                                    x = Inter3ex(uptx, lotx, ph, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 6.8
                                    lotz = 6.25
                                    z = Inter3ez(uptz, lotz, ph, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

            End If

            If ComboBox20.Text = "Alloy 825" Then
                Dim chart2(,) As Double = {{0.5, 1, 3, 40, 200},
             {0.75, 2, 5, 80, 400},
             {1.0, 10, 70, 300, 999}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 100 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 100 AndAlso temp <= 125 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 125 AndAlso temp <= 175 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 175
                                Lot = 125
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 175 AndAlso temp <= 200 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If

                If clcon < 0.5 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

            End If

            If ComboBox20.Text = "Alloy 20" Then
                Dim chart2(,) As Double = {{0.5, 1, 3, 40, 200},
             {0.75, 2, 5, 80, 400},
             {1.0, 10, 70, 300, 999}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 100 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 100 AndAlso temp <= 125 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 125 AndAlso temp <= 175 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 175
                                Lot = 125
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 175 AndAlso temp <= 200 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If
                If clcon < 0.5 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy 625" Then
                Dim chart2(,) As Double = {{0.5, 1, 2, 15, 75},
             {0.75, 1, 5, 25, 125},
             {1.0, 2, 70, 200, 400}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 100 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 100 AndAlso temp <= 125 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 125 AndAlso temp <= 175 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 175
                                Lot = 125
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 175 AndAlso temp <= 200 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If
                If clcon < 0.5 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy C-276" Then
                Dim chart2(,) As Double = {{0.5, 1, 2, 8, 30},
             {0.75, 1, 2, 15, 75},
             {1.0, 2, 10, 60, 300}}

                Dim UpNum2(4) As Double
                Dim LowNum2(4) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If temp < 100 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3b(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp >= 100 AndAlso temp <= 125 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 2
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(1)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(1)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(2)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(2)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 125
                                Lot = 100
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                            Exit For
                        End If
                    Next
                ElseIf temp > 125 AndAlso temp <= 175 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 3
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(2)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(2)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(3)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(3)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 175
                                Lot = 125
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                ElseIf temp > 175 AndAlso temp <= 200 Then
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3a(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                Else
                    For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                        If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                            For c As Integer = 0 To 4
                                LowNum2(c) = chart2(r, c)
                                UpNum2(c) = chart2(r + 1, c)

                                up1 = UpNum2(0)
                                up2 = UpNum2(3)

                                Lo1 = LowNum2(0)
                                Lo2 = LowNum2(3)

                                x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                up3 = UpNum2(0)
                                up4 = UpNum2(4)

                                Lo3 = LowNum2(0)
                                Lo4 = LowNum2(4)
                                z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                upt = 200
                                Lot = 175
                                cr = Inter3c(upt, Lot, temp, x, z)
                            Next
                        End If
                    Next
                End If
                If clcon < 0.5 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(1)
                                    xa = LowNum2(1)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(2)
                                    xb = LowNum2(2)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(2)
                                    xa = LowNum2(2)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(3)
                                    xb = LowNum2(3)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    za = UpNum2(3)
                                    xa = LowNum2(3)
                                    upta = 0.75
                                    lota = 0.5
                                    x = Inter3da(upta, lota, clcon, xa, za)

                                    zb = UpNum2(4)
                                    xb = LowNum2(4)
                                    uptb = 0.75
                                    lotb = 0.5
                                    z = Inter3db(uptb, lotb, clcon, xb, zb)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If

                If clcon > 1.0 Then
                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 2
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(1)
                                    xx = LowNum2(1)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(2)
                                    xz = LowNum2(2)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(2)
                                    xx = LowNum2(2)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(3)
                                    xz = LowNum2(3)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    zx = UpNum2(3)
                                    xx = LowNum2(3)
                                    uptx = 1.0
                                    lotx = 0.75
                                    x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                    zz = UpNum2(4)
                                    xz = LowNum2(4)
                                    uptz = 1.0
                                    lotz = 0.75
                                    z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy B-2" Then
                Dim chart2(,) As Double = {{0.5, 1, 4, 1, 4, 2, 8, 4, 16},
             {0.75, 1, 4, 1, 4, 5, 20, 20, 80},
             {1.0, 2, 8, 5, 20, 10, 40, 25, 100}}

                Dim UpNum2(8) As Double
                Dim LowNum2(8) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If ComboBox21.Text = "No" Then

                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(3)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(3)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(5)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(5)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    End If
                    If clcon < 0.5 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(2)
                                        xa = LowNum2(2)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(3)
                                        xb = LowNum2(3)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If clcon > 1.0 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(2)
                                        xx = LowNum2(2)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(3)
                                        xz = LowNum2(3)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

                If ComboBox21.Text = "Yes" Then

                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(2)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(2)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(4)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(4)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(2)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(2)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(4)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(4)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(4)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(4)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(6)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(6)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(6)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(6)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(8)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(8)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(6)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(6)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(8)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(8)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    End If
                    If clcon < 0.5 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(2)
                                        xa = LowNum2(2)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(3)
                                        xb = LowNum2(3)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If clcon > 1.0 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(2)
                                        xx = LowNum2(2)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(3)
                                        xz = LowNum2(3)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            End If

            If ComboBox20.Text = "Alloy 400" Then
                Dim chart2(,) As Double = {{0.5, 1, 4, 3, 12, 30, 120, 300, 999},
         {0.75, 2, 10, 5, 20, 80, 320, 800, 999},
         {1.0, 19, 40, 25, 100, 150, 600, 900, 999}}

                Dim UpNum2(8) As Double
                Dim LowNum2(8) As Double

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
                Dim clcon As Double
                Dim temp As Double
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


                clcon = Val(TextBox4.Text)
                temp = Val(TextBox3.Text)

                If ComboBox21.Text = "No" Then

                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 3
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(1)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(1)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(3)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(3)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 5
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(3)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(3)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(5)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(5)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 7
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(5)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(5)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(7)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(7)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    End If
                    If clcon < 0.5 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(2)
                                        xa = LowNum2(2)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(3)
                                        xb = LowNum2(3)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If clcon > 1.0 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(2)
                                        xx = LowNum2(2)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(3)
                                        xz = LowNum2(3)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

                If ComboBox21.Text = "Yes" Then

                    If temp < 100 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(2)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(2)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(4)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(4)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3b(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp >= 100 AndAlso temp <= 125 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 4
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(2)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(2)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(4)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(4)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 125
                                    Lot = 100
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                                Exit For
                            End If
                        Next
                    ElseIf temp > 125 AndAlso temp <= 175 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 6
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(4)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(4)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(6)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(6)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 175
                                    Lot = 125
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    ElseIf temp > 175 AndAlso temp <= 200 Then
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(6)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(6)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(8)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(8)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3a(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    Else
                        For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                            If chart2(r, 0) <= clcon AndAlso chart2(r + 1, 0) >= clcon Then
                                For c As Integer = 0 To 8
                                    LowNum2(c) = chart2(r, c)
                                    UpNum2(c) = chart2(r + 1, c)

                                    up1 = UpNum2(0)
                                    up2 = UpNum2(6)

                                    Lo1 = LowNum2(0)
                                    Lo2 = LowNum2(6)

                                    x = Inter4(clcon, up1, up2, Lo1, Lo2)

                                    up3 = UpNum2(0)
                                    up4 = UpNum2(8)

                                    Lo3 = LowNum2(0)
                                    Lo4 = LowNum2(8)
                                    z = Inter5(clcon, up3, up4, Lo3, Lo4)
                                    upt = 200
                                    Lot = 175
                                    cr = Inter3c(upt, Lot, temp, x, z)
                                Next
                            End If
                        Next
                    End If
                    If clcon < 0.5 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(1)
                                        xa = LowNum2(1)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(2)
                                        xb = LowNum2(2)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(2)
                                        xa = LowNum2(2)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(3)
                                        xb = LowNum2(3)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 0.5 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        za = UpNum2(3)
                                        xa = LowNum2(3)
                                        upta = 0.75
                                        lota = 0.5
                                        x = Inter3da(upta, lota, clcon, xa, za)

                                        zb = UpNum2(4)
                                        xb = LowNum2(4)
                                        uptb = 0.75
                                        lotb = 0.5
                                        z = Inter3db(uptb, lotb, clcon, xb, zb)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If clcon > 1.0 Then
                        If temp < 100 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3b(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp >= 100 AndAlso temp <= 125 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 2
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(1)
                                        xx = LowNum2(1)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(2)
                                        xz = LowNum2(2)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 125
                                        Lot = 100
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 125 AndAlso temp <= 175 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 3
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(2)
                                        xx = LowNum2(2)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(3)
                                        xz = LowNum2(3)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 175
                                        Lot = 125
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        ElseIf temp > 175 AndAlso temp <= 200 Then
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3a(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        Else
                            For r As Integer = 0 To chart2.GetUpperBound(0) - 1
                                If chart2(r, 0) <= 0.75 AndAlso chart2(r + 1, 0) >= 1.0 Then
                                    For c As Integer = 0 To 4
                                        LowNum2(c) = chart2(r, c)
                                        UpNum2(c) = chart2(r + 1, c)

                                        zx = UpNum2(3)
                                        xx = LowNum2(3)
                                        uptx = 1.0
                                        lotx = 0.75
                                        x = Inter3ex(uptx, lotx, clcon, xx, zx)

                                        zz = UpNum2(4)
                                        xz = LowNum2(4)
                                        uptz = 1.0
                                        lotz = 0.75
                                        z = Inter3ez(uptz, lotz, clcon, xz, zz)
                                        upt = 200
                                        Lot = 175
                                        cr = Inter3c(upt, Lot, temp, x, z)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            End If

        End If
    End Sub


End Class