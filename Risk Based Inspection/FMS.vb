Public Class FMS
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Close()
    End Sub

    Dim leader1 As Double
    Dim leader2a As Double
    Dim leader2b As Double
    Dim leader2c As Double
    Dim leader2d As Double
    Dim leader2e As Double
    Dim leader3 As Double
    Dim leader4 As Double
    Dim leader5 As Double
    Dim leader6 As Double
    Dim leader6a As Double
    Dim leader6b As Double
    Dim totalleader As Double

    Dim safety1 As Double
    Dim safety1a As Double
    Dim safety1b As Double
    Dim safety1c As Double
    Dim safety2 As Double
    Dim safety3a As Double
    Dim safety3b As Double
    Dim safety3c As Double
    Dim safety4 As Double
    Dim safety5 As Double
    Dim safety6 As Double
    Dim safety7a As Double
    Dim safety7b As Double
    Dim safety8a As Double
    Dim safety8b As Double
    Dim safety8c As Double
    Dim safety8d As Double
    Dim safety8e As Double
    Dim safety8f As Double
    Dim safety9 As Double
    Dim safety10 As Double
    Dim totalsafety As Double

    Dim hazard1 As Double
    Dim hazard2 As Double
    Dim hazard2a As Double
    Dim hazard2b As Double
    Dim hazard2c As Double
    Dim hazard2d As Double
    Dim hazard2e As Double
    Dim hazard3a As Double
    Dim hazard3b As Double
    Dim hazard3c As Double
    Dim hazard3d As Double
    Dim hazard3e As Double
    Dim hazard3f As Double
    Dim hazard3g As Double
    Dim hazard4a As Double
    Dim hazard4b As Double
    Dim hazard4c As Double
    Dim hazard4d As Double
    Dim hazard4e As Double
    Dim hazard5 As Double
    Dim hazard5a As Double
    Dim hazard5b As Double
    Dim hazard6 As Double
    Dim hazard7 As Double
    Dim hazard8 As Double
    Dim hazard9 As Double
    Dim totalhazard As Double

    Dim management1a As Double
    Dim management1b As Double
    Dim management2a As Double
    Dim management2b As Double
    Dim management2c As Double
    Dim management2d As Double
    Dim management3 As Double
    Dim management3a As Double
    Dim management3b As Double
    Dim management4a As Double
    Dim management4b As Double
    Dim management4c As Double
    Dim management4d As Double
    Dim management4e As Double
    Dim management4f As Double
    Dim management4g As Double
    Dim management5 As Double
    Dim management6 As Double
    Dim totalmanagement As Double

    Dim operating1a As Double
    Dim operating1b As Double
    Dim operating2a As Double
    Dim operating2b As Double
    Dim operating2c As Double
    Dim operating2d As Double
    Dim operating2e As Double
    Dim operating2f As Double
    Dim operating2g As Double
    Dim operating2h As Double
    Dim operating3a As Double
    Dim operating3b As Double
    Dim operating3c As Double
    Dim operating4 As Double
    Dim operating5 As Double
    Dim operating6a As Double
    Dim operating6b As Double
    Dim operating6c As Double
    Dim operating6d As Double
    Dim operating7a As Double
    Dim operating7b As Double
    Dim operating7c As Double
    Dim operating7d As Double
    Dim totaloperating As Double

    Dim safework1a As Double
    Dim safework1b As Double
    Dim safework1c As Double
    Dim safework1d As Double
    Dim safework1e As Double
    Dim safework1f As Double
    Dim safework1g As Double
    Dim safework1h As Double
    Dim safework1i As Double
    Dim safework1j As Double
    Dim safework2 As Double
    Dim safework3a As Double
    Dim safework3b As Double
    Dim safework3c As Double
    Dim safework3d As Double
    Dim safework3e As Double
    Dim safework4 As Double
    Dim safework5 As Double
    Dim safework6a As Double
    Dim safework6b As Double
    Dim safework6c As Double
    Dim safework6d As Double
    Dim safework7a As Double
    Dim safework7b As Double
    Dim safework8a As Double
    Dim safework8b As Double
    Dim totalsafework As Double

    Dim training1 As Double
    Dim training2 As Double
    Dim training3a As Double
    Dim training3b As Double
    Dim training3c As Double
    Dim training3d As Double
    Dim training3e As Double
    Dim training3f As Double
    Dim training4a As Double
    Dim training4b As Double
    Dim training4c As Double
    Dim training4d As Double
    Dim training5a As Double
    Dim training5b As Double
    Dim training5c As Double
    Dim training6a As Double
    Dim training6b As Double
    Dim training6c As Double
    Dim training6d As Double
    Dim training6e As Double
    Dim training7 As Double
    Dim training7a As Double
    Dim training7b As Double
    Dim training8a As Double
    Dim training8b As Double
    Dim training8c As Double
    Dim training8d As Double
    Dim totaltraining As Double

    Dim mechanicalintegrity1a As Double
    Dim mechanicalintegrity1b As Double
    Dim mechanicalintegrity1c As Double
    Dim mechanicalintegrity1d As Double
    Dim mechanicalintegrity1e As Double
    Dim mechanicalintegrity2 As Double
    Dim mechanicalintegrity2a As Double
    Dim mechanicalintegrity2b As Double
    Dim mechanicalintegrity2c As Double
    Dim mechanicalintegrity3 As Double
    Dim mechanicalintegrity4 As Double
    Dim mechanicalintegrity4a As Double
    Dim mechanicalintegrity4b As Double
    Dim mechanicalintegrity5 As Double
    Dim mechanicalintegrity5a1 As Double
    Dim mechanicalintegrity5a2 As Double
    Dim mechanicalintegrity5b As Double
    Dim mechanicalintegrity5c As Double
    Dim mechanicalintegrity5d As Double
    Dim mechanicalintegrity6a As Double
    Dim mechanicalintegrity6b As Double
    Dim mechanicalintegrity7 As Double
    Dim mechanicalintegrity8a As Double
    Dim mechanicalintegrity8b As Double
    Dim mechanicalintegrity9a As Double
    Dim mechanicalintegrity9b As Double
    Dim mechanicalintegrity10 As Double
    Dim mechanicalintegrity10a As Double
    Dim mechanicalintegrity10b As Double
    Dim mechanicalintegrity11a As Double
    Dim mechanicalintegrity11b As Double
    Dim mechanicalintegrity12 As Double
    Dim mechanicalintegrity13a As Double
    Dim mechanicalintegrity13b As Double
    Dim mechanicalintegrity14 As Double
    Dim mechanicalintegrity15 As Double
    Dim mechanicalintegrity16 As Double
    Dim mechanicalintegrity16a As Double
    Dim mechanicalintegrity16b As Double
    Dim mechanicalintegrity16c As Double
    Dim mechanicalintegrity17a As Double
    Dim mechanicalintegrity17b As Double
    Dim mechanicalintegrity17c As Double
    Dim mechanicalintegrity17d As Double
    Dim mechanicalintegrity17e As Double
    Dim mechanicalintegrity18a As Double
    Dim mechanicalintegrity18b As Double
    Dim mechanicalintegrity18c As Double
    Dim mechanicalintegrity18d As Double
    Dim mechanicalintegrity18e As Double
    Dim mechanicalintegrity19 As Double
    Dim mechanicalintegrity20 As Double
    Dim totalmechanicalintegrity As Double

    Dim prestartup1 As Double
    Dim prestartup2 As Double
    Dim prestartup3 As Double
    Dim prestartup3a As Double
    Dim prestartup3b As Double
    Dim prestartup4a As Double
    Dim prestartup4b As Double
    Dim prestartup4c As Double
    Dim prestartup5 As Double
    Dim totalprestartup As Double

    Dim emergency1 As Double
    Dim emergency2 As Double
    Dim emergency2a As Double
    Dim emergency2b As Double
    Dim emergency3a As Double
    Dim emergency3b As Double
    Dim emergency3c As Double
    Dim emergency3d As Double
    Dim emergency3e As Double
    Dim emergency3f As Double
    Dim emergency3g As Double
    Dim emergency3h As Double
    Dim emergency3i As Double
    Dim emergency4 As Double
    Dim emergency4a As Double
    Dim emergency4b As Double
    Dim emergency4c As Double
    Dim emergency5a As Double
    Dim emergency5b As Double
    Dim emergency6 As Double
    Dim totalemergency As Double

    Dim incident1a As Double
    Dim incident1b As Double
    Dim incident2a As Double
    Dim incident2b As Double
    Dim incident3a As Double
    Dim incident3b As Double
    Dim incident3c As Double
    Dim incident3d As Double
    Dim incident3e As Double
    Dim incident4a As Double
    Dim incident4b As Double
    Dim incident4c As Double
    Dim incident4d As Double
    Dim incident4e As Double
    Dim incident4f As Double
    Dim incident5 As Double
    Dim incident6 As Double
    Dim incident7 As Double
    Dim incident8 As Double
    Dim incident9 As Double
    Dim totalincident As Double

    Dim contractors1a As Double
    Dim contractors1b As Double
    Dim contractors1c As Double
    Dim contractors2a As Double
    Dim contractors2b As Double
    Dim contractors2c As Double
    Dim contractors2d As Double
    Dim contractors3 As Double
    Dim contractors4 As Double
    Dim contractors5 As Double
    Dim totalcontractors As Double

    Dim ms1a As Double
    Dim ms1b As Double
    Dim ms1c As Double
    Dim ms2 As Double
    Dim ms3a As Double
    Dim ms3b As Double
    Dim ms4 As Double
    Dim totalms As Double


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            leader1 = 10
        Else
            leader1 = 0
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            leader2a = 2
        Else
            leader2a = 0
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            leader2b = 2
        Else
            leader2b = 0
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            leader2c = 2
        Else
            leader2c = 0
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            leader2d = 2
        Else
            leader2d = 0
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            leader2e = 2
            TextBox1.Enabled = True
        Else
            leader2e = 0
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            leader3 = 10
        Else
            leader3 = 0
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            leader4 = 15
        Else
            leader4 = 0
        End If
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        leader5 = NumericUpDown1.Value * 10
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            leader6 = 5
        Else
            leader6 = 0
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            leader6a = 5
        Else
            leader6a = 0
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            leader6b = 5
        Else
            leader6b = 0
        End If
    End Sub



    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            safety1 = 5
        Else
            safety1 = 0
        End If
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            safety1a = 2
        Else
            safety1a = 0
        End If
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        If CheckBox14.Checked = True Then
            safety1b = 2
        Else
            safety1b = 0
        End If
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = True Then
            safety1c = 2
        Else
            safety1c = 0
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            safety2 = 10
        Else
            safety2 = 0
        End If
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        If CheckBox17.Checked = True Then
            safety3a = 3
        Else
            safety3a = 0
        End If
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        If CheckBox18.Checked = True Then
            safety3b = 3
        Else
            safety3b = 0
        End If
    End Sub

    Private Sub CheckBox19_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox19.CheckedChanged
        If CheckBox19.Checked = True Then
            safety3c = 3
        Else
            safety3c = 0
        End If
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox20.CheckedChanged
        If CheckBox20.Checked = True Then
            safety4 = 5
        Else
            safety4 = 0
        End If
    End Sub

    Private Sub CheckBox21_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged
        If CheckBox21.Checked = True Then
            safety5 = 10
        Else
            safety5 = 0
        End If
    End Sub

    Private Sub CheckBox22_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox22.CheckedChanged
        If CheckBox22.Checked = True Then
            safety6 = 8
        Else
            safety6 = 0
        End If
    End Sub

    Private Sub CheckBox23_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox23.CheckedChanged
        If CheckBox23.Checked = True Then
            safety7a = 4
        Else
            safety7a = 0
        End If
    End Sub

    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox24.CheckedChanged
        If CheckBox24.Checked = True Then
            safety7b = 4
        Else
            safety7b = 0
        End If
    End Sub

    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        If CheckBox25.Checked = True Then
            safety8a = 1
        Else
            safety8a = 0
        End If
    End Sub

    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox26.CheckedChanged
        If CheckBox26.Checked = True Then
            safety8b = 1
        Else
            safety8b = 0
        End If
    End Sub

    Private Sub CheckBox27_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox27.CheckedChanged
        If CheckBox27.Checked = True Then
            safety8c = 1
        Else
            safety8c = 0
        End If
    End Sub

    Private Sub CheckBox28_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox28.CheckedChanged
        If CheckBox28.Checked = True Then
            safety8d = 1
        Else
            safety8d = 0
        End If
    End Sub

    Private Sub CheckBox29_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox29.CheckedChanged
        If CheckBox29.Checked = True Then
            safety8e = 1
        Else
            safety8e = 0
        End If
    End Sub

    Private Sub CheckBox30_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox30.CheckedChanged
        If CheckBox30.Checked = True Then
            safety8f = 1
        Else
            safety8f = 0
        End If
    End Sub

    Private Sub CheckBox31_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox31.CheckedChanged
        If CheckBox31.Checked = True Then
            safety9 = 5
        Else
            safety9 = 0
        End If
    End Sub

    Private Sub CheckBox32_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox32.CheckedChanged
        If CheckBox32.Checked = True Then
            safety10 = 8
        Else
            safety10 = 0
        End If
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        hazard1 = NumericUpDown2.Value * 10
    End Sub

    Private Sub CheckBox33_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox33.CheckedChanged
        If CheckBox33.Checked = True Then
            hazard2 = 5
        Else
            hazard2 = 0
        End If
    End Sub

    Private Sub CheckBox34_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox34.CheckedChanged
        If CheckBox34.Checked = True Then
            hazard2a = 1
        Else
            hazard2a = 0
        End If
    End Sub

    Private Sub CheckBox35_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox35.CheckedChanged
        If CheckBox35.Checked = True Then
            hazard2b = 1
        Else
            hazard2b = 0
        End If
    End Sub

    Private Sub CheckBox36_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox36.CheckedChanged
        If CheckBox36.Checked = True Then
            hazard2c = 1
        Else
            hazard2c = 0
        End If
    End Sub

    Private Sub CheckBox37_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox37.CheckedChanged
        If CheckBox37.Checked = True Then
            hazard2d = 1
        Else
            hazard2d = 0
        End If
    End Sub

    Private Sub CheckBox38_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox38.CheckedChanged
        If CheckBox38.Checked = True Then
            hazard2e = 1
        Else
            hazard2e = 0
        End If
    End Sub

    Private Sub CheckBox39_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox39.CheckedChanged
        If CheckBox39.Checked = True Then
            hazard3a = 2
        Else
            hazard3a = 0
        End If
    End Sub

    Private Sub CheckBox40_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox40.CheckedChanged
        If CheckBox40.Checked = True Then
            hazard3b = 2
        Else
            hazard3b = 0
        End If
    End Sub

    Private Sub CheckBox41_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox41.CheckedChanged
        If CheckBox41.Checked = True Then
            hazard3c = 2
        Else
            hazard3c = 0
        End If
    End Sub

    Private Sub CheckBox42_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox42.CheckedChanged
        If CheckBox42.Checked = True Then
            hazard3d = 2
        Else
            hazard3d = 0
        End If
    End Sub

    Private Sub CheckBox43_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox43.CheckedChanged
        If CheckBox43.Checked = True Then
            hazard3e = 2
        Else
            hazard3e = 0
        End If
    End Sub

    Private Sub CheckBox44_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox44.CheckedChanged
        If CheckBox44.Checked = True Then
            hazard3f = 2
        Else
            hazard3f = 0
        End If
    End Sub

    Private Sub CheckBox45_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox45.CheckedChanged
        If CheckBox45.Checked = True Then
            hazard3g = 2
        Else
            hazard3g = 0
        End If
    End Sub

    Private Sub CheckBox46_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox46.CheckedChanged
        If CheckBox46.Checked = True Then
            hazard4a = 3
        Else
            hazard4a = 0
        End If
    End Sub

    Private Sub CheckBox47_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox47.CheckedChanged
        If CheckBox47.Checked = True Then
            hazard4b = 3
        Else
            hazard4b = 0
        End If
    End Sub

    Private Sub CheckBox48_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox48.CheckedChanged
        If CheckBox48.Checked = True Then
            hazard4c = 3
        Else
            hazard4c = 0
        End If
    End Sub

    Private Sub CheckBox49_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox49.CheckedChanged
        If CheckBox49.Checked = True Then
            hazard4d = 3
        Else
            hazard4d = 0
        End If
    End Sub

    Private Sub CheckBox50_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox50.CheckedChanged
        If CheckBox50.Checked = True Then
            hazard4e = 3
        Else
            hazard4e = 0
        End If
    End Sub

    Private Sub CheckBox51_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox51.CheckedChanged
        If CheckBox51.Checked = True Then
            hazard5 = 8
        Else
            hazard5 = 0
        End If
    End Sub

    Private Sub CheckBox52_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox52.CheckedChanged
        If CheckBox52.Checked = True Then
            hazard5a = 3
        Else
            hazard5a = 0
        End If
    End Sub

    Private Sub CheckBox53_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox53.CheckedChanged
        If CheckBox53.Checked = True Then
            hazard5b = 3
        Else
            hazard5b = 0
        End If
    End Sub

    Private Sub CheckBox54_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox54.CheckedChanged
        If CheckBox54.Checked = True Then
            hazard6 = 10
        Else
            hazard6 = 0
        End If
    End Sub

    Private Sub CheckBox55_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox55.CheckedChanged
        If CheckBox55.Checked = True Then
            hazard7 = 12
        Else
            hazard7 = 0
        End If
    End Sub

    Private Sub CheckBox56_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox56.CheckedChanged
        If CheckBox56.Checked = True Then
            hazard8 = 10
        Else
            hazard8 = 0
        End If
    End Sub

    Private Sub CheckBox57_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox57.CheckedChanged
        If CheckBox57.Checked = True Then
            hazard9 = 5
        Else
            hazard9 = 0
        End If
    End Sub

    Private Sub CheckBox58_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox58.CheckedChanged
        If CheckBox58.Checked = True Then
            management1a = 9
        Else
            management1a = 0
        End If
    End Sub

    Private Sub CheckBox59_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox59.CheckedChanged
        If CheckBox59.Checked = True Then
            management1b = 5
        Else
            management1b = 0
        End If
    End Sub

    Private Sub CheckBox60_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox60.CheckedChanged
        If CheckBox60.Checked = True Then
            management2a = 4
        Else
            management2a = 0
        End If
    End Sub

    Private Sub CheckBox61_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox61.CheckedChanged
        If CheckBox61.Checked = True Then
            management2b = 4
        Else
            management2b = 0
        End If
    End Sub

    Private Sub CheckBox62_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox62.CheckedChanged
        If CheckBox62.Checked = True Then
            management2c = 4
        Else
            management2c = 0
        End If
    End Sub

    Private Sub CheckBox63_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox63.CheckedChanged
        If CheckBox63.Checked = True Then
            management2d = 4
        Else
            management2d = 0
        End If
    End Sub

    Private Sub CheckBox268_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox268.CheckedChanged
        If CheckBox268.Checked = True Then
            management3 = 5
        Else
            management3 = 0
        End If
    End Sub

    Private Sub CheckBox64_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox64.CheckedChanged
        If CheckBox64.Checked = True Then
            management3a = 4
        Else
            management3a = 0
        End If
    End Sub

    Private Sub CheckBox65_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox65.CheckedChanged
        If CheckBox65.Checked = True Then
            management3b = 5
        Else
            management3b = 0
        End If
    End Sub

    Private Sub CheckBox66_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox66.CheckedChanged
        If CheckBox66.Checked = True Then
            management4a = 3
        Else
            management4a = 0
        End If
    End Sub

    Private Sub CheckBox67_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox67.CheckedChanged
        If CheckBox67.Checked = True Then
            management4b = 3
        Else
            management4b = 0
        End If
    End Sub

    Private Sub CheckBox68_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox68.CheckedChanged
        If CheckBox68.Checked = True Then
            management4c = 3
        Else
            management4c = 0
        End If
    End Sub

    Private Sub CheckBox69_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox69.CheckedChanged
        If CheckBox69.Checked = True Then
            management4d = 3
        Else
            management4d = 0
        End If
    End Sub

    Private Sub CheckBox70_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox70.CheckedChanged
        If CheckBox70.Checked = True Then
            management4e = 3
        Else
            management4e = 0
        End If
    End Sub

    Private Sub CheckBox71_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox71.CheckedChanged
        If CheckBox71.Checked = True Then
            management4f = 3
        Else
            management4f = 0
        End If
    End Sub

    Private Sub CheckBox72_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox72.CheckedChanged
        If CheckBox72.Checked = True Then
            management4g = 3
        Else
            management4g = 0
        End If
    End Sub

    Private Sub CheckBox73_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox73.CheckedChanged
        If CheckBox73.Checked = True Then
            management5 = 10
        Else
            management5 = 0
        End If
    End Sub

    Private Sub CheckBox74_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox74.CheckedChanged
        If CheckBox74.Checked = True Then
            management6 = 5
        Else
            management6 = 0
        End If
    End Sub

    Private Sub CheckBox75_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox75.CheckedChanged
        If CheckBox75.Checked = True Then
            operating1a = 10
        Else
            operating1a = 0
        End If
    End Sub

    Private Sub CheckBox76_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox76.CheckedChanged
        If CheckBox76.Checked = True Then
            operating1b = 5
        Else
            operating1b = 0
        End If
    End Sub

    Private Sub CheckBox77_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox77.CheckedChanged
        If CheckBox77.Checked = True Then
            operating2a = 2
        Else
            operating2a = 0
        End If
    End Sub

    Private Sub CheckBox78_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox78.CheckedChanged
        If CheckBox78.Checked = True Then
            operating2b = 2
        Else
            operating2b = 0
        End If
    End Sub

    Private Sub CheckBox79_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox79.CheckedChanged
        If CheckBox79.Checked = True Then
            operating2c = 2
        Else
            operating2c = 0
        End If
    End Sub

    Private Sub CheckBox80_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox80.CheckedChanged
        If CheckBox80.Checked = True Then
            operating2d = 2
        Else
            operating2d = 0
        End If
    End Sub

    Private Sub CheckBox81_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox81.CheckedChanged
        If CheckBox81.Checked = True Then
            operating2e = 2
        Else
            operating2e = 0
        End If
    End Sub

    Private Sub CheckBox82_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox82.CheckedChanged
        If CheckBox82.Checked = True Then
            operating2f = 2
        Else
            operating2f = 0
        End If
    End Sub

    Private Sub CheckBox83_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox83.CheckedChanged
        If CheckBox83.Checked = True Then
            operating2g = 2
        Else
            operating2g = 0
        End If
    End Sub

    Private Sub CheckBox84_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox84.CheckedChanged
        If CheckBox84.Checked = True Then
            operating2h = 2
        Else
            operating2h = 0
        End If
    End Sub

    Private Sub CheckBox85_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox85.CheckedChanged
        If CheckBox85.Checked = True Then
            operating3a = 3
        Else
            operating3a = 0
        End If
    End Sub

    Private Sub CheckBox86_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox86.CheckedChanged
        If CheckBox86.Checked = True Then
            operating3b = 4
        Else
            operating3b = 0
        End If
    End Sub

    Private Sub CheckBox87_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox87.CheckedChanged
        If CheckBox87.Checked = True Then
            operating3c = 3
        Else
            operating3c = 0
        End If
    End Sub

    Private Sub CheckBox88_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox88.CheckedChanged
        If CheckBox88.Checked = True Then
            operating4 = 10
        Else
            operating4 = 0
        End If
    End Sub

    Private Sub CheckBox89_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox89.CheckedChanged
        If CheckBox89.Checked = True Then
            operating5 = 10
        Else
            operating5 = 0
        End If
    End Sub

    Private Sub CheckBox90_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox90.CheckedChanged
        If CheckBox90.Checked = True Then
            operating6a = 11
            CheckBox90.Enabled = True
            CheckBox91.Enabled = False
            CheckBox92.Enabled = False
            CheckBox93.Enabled = False
        Else
            operating6a = 0
            CheckBox90.Enabled = True
            CheckBox91.Enabled = True
            CheckBox92.Enabled = True
            CheckBox93.Enabled = True
        End If
    End Sub

    Private Sub CheckBox91_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox91.CheckedChanged
        If CheckBox91.Checked = True Then
            operating6b = 6
            CheckBox90.Enabled = False
            CheckBox91.Enabled = True
            CheckBox92.Enabled = False
            CheckBox93.Enabled = False
        Else
            operating6b = 0
            CheckBox90.Enabled = True
            CheckBox91.Enabled = True
            CheckBox92.Enabled = True
            CheckBox93.Enabled = True
        End If
    End Sub

    Private Sub CheckBox92_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox92.CheckedChanged
        If CheckBox92.Checked = True Then
            operating6c = 3
            CheckBox90.Enabled = False
            CheckBox91.Enabled = False
            CheckBox92.Enabled = True
            CheckBox93.Enabled = False
        Else
            operating6c = 0
            CheckBox90.Enabled = True
            CheckBox91.Enabled = True
            CheckBox92.Enabled = True
            CheckBox93.Enabled = True
        End If
    End Sub

    Private Sub CheckBox93_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox93.CheckedChanged
        If CheckBox93.Checked = True Then
            operating6d = 0
            CheckBox90.Enabled = False
            CheckBox91.Enabled = False
            CheckBox92.Enabled = False
            CheckBox93.Enabled = True
        Else
            operating6d = 0
            CheckBox90.Enabled = True
            CheckBox91.Enabled = True
            CheckBox92.Enabled = True
            CheckBox93.Enabled = True
        End If
    End Sub

    Private Sub CheckBox94_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox94.CheckedChanged
        If CheckBox94.Checked = True Then
            operating7a = 8
            CheckBox94.Enabled = True
            CheckBox95.Enabled = False
            CheckBox96.Enabled = False
            CheckBox97.Enabled = False
        Else
            operating7a = 0
            CheckBox94.Enabled = True
            CheckBox95.Enabled = True
            CheckBox96.Enabled = True
            CheckBox97.Enabled = True
        End If
    End Sub

    Private Sub CheckBox95_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox95.CheckedChanged
        If CheckBox95.Checked = True Then
            operating7b = 4
            CheckBox94.Enabled = True
            CheckBox95.Enabled = False
            CheckBox96.Enabled = False
            CheckBox97.Enabled = False
        Else
            operating7b = 0
            CheckBox94.Enabled = True
            CheckBox95.Enabled = True
            CheckBox96.Enabled = True
            CheckBox97.Enabled = True
        End If
    End Sub

    Private Sub CheckBox96_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox96.CheckedChanged
        If CheckBox96.Checked = True Then
            operating7c = 2
            CheckBox94.Enabled = True
            CheckBox95.Enabled = False
            CheckBox96.Enabled = False
            CheckBox97.Enabled = False
        Else
            operating7c = 0
            CheckBox94.Enabled = True
            CheckBox95.Enabled = True
            CheckBox96.Enabled = True
            CheckBox97.Enabled = True
        End If
    End Sub

    Private Sub CheckBox97_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox97.CheckedChanged
        If CheckBox97.Checked = True Then
            operating7d = 0
            CheckBox94.Enabled = True
            CheckBox95.Enabled = False
            CheckBox96.Enabled = False
            CheckBox97.Enabled = False
        Else
            operating7d = 0
            CheckBox94.Enabled = True
            CheckBox95.Enabled = True
            CheckBox96.Enabled = True
            CheckBox97.Enabled = True
        End If
    End Sub

    Private Sub CheckBox98_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox98.CheckedChanged
        If CheckBox98.Checked = True Then
            safework1a = 2
        Else
            safework1a = 0
        End If
    End Sub

    Private Sub CheckBox99_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox99.CheckedChanged
        If CheckBox99.Checked = True Then
            safework1b = 2
        Else
            safework1b = 0
        End If
    End Sub

    Private Sub CheckBox100_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox100.CheckedChanged
        If CheckBox100.Checked = True Then
            safework1c = 2
        Else
            safework1c = 0
        End If
    End Sub

    Private Sub CheckBox101_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox101.CheckedChanged
        If CheckBox101.Checked = True Then
            safework1d = 2
        Else
            safework1d = 0
        End If
    End Sub

    Private Sub CheckBox102_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox102.CheckedChanged
        If CheckBox102.Checked = True Then
            safework1e = 2
        Else
            safework1e = 0
        End If
    End Sub

    Private Sub CheckBox103_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox103.CheckedChanged
        If CheckBox103.Checked = True Then
            safework1f = 2
        Else
            safework1f = 0
        End If
    End Sub

    Private Sub CheckBox104_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox104.CheckedChanged
        If CheckBox104.Checked = True Then
            safework1g = 2
        Else
            safework1g = 0
        End If
    End Sub

    Private Sub CheckBox105_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox105.CheckedChanged
        If CheckBox105.Checked = True Then
            safework1h = 2
        Else
            safework1h = 0
        End If
    End Sub

    Private Sub CheckBox106_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox106.CheckedChanged
        If CheckBox106.Checked = True Then
            safework1i = 2
        Else
            safework1i = 0
        End If
    End Sub

    Private Sub CheckBox107_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox107.CheckedChanged
        If CheckBox107.Checked = True Then
            safework1j = 2
        Else
            safework1j = 0
        End If
    End Sub

    Private Sub CheckBox108_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox108.CheckedChanged
        If CheckBox108.Checked = True Then
            safework2 = 10
        Else
            safework2 = 0
        End If
    End Sub

    Private Sub CheckBox109_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox109.CheckedChanged
        If CheckBox109.Checked = True Then
            safework3a = 1
        Else
            safework3a = 0
        End If
    End Sub

    Private Sub CheckBox110_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox110.CheckedChanged
        If CheckBox110.Checked = True Then
            safework3b = 1
        Else
            safework3b = 0
        End If
    End Sub

    Private Sub CheckBox111_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox111.CheckedChanged
        If CheckBox111.Checked = True Then
            safework3c = 1
        Else
            safework3c = 0
        End If
    End Sub

    Private Sub CheckBox112_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox112.CheckedChanged
        If CheckBox112.Checked = True Then
            safework3d = 1
        Else
            safework3d = 0
        End If
    End Sub

    Private Sub CheckBox113_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox113.CheckedChanged
        If CheckBox113.Checked = True Then
            safework3e = 1
        Else
            safework3e = 0
        End If
    End Sub

    Private Sub CheckBox114_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox114.CheckedChanged
        If CheckBox114.Checked = True Then
            safework4 = 10
        Else
            safework4 = 0
        End If
    End Sub

    Private Sub CheckBox115_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox115.CheckedChanged
        If CheckBox115.Checked = True Then
            safework5 = 10
        Else
            safework5 = 0
        End If
    End Sub

    Private Sub CheckBox116_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox116.CheckedChanged
        If CheckBox116.Checked = True Then
            safework6a = 7
            CheckBox116.Enabled = True
            CheckBox117.Enabled = False
            CheckBox118.Enabled = False
            CheckBox119.Enabled = False
        Else
            safework6a = 0
            CheckBox116.Enabled = True
            CheckBox117.Enabled = True
            CheckBox118.Enabled = True
            CheckBox119.Enabled = True
        End If
    End Sub

    Private Sub CheckBox117_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox117.CheckedChanged
        If CheckBox117.Checked = True Then
            safework6b = 4
            CheckBox116.Enabled = False
            CheckBox117.Enabled = True
            CheckBox118.Enabled = False
            CheckBox119.Enabled = False
        Else
            safework6b = 0
            CheckBox116.Enabled = True
            CheckBox117.Enabled = True
            CheckBox118.Enabled = True
            CheckBox119.Enabled = True
        End If
    End Sub

    Private Sub CheckBox118_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox118.CheckedChanged
        If CheckBox118.Checked = True Then
            safework6c = 2
            CheckBox116.Enabled = False
            CheckBox117.Enabled = False
            CheckBox118.Enabled = True
            CheckBox119.Enabled = False
        Else
            safework6c = 0
            CheckBox116.Enabled = True
            CheckBox117.Enabled = True
            CheckBox118.Enabled = True
            CheckBox119.Enabled = True
        End If
    End Sub

    Private Sub CheckBox119_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox119.CheckedChanged
        If CheckBox119.Checked = True Then
            safework6d = 0
            CheckBox116.Enabled = False
            CheckBox117.Enabled = False
            CheckBox118.Enabled = False
            CheckBox119.Enabled = True
        Else
            safework6d = 0
            CheckBox116.Enabled = True
            CheckBox117.Enabled = True
            CheckBox118.Enabled = True
            CheckBox119.Enabled = True
        End If
    End Sub

    Private Sub CheckBox120_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox120.CheckedChanged
        If CheckBox120.Checked = True Then
            safework7a = 10
        Else
            safework7a = 0
        End If
    End Sub

    Private Sub CheckBox121_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox121.CheckedChanged
        If CheckBox121.Checked = True Then
            safework7b = 5
        Else
            safework7b = 0
        End If
    End Sub

    Private Sub CheckBox122_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox122.CheckedChanged
        If CheckBox122.Checked = True Then
            safework8a = 4
        Else
            safework8a = 0
        End If
    End Sub

    Private Sub CheckBox123_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox123.CheckedChanged
        If CheckBox123.Checked = True Then
            safework8b = 4
        Else
            safework8b = 0
        End If
    End Sub

    Private Sub CheckBox124_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox124.CheckedChanged
        If CheckBox124.Checked = True Then
            training1 = 10
        Else
            training1 = 0
        End If
    End Sub

    Private Sub CheckBox125_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox125.CheckedChanged
        If CheckBox125.Checked = True Then
            training2 = 10
        Else
            training2 = 0
        End If
    End Sub

    Private Sub CheckBox126_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox126.CheckedChanged
        If CheckBox126.Checked = True Then
            training3a = 3
        Else
            training3a = 0
        End If
    End Sub

    Private Sub CheckBox127_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox127.CheckedChanged
        If CheckBox127.Checked = True Then
            training3b = 3
        Else
            training3b = 0
        End If
    End Sub

    Private Sub CheckBox128_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox128.CheckedChanged
        If CheckBox128.Checked = True Then
            training3c = 3
        Else
            training3c = 0
        End If
    End Sub

    Private Sub CheckBox129_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox129.CheckedChanged
        If CheckBox129.Checked = True Then
            training3d = 3
        Else
            training3d = 0
        End If
    End Sub

    Private Sub CheckBox130_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox130.CheckedChanged
        If CheckBox130.Checked = True Then
            training3e = 3
        Else
            training3e = 0
        End If
    End Sub

    Private Sub CheckBox131_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox131.CheckedChanged
        If CheckBox131.Checked = True Then
            training3f = 3
        Else
            training3f = 0
        End If
    End Sub

    Private Sub CheckBox132_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox132.CheckedChanged
        If CheckBox132.Checked = True Then
            training4a = 10
            CheckBox132.Enabled = True
            CheckBox133.Enabled = False
            CheckBox134.Enabled = False
            CheckBox135.Enabled = False
        Else
            training4a = 0
            CheckBox132.Enabled = True
            CheckBox133.Enabled = True
            CheckBox134.Enabled = True
            CheckBox135.Enabled = True
        End If
    End Sub

    Private Sub CheckBox133_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox133.CheckedChanged
        If CheckBox133.Checked = True Then
            training4b = 7
            CheckBox132.Enabled = False
            CheckBox133.Enabled = True
            CheckBox134.Enabled = False
            CheckBox135.Enabled = False
        Else
            training4b = 0
            CheckBox132.Enabled = True
            CheckBox133.Enabled = True
            CheckBox134.Enabled = True
            CheckBox135.Enabled = True
        End If
    End Sub

    Private Sub CheckBox134_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox134.CheckedChanged
        If CheckBox134.Checked = True Then
            training4c = 3
            CheckBox132.Enabled = False
            CheckBox133.Enabled = False
            CheckBox134.Enabled = True
            CheckBox135.Enabled = False
        Else
            training4c = 0
            CheckBox132.Enabled = True
            CheckBox133.Enabled = True
            CheckBox134.Enabled = True
            CheckBox135.Enabled = True
        End If
    End Sub

    Private Sub CheckBox135_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox135.CheckedChanged
        If CheckBox135.Checked = True Then
            training4d = 0
            CheckBox132.Enabled = False
            CheckBox133.Enabled = False
            CheckBox134.Enabled = False
            CheckBox135.Enabled = True
        Else
            training4d = 0
            CheckBox132.Enabled = True
            CheckBox133.Enabled = True
            CheckBox134.Enabled = True
            CheckBox135.Enabled = True
        End If
    End Sub

    Private Sub CheckBox136_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox136.CheckedChanged
        If CheckBox136.Checked = True Then
            training5a = 10
            CheckBox136.Enabled = True
            CheckBox137.Enabled = False
            CheckBox138.Enabled = False
        Else
            training5a = 0
            CheckBox136.Enabled = True
            CheckBox137.Enabled = True
            CheckBox138.Enabled = True
        End If
    End Sub

    Private Sub CheckBox137_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox137.CheckedChanged
        If CheckBox137.Checked = True Then
            training5b = 5
            CheckBox136.Enabled = False
            CheckBox137.Enabled = True
            CheckBox138.Enabled = False
        Else
            training5b = 0
            CheckBox136.Enabled = True
            CheckBox137.Enabled = True
            CheckBox138.Enabled = True
        End If
    End Sub

    Private Sub CheckBox138_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox138.CheckedChanged
        If CheckBox138.Checked = True Then
            training5c = 0
            CheckBox136.Enabled = False
            CheckBox137.Enabled = False
            CheckBox138.Enabled = True
        Else
            training5c = 0
            CheckBox136.Enabled = True
            CheckBox137.Enabled = True
            CheckBox138.Enabled = True
        End If
    End Sub

    Private Sub CheckBox139_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox139.CheckedChanged
        If CheckBox139.Checked = True Then
            training6a = 10
            CheckBox139.Enabled = True
            CheckBox140.Enabled = False
            CheckBox141.Enabled = False
            CheckBox142.Enabled = False
            CheckBox143.Enabled = False
        Else
            training6a = 0
            CheckBox139.Enabled = True
            CheckBox140.Enabled = True
            CheckBox141.Enabled = True
            CheckBox142.Enabled = True
            CheckBox143.Enabled = True
        End If
    End Sub

    Private Sub CheckBox140_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox140.CheckedChanged
        If CheckBox140.Checked = True Then
            training6b = 7
            CheckBox139.Enabled = False
            CheckBox140.Enabled = True
            CheckBox141.Enabled = False
            CheckBox142.Enabled = False
            CheckBox143.Enabled = False
        Else
            training6b = 0
            CheckBox139.Enabled = True
            CheckBox140.Enabled = True
            CheckBox141.Enabled = True
            CheckBox142.Enabled = True
            CheckBox143.Enabled = True
        End If
    End Sub

    Private Sub CheckBox141_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox141.CheckedChanged
        If CheckBox141.Checked = True Then
            training6c = 5
            CheckBox139.Enabled = False
            CheckBox140.Enabled = False
            CheckBox141.Enabled = True
            CheckBox142.Enabled = False
            CheckBox143.Enabled = False
        Else
            training6c = 0
            CheckBox139.Enabled = True
            CheckBox140.Enabled = True
            CheckBox141.Enabled = True
            CheckBox142.Enabled = True
            CheckBox143.Enabled = True
        End If
    End Sub

    Private Sub CheckBox142_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox142.CheckedChanged
        If CheckBox142.Checked = True Then
            training6d = 3
            CheckBox139.Enabled = False
            CheckBox140.Enabled = False
            CheckBox141.Enabled = False
            CheckBox142.Enabled = True
            CheckBox143.Enabled = False
        Else
            training6d = 0
            CheckBox139.Enabled = True
            CheckBox140.Enabled = True
            CheckBox141.Enabled = True
            CheckBox142.Enabled = True
            CheckBox143.Enabled = True
        End If
    End Sub

    Private Sub CheckBox143_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox143.CheckedChanged
        If CheckBox143.Checked = True Then
            training6e = 0
            CheckBox139.Enabled = False
            CheckBox140.Enabled = False
            CheckBox141.Enabled = False
            CheckBox142.Enabled = False
            CheckBox143.Enabled = True
        Else
            training6e = 0
            CheckBox139.Enabled = True
            CheckBox140.Enabled = True
            CheckBox141.Enabled = True
            CheckBox142.Enabled = True
            CheckBox143.Enabled = True
        End If
    End Sub

    Private Sub CheckBox144_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox144.CheckedChanged
        If CheckBox144.Checked = True Then
            training7 = 4
        Else
            training7 = 0
        End If
    End Sub

    Private Sub CheckBox145_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox145.CheckedChanged
        If CheckBox145.Checked = True Then
            training7a = 4
        Else
            training7a = 0
        End If
    End Sub

    Private Sub CheckBox146_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox146.CheckedChanged
        If CheckBox146.Checked = True Then
            training7b = 4
        Else
            training7b = 0
        End If
    End Sub

    Private Sub CheckBox147_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox147.CheckedChanged
        If CheckBox147.Checked = True Then
            training8a = 5
        Else
            training8a = 0
        End If
    End Sub

    Private Sub CheckBox148_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox148.CheckedChanged
        If CheckBox148.Checked = True Then
            training8b = 5
        Else
            training8b = 0
        End If
    End Sub

    Private Sub CheckBox149_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox149.CheckedChanged
        If CheckBox149.Checked = True Then
            training8c = 5
        Else
            training8c = 0
        End If
    End Sub

    Private Sub CheckBox150_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox150.CheckedChanged
        If CheckBox150.Checked = True Then
            training8d = 5
        Else
            training8d = 0
        End If
    End Sub

    Private Sub CheckBox151_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox151.CheckedChanged
        If CheckBox151.Checked = True Then
            mechanicalintegrity1a = 2
        Else
            mechanicalintegrity1a = 0

        End If
    End Sub

    Private Sub CheckBox152_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox152.CheckedChanged
        If CheckBox152.Checked = True Then
            mechanicalintegrity1b = 2
        Else
            mechanicalintegrity1b = 0

        End If
    End Sub

    Private Sub CheckBox153_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox153.CheckedChanged
        If CheckBox153.Checked = True Then
            mechanicalintegrity1c = 2
        Else
            mechanicalintegrity1c = 0

        End If
    End Sub

    Private Sub CheckBox154_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox154.CheckedChanged
        If CheckBox154.Checked = True Then
            mechanicalintegrity1d = 2
        Else
            mechanicalintegrity1d = 0

        End If
    End Sub

    Private Sub CheckBox155_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox155.CheckedChanged
        If CheckBox155.Checked = True Then
            mechanicalintegrity1e = 2
        Else
            mechanicalintegrity1e = 0

        End If
    End Sub

    Private Sub CheckBox156_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox156.CheckedChanged
        If CheckBox156.Checked = True Then
            mechanicalintegrity2 = 2
        Else
            mechanicalintegrity2 = 0

        End If
    End Sub

    Private Sub CheckBox157_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox157.CheckedChanged
        If CheckBox157.Checked = True Then
            mechanicalintegrity2a = 1
        Else
            mechanicalintegrity2a = 0

        End If
    End Sub

    Private Sub CheckBox158_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox158.CheckedChanged
        If CheckBox158.Checked = True Then
            mechanicalintegrity2b = 2
        Else
            mechanicalintegrity2b = 0

        End If
    End Sub

    Private Sub CheckBox159_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox159.CheckedChanged
        If CheckBox159.Checked = True Then
            mechanicalintegrity2c = 2
        Else
            mechanicalintegrity2c = 0

        End If
    End Sub

    Private Sub CheckBox160_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox160.CheckedChanged
        If CheckBox160.Checked = True Then
            mechanicalintegrity3 = 5
        Else
            mechanicalintegrity3 = 0

        End If
    End Sub

    Private Sub CheckBox161_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox161.CheckedChanged
        If CheckBox161.Checked = True Then
            mechanicalintegrity4 = 5
        Else
            mechanicalintegrity4 = 0

        End If
    End Sub

    Private Sub CheckBox162_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox162.CheckedChanged
        If CheckBox162.Checked = True Then
            mechanicalintegrity4a = 1
        Else
            mechanicalintegrity4a = 0

        End If
    End Sub

    Private Sub CheckBox163_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox163.CheckedChanged
        If CheckBox163.Checked = True Then
            mechanicalintegrity4b = 1
        Else
            mechanicalintegrity4b = 0

        End If
    End Sub

    Private Sub CheckBox164_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox164.CheckedChanged
        If CheckBox164.Checked = True Then
            mechanicalintegrity5 = 3
        Else
            mechanicalintegrity5 = 0

        End If
    End Sub

    Private Sub CheckBox165_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox165.CheckedChanged
        If CheckBox165.Checked = True Then
            mechanicalintegrity5a1 = 1
        Else
            mechanicalintegrity5a1 = 0

        End If
    End Sub

    Private Sub CheckBox166_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox166.CheckedChanged
        If CheckBox166.Checked = True Then
            mechanicalintegrity5a2 = 1
        Else
            mechanicalintegrity5a2 = 0

        End If
    End Sub

    Private Sub CheckBox167_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox167.CheckedChanged
        If CheckBox167.Checked = True Then
            mechanicalintegrity5b = 2
        Else
            mechanicalintegrity5b = 0

        End If
    End Sub

    Private Sub CheckBox168_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox168.CheckedChanged
        If CheckBox168.Checked = True Then
            mechanicalintegrity5c = 2
        Else
            mechanicalintegrity5c = 0

        End If
    End Sub

    Private Sub CheckBox169_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox169.CheckedChanged
        If CheckBox169.Checked = True Then
            mechanicalintegrity5d = 2
        Else
            mechanicalintegrity5d = 0

        End If
    End Sub

    Private Sub CheckBox170_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox170.CheckedChanged
        If CheckBox170.Checked = True Then
            mechanicalintegrity6a = 3
        Else
            mechanicalintegrity6a = 0

        End If
    End Sub

    Private Sub CheckBox171_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox171.CheckedChanged
        If CheckBox171.Checked = True Then
            mechanicalintegrity6b = 2
        Else
            mechanicalintegrity6b = 0

        End If
    End Sub

    Private Sub CheckBox172_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox172.CheckedChanged
        If CheckBox172.Checked = True Then
            mechanicalintegrity7 = 5
        Else
            mechanicalintegrity7 = 0

        End If
    End Sub

    Private Sub CheckBox173_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox173.CheckedChanged
        If CheckBox173.Checked = True Then
            mechanicalintegrity8a = 3
        Else
            mechanicalintegrity8a = 0

        End If
    End Sub

    Private Sub CheckBox174_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox174.CheckedChanged
        If CheckBox174.Checked = True Then
            mechanicalintegrity8b = 2
        Else
            mechanicalintegrity8b = 0

        End If
    End Sub

    Private Sub CheckBox175_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox175.CheckedChanged
        If CheckBox175.Checked = True Then
            mechanicalintegrity9a = 3
        Else
            mechanicalintegrity9a = 0

        End If
    End Sub

    Private Sub CheckBox176_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox176.CheckedChanged
        If CheckBox176.Checked = True Then
            mechanicalintegrity9b = 3
        Else
            mechanicalintegrity9b = 0

        End If
    End Sub

    Private Sub CheckBox177_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox177.CheckedChanged
        If CheckBox177.Checked = True Then
            mechanicalintegrity10 = 5
        Else
            mechanicalintegrity10 = 0

        End If
    End Sub

    Private Sub CheckBox178_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox178.CheckedChanged
        If CheckBox178.Checked = True Then
            mechanicalintegrity10a = 1
        Else
            mechanicalintegrity10a = 0

        End If
    End Sub

    Private Sub CheckBox179_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox179.CheckedChanged
        If CheckBox179.Checked = True Then
            mechanicalintegrity10b = 2
        Else
            mechanicalintegrity10b = 0

        End If
    End Sub

    Private Sub CheckBox180_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox180.CheckedChanged
        If CheckBox180.Checked = True Then
            mechanicalintegrity11a = 3
        Else
            mechanicalintegrity11a = 0

        End If
    End Sub

    Private Sub CheckBox269_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox269.CheckedChanged
        If CheckBox269.Checked = True Then
            mechanicalintegrity11b = 2
        Else
            mechanicalintegrity11b = 0

        End If
    End Sub

    Private Sub CheckBox181_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox181.CheckedChanged
        If CheckBox181.Checked = True Then
            mechanicalintegrity12 = 5
        Else
            mechanicalintegrity12 = 0

        End If
    End Sub

    Private Sub CheckBox182_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox182.CheckedChanged
        If CheckBox182.Checked = True Then
            mechanicalintegrity13a = 3
        Else
            mechanicalintegrity13a = 0

        End If
    End Sub

    Private Sub CheckBox183_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox183.CheckedChanged
        If CheckBox183.Checked = True Then
            mechanicalintegrity13b = 2
        Else
            mechanicalintegrity13b = 0

        End If
    End Sub

    Private Sub CheckBox184_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox184.CheckedChanged
        If CheckBox184.Checked = True Then
            mechanicalintegrity14 = 5
        Else
            mechanicalintegrity14 = 0

        End If
    End Sub

    Private Sub CheckBox185_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox185.CheckedChanged
        If CheckBox185.Checked = True Then
            mechanicalintegrity15 = 5
        Else
            mechanicalintegrity15 = 0

        End If
    End Sub

    Private Sub CheckBox186_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox186.CheckedChanged
        If CheckBox186.Checked = True Then
            mechanicalintegrity16 = 3
        Else
            mechanicalintegrity16 = 0

        End If
    End Sub

    Private Sub CheckBox187_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox187.CheckedChanged
        If CheckBox187.Checked = True Then
            mechanicalintegrity16a = 1
        Else
            mechanicalintegrity16a = 0

        End If
    End Sub

    Private Sub CheckBox188_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox188.CheckedChanged
        If CheckBox188.Checked = True Then
            mechanicalintegrity16b = 1
        Else
            mechanicalintegrity16b = 0

        End If
    End Sub

    Private Sub CheckBox189_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox189.CheckedChanged
        If CheckBox189.Checked = True Then
            mechanicalintegrity16c = 1
        Else
            mechanicalintegrity16c = 0
        End If
    End Sub

    Private Sub CheckBox190_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox190.CheckedChanged
        If CheckBox190.Checked = True Then
            mechanicalintegrity17a = 1
        Else
            mechanicalintegrity17a = 0
        End If
    End Sub

    Private Sub CheckBox191_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox191.CheckedChanged
        If CheckBox191.Checked = True Then
            mechanicalintegrity17b = 1
        Else
            mechanicalintegrity17b = 0
        End If
    End Sub

    Private Sub CheckBox192_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox192.CheckedChanged
        If CheckBox192.Checked = True Then
            mechanicalintegrity17c = 1
        Else
            mechanicalintegrity17c = 0
        End If
    End Sub

    Private Sub CheckBox193_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox193.CheckedChanged
        If CheckBox193.Checked = True Then
            mechanicalintegrity17d = 1
        Else
            mechanicalintegrity17d = 0
        End If
    End Sub

    Private Sub CheckBox194_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox194.CheckedChanged
        If CheckBox194.Checked = True Then
            mechanicalintegrity17e = 1
        Else
            mechanicalintegrity17e = 0
        End If
    End Sub

    Private Sub CheckBox195_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox195.CheckedChanged
        If CheckBox195.Checked = True Then
            mechanicalintegrity18a = 1
        Else
            mechanicalintegrity18a = 0
        End If
    End Sub

    Private Sub CheckBox196_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox196.CheckedChanged
        If CheckBox196.Checked = True Then
            mechanicalintegrity18b = 1
        Else
            mechanicalintegrity18b = 0
        End If
    End Sub

    Private Sub CheckBox197_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox197.CheckedChanged
        If CheckBox197.Checked = True Then
            mechanicalintegrity18c = 1
        Else
            mechanicalintegrity18c = 0
        End If
    End Sub

    Private Sub CheckBox198_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox198.CheckedChanged
        If CheckBox198.Checked = True Then
            mechanicalintegrity18d = 1
        Else
            mechanicalintegrity18d = 0
        End If
    End Sub

    Private Sub CheckBox199_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox199.CheckedChanged
        If CheckBox199.Checked = True Then
            mechanicalintegrity18e = 1
        Else
            mechanicalintegrity18e = 0
        End If
    End Sub

    Private Sub CheckBox200_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox200.CheckedChanged
        If CheckBox200.Checked = True Then
            mechanicalintegrity19 = 5
        Else
            mechanicalintegrity19 = 0
        End If
    End Sub

    Private Sub CheckBox201_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox201.CheckedChanged
        If CheckBox201.Checked = True Then
            mechanicalintegrity20 = 5
        Else
            mechanicalintegrity20 = 0
        End If
    End Sub

    Private Sub CheckBox202_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox202.CheckedChanged
        If CheckBox202.Checked = True Then
            prestartup1 = 10
        Else
            prestartup1 = 0
        End If
    End Sub

    Private Sub CheckBox203_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox203.CheckedChanged
        If CheckBox203.Checked = True Then
            prestartup2 = 10
        Else
            prestartup2 = 0
        End If
    End Sub

    Private Sub CheckBox204_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox204.CheckedChanged
        If CheckBox204.Checked = True Then
            prestartup3 = 10
        Else
            prestartup3 = 0
        End If
    End Sub

    Private Sub CheckBox205_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox205.CheckedChanged
        If CheckBox205.Checked = True Then
            prestartup3a = 5
        Else
            prestartup3a = 0
        End If
    End Sub

    Private Sub CheckBox206_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox206.CheckedChanged
        If CheckBox206.Checked = True Then
            prestartup3b = 5
        Else
            prestartup3b = 0
        End If
    End Sub

    Private Sub CheckBox207_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox207.CheckedChanged
        If CheckBox207.Checked = True Then
            prestartup4a = 5
        Else
            prestartup4a = 0
        End If
    End Sub

    Private Sub CheckBox208_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox208.CheckedChanged
        If CheckBox208.Checked = True Then
            prestartup4b = 5
        Else
            prestartup4b = 0
        End If
    End Sub

    Private Sub CheckBox209_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox209.CheckedChanged
        If CheckBox209.Checked = True Then
            prestartup4c = 5
        Else
            prestartup4c = 0
        End If
    End Sub

    Private Sub CheckBox210_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox210.CheckedChanged
        If CheckBox210.Checked = True Then
            prestartup5 = 5
        Else
            prestartup5 = 0
        End If
    End Sub

    Private Sub CheckBox211_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox211.CheckedChanged
        If CheckBox211.Checked = True Then
            emergency1 = 10
        Else
            emergency1 = 0
        End If
    End Sub

    Private Sub CheckBox212_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox212.CheckedChanged
        If CheckBox212.Checked = True Then
            emergency2 = 5
        Else
            emergency2 = 0
        End If
    End Sub

    Private Sub CheckBox213_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox213.CheckedChanged
        If CheckBox213.Checked = True Then
            emergency2a = 2
        Else
            emergency2a = 0
        End If
    End Sub

    Private Sub CheckBox214_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox214.CheckedChanged
        If CheckBox214.Checked = True Then
            emergency2b = 2
        Else
            emergency2b = 0
        End If
    End Sub

    Private Sub CheckBox215_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox215.CheckedChanged
        If CheckBox215.Checked = True Then
            emergency3a = 2
        Else
            emergency3a = 0
        End If
    End Sub

    Private Sub CheckBox216_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox216.CheckedChanged
        If CheckBox216.Checked = True Then
            emergency3b = 2
        Else
            emergency3b = 0
        End If
    End Sub

    Private Sub CheckBox217_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox217.CheckedChanged
        If CheckBox217.Checked = True Then
            emergency3c = 2
        Else
            emergency3c = 0
        End If
    End Sub

    Private Sub CheckBox218_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox218.CheckedChanged
        If CheckBox218.Checked = True Then
            emergency3d = 2
        Else
            emergency3d = 0
        End If
    End Sub

    Private Sub CheckBox219_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox219.CheckedChanged
        If CheckBox219.Checked = True Then
            emergency3e = 2
        Else
            emergency3e = 0
        End If
    End Sub

    Private Sub CheckBox220_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox220.CheckedChanged
        If CheckBox220.Checked = True Then
            emergency3f = 2
        Else
            emergency3f = 0
        End If
    End Sub

    Private Sub CheckBox221_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox221.CheckedChanged
        If CheckBox221.Checked = True Then
            emergency3g = 2
        Else
            emergency3g = 0
        End If
    End Sub

    Private Sub CheckBox222_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox222.CheckedChanged
        If CheckBox222.Checked = True Then
            emergency3h = 2
        Else
            emergency3h = 0
        End If
    End Sub

    Private Sub CheckBox223_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox223.CheckedChanged
        If CheckBox223.Checked = True Then
            emergency3i = 2
        Else
            emergency3i = 0
        End If
    End Sub

    Private Sub CheckBox224_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox224.CheckedChanged
        If CheckBox224.Checked = True Then
            emergency4 = 5
        Else
            emergency4 = 0
        End If
    End Sub

    Private Sub CheckBox225_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox225.CheckedChanged
        If CheckBox225.Checked = True Then
            emergency4a = 2
        Else
            emergency4a = 0
        End If
    End Sub

    Private Sub CheckBox226_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox226.CheckedChanged
        If CheckBox226.Checked = True Then
            emergency4b = 2
        Else
            emergency4b = 0
        End If
    End Sub

    Private Sub CheckBox227_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox227.CheckedChanged
        If CheckBox227.Checked = True Then
            emergency4c = 2
        Else
            emergency4c = 0
        End If
    End Sub

    Private Sub CheckBox228_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox228.CheckedChanged
        If CheckBox228.Checked = True Then
            emergency5a = 5
        Else
            emergency5a = 0
        End If
    End Sub

    Private Sub CheckBox229_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox229.CheckedChanged
        If CheckBox229.Checked = True Then
            emergency5b = 2
        Else
            emergency5b = 0
        End If
    End Sub

    Private Sub CheckBox230_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox230.CheckedChanged
        If CheckBox230.Checked = True Then
            emergency6 = 10
        Else
            emergency6 = 0
        End If
    End Sub

    Private Sub CheckBox231_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox231.CheckedChanged
        If CheckBox231.Checked = True Then
            incident1a = 10
        Else
            incident1a = 0
        End If
    End Sub

    Private Sub CheckBox232_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox232.CheckedChanged
        If CheckBox232.Checked = True Then
            incident1b = 5
        Else
            incident1b = 0
        End If
    End Sub

    Private Sub CheckBox233_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox233.CheckedChanged
        If CheckBox233.Checked = True Then
            incident2a = 3
        Else
            incident2a = 0
        End If
    End Sub

    Private Sub CheckBox234_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox234.CheckedChanged
        If CheckBox234.Checked = True Then
            incident2b = 3
        Else
            incident2b = 0
        End If
    End Sub

    Private Sub CheckBox235_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox235.CheckedChanged
        If CheckBox235.Checked = True Then
            incident3a = 2
        Else
            incident3a = 0
        End If
    End Sub

    Private Sub CheckBox236_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox236.CheckedChanged
        If CheckBox236.Checked = True Then
            incident3b = 2
        Else
            incident3b = 0
        End If
    End Sub

    Private Sub CheckBox237_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox237.CheckedChanged
        If CheckBox237.Checked = True Then
            incident3c = 2
        Else
            incident3c = 0
        End If
    End Sub

    Private Sub CheckBox238_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox238.CheckedChanged
        If CheckBox238.Checked = True Then
            incident3d = 2
        Else
            incident3d = 0
        End If
    End Sub

    Private Sub CheckBox239_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox239.CheckedChanged
        If CheckBox239.Checked = True Then
            incident3e = 2
        Else
            incident3e = 0
        End If
    End Sub

    Private Sub CheckBox240_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox240.CheckedChanged
        If CheckBox240.Checked = True Then
            incident4a = 2
        Else
            incident4a = 0
        End If
    End Sub

    Private Sub CheckBox241_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox241.CheckedChanged
        If CheckBox241.Checked = True Then
            incident4b = 2
        Else
            incident4b = 0
        End If
    End Sub

    Private Sub CheckBox242_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox242.CheckedChanged
        If CheckBox242.Checked = True Then
            incident4c = 2
        Else
            incident4c = 0
        End If
    End Sub

    Private Sub CheckBox243_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox243.CheckedChanged
        If CheckBox243.Checked = True Then
            incident4d = 2
        Else
            incident4d = 0
        End If
    End Sub

    Private Sub CheckBox244_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox244.CheckedChanged
        If CheckBox244.Checked = True Then
            incident4e = 2
        Else
            incident4e = 0
        End If
    End Sub

    Private Sub CheckBox245_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox245.CheckedChanged
        If CheckBox245.Checked = True Then
            incident4f = 2
        Else
            incident4f = 0
        End If
    End Sub

    Private Sub CheckBox246_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox246.CheckedChanged
        If CheckBox246.Checked = True Then
            incident5 = 5
        Else
            incident5 = 0
        End If
    End Sub

    Private Sub CheckBox247_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox247.CheckedChanged
        If CheckBox247.Checked = True Then
            incident6 = 10
        Else
            incident6 = 0
        End If
    End Sub

    Private Sub CheckBox248_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox248.CheckedChanged
        If CheckBox248.Checked = True Then
            incident7 = 5
        Else
            incident7 = 0
        End If
    End Sub

    Private Sub CheckBox249_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox249.CheckedChanged
        If CheckBox249.Checked = True Then
            incident8 = 6
        Else
            incident8 = 0
        End If
    End Sub

    Private Sub CheckBox250_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox250.CheckedChanged
        If CheckBox250.Checked = True Then
            incident9 = 6
        Else
            incident9 = 0
        End If
    End Sub

    Private Sub CheckBox251_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox251.CheckedChanged
        If CheckBox251.Checked = True Then
            contractors1a = 3
        Else
            contractors1a = 0
        End If
    End Sub

    Private Sub CheckBox252_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox252.CheckedChanged
        If CheckBox252.Checked = True Then
            contractors1b = 3
        Else
            contractors1b = 0
        End If
    End Sub

    Private Sub CheckBox253_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox253.CheckedChanged
        If CheckBox253.Checked = True Then
            contractors1c = 3
        Else
            contractors1c = 0
        End If
    End Sub

    Private Sub CheckBox254_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox254.CheckedChanged
        If CheckBox254.Checked = True Then
            contractors2a = 2
        Else
            contractors2a = 0
        End If
    End Sub

    Private Sub CheckBox255_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox255.CheckedChanged
        If CheckBox255.Checked = True Then
            contractors2b = 2
        Else
            contractors2b = 0
        End If
    End Sub

    Private Sub CheckBox256_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox256.CheckedChanged
        If CheckBox256.Checked = True Then
            contractors2c = 2
        Else
            contractors2c = 0
        End If
    End Sub

    Private Sub CheckBox257_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox257.CheckedChanged
        If CheckBox257.Checked = True Then
            contractors2d = 2
        Else
            contractors2d = 0
        End If
    End Sub

    Private Sub CheckBox258_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox258.CheckedChanged
        If CheckBox258.Checked = True Then
            contractors3 = 9
        Else
            contractors3 = 0
        End If
    End Sub

    Private Sub CheckBox259_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox259.CheckedChanged
        If CheckBox259.Checked = True Then
            contractors4 = 9
        Else
            contractors4 = 0
        End If
    End Sub

    Private Sub CheckBox260_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox260.CheckedChanged
        If CheckBox260.Checked = True Then
            contractors5 = 10
        Else
            contractors5 = 0
        End If
    End Sub

    Private Sub CheckBox261_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox261.CheckedChanged
        If CheckBox261.Checked = True Then
            ms1a = 10
            CheckBox261.Enabled = True
            CheckBox262.Enabled = False
            CheckBox263.Enabled = False
        Else
            ms1a = 0
            CheckBox261.Enabled = True
            CheckBox262.Enabled = True
            CheckBox263.Enabled = True
        End If
    End Sub

    Private Sub CheckBox262_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox262.CheckedChanged
        If CheckBox262.Checked = True Then
            ms1b = 7
            CheckBox261.Enabled = False
            CheckBox262.Enabled = True
            CheckBox263.Enabled = False
        Else
            ms1b = 0
            CheckBox261.Enabled = True
            CheckBox262.Enabled = True
            CheckBox263.Enabled = True
        End If
    End Sub

    Private Sub CheckBox263_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox263.CheckedChanged
        If CheckBox263.Checked = True Then
            ms1c = 0
            CheckBox261.Enabled = False
            CheckBox262.Enabled = False
            CheckBox263.Enabled = True
        Else
            ms1c = 0
            CheckBox261.Enabled = True
            CheckBox262.Enabled = True
            CheckBox263.Enabled = True
        End If
    End Sub

    Private Sub CheckBox264_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox264.CheckedChanged
        If CheckBox264.Checked = True Then
            ms2 = 10
        Else
            ms2 = 0
        End If
    End Sub

    Private Sub CheckBox265_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox265.CheckedChanged
        If CheckBox265.Checked = True Then
            ms3a = 5
        Else
            ms3a = 0
        End If
    End Sub

    Private Sub CheckBox266_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox266.CheckedChanged
        If CheckBox266.Checked = True Then
            ms3b = 5
        Else
            ms3b = 0
        End If
    End Sub

    Private Sub CheckBox267_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox267.CheckedChanged
        If CheckBox267.Checked = True Then
            ms4 = 10
        Else
            ms4 = 0
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        totalleader = leader1 + leader2a + leader2b + leader2c + leader2d + leader2e + leader3 + leader4 + leader5 + leader6 + leader6a + leader6b
        totalsafety = safety1 + safety1a + safety1b + safety1c + safety2 + safety3a + safety3b + safety3c + safety4 + safety5 + safety6 + safety7a + safety7b + safety8a + safety8b + safety8c + safety8d + safety8e + safety8f + safety9 + safety10
        totalhazard = hazard1 + hazard2 + hazard2a + hazard2b + hazard2c + hazard2d + hazard2e + hazard3a + hazard3b + hazard3c + hazard3d + hazard3e + hazard3f + hazard3g + hazard4a + hazard4b + hazard4c + hazard4d + hazard4e + hazard5 + hazard5a + hazard5b + hazard6 + hazard7 + hazard8 + hazard9
        totalmanagement = management1a + management1b + management2a + management2b + management2c + management2d + management3 + management3a + management3b + management4a + management4b + management4c + management4d + management4e + management4f + management4g + management5 + management6
        totaloperating = operating1a + operating1b + operating2a + operating2b + operating2c + operating2d + operating2e + operating2f + operating2g + operating2h + operating3a + operating3b + operating3c + operating4 + operating5 + operating6a + operating6b + operating6c + operating6d + operating7a + operating7b + operating7c + operating7d
        totalsafework = safework1a + safework1b + safework1c + safework1d + safework1e + safework1f + safework1g + safework1h + safework1i + safework1j + safework2 + safework3a + safework3b + safework3c + safework3d + safework3e + safework4 + safework5 + safework6a + safework6b + safework6c + safework6d + safework7a + safework7b + safework8a + safework8b
        totaltraining = training1 + training2 + training3a + training3b + training3c + training3d + training3e + training3f + training4a + training4b + training4c + training4d + training5a + training5b + training5c + training6a + training6b + training6c + training6d + training6e + training7 + training7a + training7b + training8a + training8b + training8c + training8d
        totalmechanicalintegrity = mechanicalintegrity1a + mechanicalintegrity1b + mechanicalintegrity1c + mechanicalintegrity1d + mechanicalintegrity1e + mechanicalintegrity2 + mechanicalintegrity2a + mechanicalintegrity2b + mechanicalintegrity2c + mechanicalintegrity3 + mechanicalintegrity4 + mechanicalintegrity4a + mechanicalintegrity4b + mechanicalintegrity5 + mechanicalintegrity5a1 + mechanicalintegrity5a2 + mechanicalintegrity5b + mechanicalintegrity5c + mechanicalintegrity5d + mechanicalintegrity6a + mechanicalintegrity6b + mechanicalintegrity7 + mechanicalintegrity8a + mechanicalintegrity8b + mechanicalintegrity9a + mechanicalintegrity9b + mechanicalintegrity10 + mechanicalintegrity10a + mechanicalintegrity10b + mechanicalintegrity11a + mechanicalintegrity11b + mechanicalintegrity12 + mechanicalintegrity13a + mechanicalintegrity13b + mechanicalintegrity14 + mechanicalintegrity15 + mechanicalintegrity16 + mechanicalintegrity16a + mechanicalintegrity16b + mechanicalintegrity16c + mechanicalintegrity17a + mechanicalintegrity17b + mechanicalintegrity17c + mechanicalintegrity17d + mechanicalintegrity17e + mechanicalintegrity18a + mechanicalintegrity18b + mechanicalintegrity18c + mechanicalintegrity18d + mechanicalintegrity18e + mechanicalintegrity19 + mechanicalintegrity20
        totalprestartup = prestartup1 + prestartup2 + prestartup3 + prestartup3a + prestartup3b + prestartup4a + prestartup4b + prestartup4c + prestartup5
        totalemergency = emergency1 + emergency2 + emergency2a + emergency2b + emergency3a + emergency3b + emergency3c + emergency3d + emergency3e + emergency3f + emergency3g + emergency3h + emergency3i + emergency4 + emergency4a + emergency4b + emergency4c + emergency5a + emergency5b + emergency6
        totalincident = incident1a + incident1b + incident2a + incident2b + incident3a + incident3b + incident3c + incident3d + incident3e + incident4a + incident4b + incident4c + incident4d + incident4e + incident4f + incident5 + incident6 + incident7 + incident8 + incident9
        totalcontractors = contractors1a + contractors1b + contractors1c + contractors2a + contractors2b + contractors2c + contractors2d + contractors3 + contractors4 + contractors5
        totalms = ms1a + ms1b + ms1c + ms2 + ms3a + ms3b + ms4 

        Dim total As Double

        total = totalleader + totalsafety + totalhazard + totalmanagement + totaloperating + totalsafework + totaltraining + totalmechanicalintegrity + totalprestartup + totalemergency + totalincident + totalcontractors + totalms
        TextBox2.Text = total
    End Sub
End Class