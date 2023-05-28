Imports System.Runtime.Remoting
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon

Public Class Form1

    Public Limits(20) As Double

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Load

        'Label2.Text = CStr(Globals.ThisAddIn.Count_tasks(Globals.ThisAddIn.Application.ActiveProject)) '1
        'Label4.Text = CStr(Globals.ThisAddIn.Count_tasksover25(Globals.ThisAddIn.Application.ActiveProject)) '2
        ''NOT IN DAYS
        ''IN THE PAST
        'Label10.Text = CStr(Globals.ThisAddIn.C_NoSuccessor(Globals.ThisAddIn.Application.ActiveProject)) '5
        'Label12.Text = CStr(Globals.ThisAddIn.C_NoPredecessor(Globals.ThisAddIn.Application.ActiveProject)) '6
        'Label14.Text = CStr(Globals.ThisAddIn.C_Linkstosummarytasks(Globals.ThisAddIn.Application.ActiveProject)) '7
        'Label16.Text = CStr(Globals.ThisAddIn.C_StartFinish(Globals.ThisAddIn.Application.ActiveProject)) '8
        'Label18.Text = CStr(Globals.ThisAddIn.C_StartStart(Globals.ThisAddIn.Application.ActiveProject)) '9
        'Label20.Text = CStr(Globals.ThisAddIn.C_FinishFinish(Globals.ThisAddIn.Application.ActiveProject)) '10
        'Label22.Text = CStr(Globals.ThisAddIn.C_PositiveLag(Globals.ThisAddIn.Application.ActiveProject)) '11
        'Label24.Text = CStr(Globals.ThisAddIn.C_NegativeLag(Globals.ThisAddIn.Application.ActiveProject)) '12
        'Label26.Text = CStr(Globals.ThisAddIn.C_AsLateaspossible(Globals.ThisAddIn.Application.ActiveProject)) '13
        'Label28.Text = CStr(Globals.ThisAddIn.C_MustStartOn(Globals.ThisAddIn.Application.ActiveProject)) '14
        'Label30.Text = CStr(Globals.ThisAddIn.C_MustFinishOn(Globals.ThisAddIn.Application.ActiveProject)) '15
        'Label32.Text = CStr(Globals.ThisAddIn.C_StartNoEarlierThan(Globals.ThisAddIn.Application.ActiveProject)) '16
        'Label34.Text = CStr(Globals.ThisAddIn.C_StartNoLaterThan(Globals.ThisAddIn.Application.ActiveProject)) '17
        'Label36.Text = CStr(Globals.ThisAddIn.C_FinishNoEarlierThan(Globals.ThisAddIn.Application.ActiveProject)) '18
        'Label38.Text = CStr(Globals.ThisAddIn.C_FinishNoLaterThan(Globals.ThisAddIn.Application.ActiveProject)) '19
        'Label40.Text = CStr(Globals.ThisAddIn.C_LargeTotalSlack(Globals.ThisAddIn.Application.ActiveProject)) '20
        'Label42.Text = CStr(Globals.ThisAddIn.C_NegativeSlack(Globals.ThisAddIn.Application.ActiveProject)) '21
        Globals.ThisAddIn.SetupGrades(Globals.ThisAddIn.Application.ActiveProject)
        Label2.Text = CStr(Globals.ThisAddIn.C1_Tasks) '1
        Label4.Text = CStr(Globals.ThisAddIn.C2_Tasksover25Days) '2
        Label6.Text = CStr(Globals.ThisAddIn.C3_Notindays) ' 3
        Label8.Text = CStr(Globals.ThisAddIn.C4_PlannedinPast) '4
        Label10.Text = CStr(Globals.ThisAddIn.C5_NoSuccessor) '5
        Label12.Text = CStr(Globals.ThisAddIn.C6_NoPredeccessor) '6
        Label14.Text = CStr(Globals.ThisAddIn.C7_LinkstosummaryTasks) '7
        Label16.Text = CStr(Globals.ThisAddIn.C8_StartFinish) '8
        Label18.Text = CStr(Globals.ThisAddIn.C9_StartStart) '9
        Label20.Text = CStr(Globals.ThisAddIn.C10_FinishFinish) '10
        Label22.Text = CStr(Globals.ThisAddIn.C11_PositiveLag) '11
        Label24.Text = CStr(Globals.ThisAddIn.C12_NegativeLag) '12
        Label26.Text = CStr(Globals.ThisAddIn.C13_AsLateAsPossible) '13
        Label28.Text = CStr(Globals.ThisAddIn.C14_MustStartOn) '14
        Label30.Text = CStr(Globals.ThisAddIn.C15_MustFinishOn) '15
        Label32.Text = CStr(Globals.ThisAddIn.C16_StartNoEarlierThan) '16
        Label34.Text = CStr(Globals.ThisAddIn.C17_StarNoLaterThan) '17
        Label36.Text = CStr(Globals.ThisAddIn.C18_FinishNoEarlierThan) '18
        Label38.Text = CStr(Globals.ThisAddIn.C19_FinishNoLaterThan) '19
        Label40.Text = CStr(Globals.ThisAddIn.C20_LargeTotalSlack) '20
        Label42.Text = CStr(Globals.ThisAddIn.C21_NegativeSlack) '21



        If Globals.ThisAddIn.C1_Tasks <> 0 Then
            Labelper1.Text = "100%"
            Labelper2.Text = CStr(Math.Round((Globals.ThisAddIn.C2_Tasksover25Days / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper3.Text = CStr(Math.Round((Globals.ThisAddIn.C3_Notindays / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper4.Text = CStr(Math.Round((Globals.ThisAddIn.C4_PlannedinPast / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            LabelPer5.Text = CStr(Math.Round((Globals.ThisAddIn.C5_NoSuccessor / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            LabelPer6.Text = CStr(Math.Round((Globals.ThisAddIn.C6_NoPredeccessor / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper7.Text = CStr(Math.Round((Globals.ThisAddIn.C7_LinkstosummaryTasks / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            LabelPer8.Text = CStr(Math.Round((Globals.ThisAddIn.C8_StartFinish / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper9.Text = CStr(Math.Round((Globals.ThisAddIn.C9_StartStart / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper10.Text = CStr(Math.Round((Globals.ThisAddIn.C10_FinishFinish / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper11.Text = CStr(Math.Round((Globals.ThisAddIn.C11_PositiveLag / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper12.Text = CStr(Math.Round((Globals.ThisAddIn.C12_NegativeLag / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper13.Text = CStr(Math.Round((Globals.ThisAddIn.C13_AsLateAsPossible / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper14.Text = CStr(Math.Round((Globals.ThisAddIn.C14_MustStartOn / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper15.Text = CStr(Math.Round((Globals.ThisAddIn.C15_MustFinishOn / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper16.Text = CStr(Math.Round((Globals.ThisAddIn.C16_StartNoEarlierThan / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper17.Text = CStr(Math.Round((Globals.ThisAddIn.C17_StarNoLaterThan / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper18.Text = CStr(Math.Round((Globals.ThisAddIn.C18_FinishNoEarlierThan / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper19.Text = CStr(Math.Round((Globals.ThisAddIn.C19_FinishNoLaterThan / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper20.Text = CStr(Math.Round((Globals.ThisAddIn.C20_LargeTotalSlack / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
            Labelper21.Text = CStr(Math.Round((Globals.ThisAddIn.C21_NegativeSlack / Globals.ThisAddIn.C1_Tasks) * 100)) & "%"
        Else
            Labelper1.Text = "0"
            Labelper2.Text = "0"
            Labelper3.Text = "0"
            Labelper4.Text = "0"
            LabelPer5.Text = "0"
            LabelPer6.Text = "0"
            Labelper7.Text = "0"
            LabelPer8.Text = "0"
            Labelper9.Text = "0"
            Labelper10.Text = "0"
            Labelper11.Text = "0"
            Labelper12.Text = "0"
            Labelper13.Text = "0"
            Labelper14.Text = "0"
            Labelper15.Text = "0"
            Labelper16.Text = "0"
            Labelper17.Text = "0"
            Labelper18.Text = "0"
            Labelper19.Text = "0"
            Labelper20.Text = "0"
            Labelper21.Text = "0"
        End If
        'תזכורת LIMITS(1) הוא עבר השני במערך
        Limits(1) = 0.15 '2
        Limits(2) = 0 '3
        Limits(3) = 0   '4
        Limits(4) = 0.1 '5
        Limits(5) = 0.05    '6
        Limits(6) = 0.05 '7
        Limits(7) = 0   '8
        Limits(8) = 0.1
        Limits(9) = 0.2
        Limits(10) = 0.2
        Limits(11) = 0.1
        Limits(12) = 0.05
        Limits(13) = 0.1
        Limits(14) = 0.1
        Limits(15) = 0.1
        Limits(16) = Limits(17) = Limits(18) = Limits(19) = Limits(20) = 0

        TextBox2.Text = FormatPercent(Limits(1), 0)
        TextBox3.Text = FormatPercent(Limits(2), 0)
        TextBox4.Text = FormatPercent(Limits(3), 0)
        TextBox5.Text = FormatPercent(Limits(4), 0)
        TextBox6.Text = FormatPercent(Limits(5), 0)
        TextBox7.Text = FormatPercent(Limits(6), 0)
        TextBox8.Text = FormatPercent(Limits(7), 0)
        TextBox9.Text = FormatPercent(Limits(8), 0)
        TextBox10.Text = FormatPercent(Limits(9), 0)
        TextBox11.Text = FormatPercent(Limits(10), 0)
        TextBox12.Text = FormatPercent(Limits(11), 0)
        TextBox13.Text = FormatPercent(Limits(12), 0)
        TextBox14.Text = FormatPercent(Limits(13), 0)
        TextBox15.Text = FormatPercent(Limits(14), 0)
        TextBox16.Text = FormatPercent(Limits(15), 0)
        TextBox17.Text = FormatPercent(Limits(16), 0)
        TextBox18.Text = FormatPercent(Limits(17), 0)
        TextBox19.Text = FormatPercent(Limits(18), 0)
        TextBox20.Text = FormatPercent(Limits(19), 0)
        TextBox21.Text = FormatPercent(Limits(20), 0)
        CheckBox1.Checked = False
        CheckBox1.Text = "לפתוח"

    End Sub


    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click

    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub Label42_Click(sender As Object, e As EventArgs) Handles Label42.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub

    Private Sub Label40_Click(sender As Object, e As EventArgs) Handles Label40.Click

    End Sub

    Private Sub Labelper4_Click(sender As Object, e As EventArgs) Handles Labelper4.Click

    End Sub

    Private Sub Label50_Click(sender As Object, e As EventArgs) Handles Labelper12.Click

    End Sub

    Private Sub Label43_Click(sender As Object, e As EventArgs) Handles Label43.Click

    End Sub

    Private Sub Label44_Click(sender As Object, e As EventArgs) Handles Label44.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            CheckBox1.Text = "לנעול"
            lockLimit.Text = "פתוח לשינוים"
        Else
            CheckBox1.Text = "לפתוח"
            lockLimit.Text = "נעול לשינוים"

        End If
        TextBox2.Enabled = CheckBox1.Checked
        TextBox3.Enabled = CheckBox1.Checked
        TextBox4.Enabled = CheckBox1.Checked
        TextBox5.Enabled = CheckBox1.Checked
        TextBox6.Enabled = CheckBox1.Checked
        TextBox7.Enabled = CheckBox1.Checked
        TextBox8.Enabled = CheckBox1.Checked
        TextBox9.Enabled = CheckBox1.Checked
        TextBox10.Enabled = CheckBox1.Checked
        TextBox11.Enabled = CheckBox1.Checked
        TextBox12.Enabled = CheckBox1.Checked
        TextBox13.Enabled = CheckBox1.Checked
        TextBox14.Enabled = CheckBox1.Checked
        TextBox15.Enabled = CheckBox1.Checked
        TextBox16.Enabled = CheckBox1.Checked
        TextBox17.Enabled = CheckBox1.Checked
        TextBox18.Enabled = CheckBox1.Checked
        TextBox19.Enabled = CheckBox1.Checked
        TextBox20.Enabled = CheckBox1.Checked
        TextBox21.Enabled = CheckBox1.Checked

    End Sub
    'Private Function TextBox_Check(ByVal t As String)
    '    If t <> String.Empty Then
    '        Dim TypedNumber As String = t
    '        Dim NumberRegex As String = "^[0-9]+\.?[0-9]*$"
    '        If Not System.Text.RegularExpressions.Regex.Match(TypedNumber, NumberRegex).Success Then
    '            t = t.Remove(t.Length - 1, 1)
    '        End If

    '    End If
    'End Function
    Public FixnumberFormatting2 As Boolean
    Private Sub TextBox_Leave(sender As Object, e As EventArgs) Handles TextBox21.Leave, TextBox20.Leave, TextBox19.Leave, TextBox18.Leave, TextBox17.Leave, TextBox16.Leave, TextBox15.Leave, TextBox14.Leave, TextBox13.Leave, TextBox12.Leave, TextBox11.Leave, TextBox10.Leave, TextBox9.Leave, TextBox8.Leave, TextBox7.Leave, TextBox6.Leave, TextBox5.Leave, TextBox4.Leave, TextBox3.Leave, TextBox2.Leave
        Dim i As Integer = CInt(sender.name.ToString.Remove(0, 7)) - 1
        '  If IsNumeric(CInt(Text.ToString.TrimEnd("%"))) Then

        If CDbl(sender.Text.ToString.TrimEnd("%")) > 1 Then
                Limits(i) = (CDbl(sender.Text.ToString.TrimEnd("%")) / 100)
            ElseIf CDbl(sender.Text.ToString.TrimEnd("%")) <= 1 And CDbl(sender.Text.ToString.TrimEnd("%")) >= 0 Then
                Limits(i) = CDbl(sender.Text.ToString.TrimEnd("%"))
            ElseIf CDbl(sender.Text.ToString.TrimEnd("%")) < 0 Then
                MsgBox("מספר לא יכול להיות שלילי")
                sender.text = FormatPercent(Limits(i), 0)
                Exit Sub
            End If
            sender.text = FormatPercent(Limits(i), 0)
            MsgBox("Limits(" & i & ")=" & Limits(i) & "  and " & sender.name & ".text = " & sender.text & "")
        'Else
        '    MsgBox("השדה רק יכול לקבל מספרים")
        '    sender.text = FormatPercent(Limits(i), 0)
        'End If
    End Sub

    Private Sub recalculateDiff(Sender As Object, e As EventArgs) Handles TextBox21.TextChanged, TextBox20.TextChanged, TextBox19.TextChanged, TextBox18.TextChanged, TextBox17.TextChanged, TextBox16.TextChanged, TextBox15.TextChanged, TextBox14.TextChanged, TextBox13.TextChanged, TextBox12.TextChanged, TextBox11.TextChanged, TextBox10.TextChanged, TextBox9.TextChanged, TextBox8.TextChanged, TextBox7.TextChanged, TextBox6.TextChanged, TextBox5.TextChanged, TextBox4.TextChanged, TextBox3.TextChanged, TextBox2.TextChanged, Labelper21.TextChanged, Labelper20.TextChanged, Labelper19.TextChanged, Labelper18.TextChanged, Labelper17.TextChanged, Labelper16.TextChanged, Labelper15.TextChanged, Labelper14.TextChanged, Labelper13.TextChanged, Labelper12.TextChanged, Labelper11.TextChanged, Labelper10.TextChanged, Labelper9.TextChanged, LabelPer8.TextChanged, Labelper7.TextChanged, LabelPer6.TextChanged, LabelPer5.TextChanged, Labelper4.TextChanged, Labelper3.TextChanged, Labelper2.TextChanged
        '
        ' להבין איזה כותרת של הפשרשים צריך לעדכן 
        ' לחשב את הכותרת להיות שווה לגבול מינוס האחוז 

    End Sub

    Private Sub TestCalc(Sender As Object, e As EventArgs) Handles TextBox21.TextChanged, Labelper2.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub
End Class