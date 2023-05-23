Imports Microsoft.Office.Tools.Ribbon

Public Class Form1

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
End Class