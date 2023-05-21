Imports Microsoft.Office.Tools.Ribbon

Public Class Form1

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Load

        Label2.Text = CStr(Globals.ThisAddIn.Count_tasks(Globals.ThisAddIn.Application.ActiveProject))
        Label4.Text = CStr(Globals.ThisAddIn.Count_tasksover25(Globals.ThisAddIn.Application.ActiveProject))
        Label10.Text = CStr(Globals.ThisAddIn.C_NoSuccessor(Globals.ThisAddIn.Application.ActiveProject))
        Label12.Text = CStr(Globals.ThisAddIn.C_NoPredecessor(Globals.ThisAddIn.Application.ActiveProject))
        Label14.Text = CStr(Globals.ThisAddIn.C_Linkstosummarytasks(Globals.ThisAddIn.Application.ActiveProject))
        Label16.Text = CStr(Globals.ThisAddIn.C_StartFinish(Globals.ThisAddIn.Application.ActiveProject))
        Label18.Text = CStr(Globals.ThisAddIn.C_StartStart(Globals.ThisAddIn.Application.ActiveProject))
        Label20.Text = CStr(Globals.ThisAddIn.C_FinishFinish(Globals.ThisAddIn.Application.ActiveProject))

        Label26.Text = CStr(Globals.ThisAddIn.C_AsLateaspossible(Globals.ThisAddIn.Application.ActiveProject))
        Label28.Text = CStr(Globals.ThisAddIn.C_MustStartOn(Globals.ThisAddIn.Application.ActiveProject))
        Label30.Text = CStr(Globals.ThisAddIn.C_MustFinishOn(Globals.ThisAddIn.Application.ActiveProject))
        Label32.Text = CStr(Globals.ThisAddIn.C_StartNoEarlierThan(Globals.ThisAddIn.Application.ActiveProject))
        Label34.Text = CStr(Globals.ThisAddIn.C_StartNoLaterThan(Globals.ThisAddIn.Application.ActiveProject))
        Label36.Text = CStr(Globals.ThisAddIn.C_FinishNoEarlierThan(Globals.ThisAddIn.Application.ActiveProject))
        Label38.Text = CStr(Globals.ThisAddIn.C_FinishNoLaterThan(Globals.ThisAddIn.Application.ActiveProject))


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
End Class