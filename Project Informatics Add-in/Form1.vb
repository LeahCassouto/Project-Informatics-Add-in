Imports System.Windows.Forms

Public Class Form1
    Dim Project As String

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'לפתוח קובץ פרןיקט 
        'סינון להיות רק סוג קובץ של פרויקט
        OpenFileDialog1.Filter = "Projects | *.mpp"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
            Project = OpenFileDialog1.FileName
            Button2.Show()


        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class