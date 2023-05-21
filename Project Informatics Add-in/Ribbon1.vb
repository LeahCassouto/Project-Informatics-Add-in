Imports System.Drawing.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Core

Public Class Ribbon1
    Dim qualityf As Form1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'Want to define label1.label only when file open 
        'If Globals.ThisAddIn.Application.ActiveProject IsNot Null Then
        '                 Then
        '
        'End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Label1.Label = Globals.ThisAddIn.Application.ActiveProject.Name
        qualityf = New Form1
        qualityf.Show()


    End Sub
End Class
