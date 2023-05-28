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
    Private Function TheFormIsAlreadyLoaded(ByVal pFormName As String) As Boolean

        TheFormIsAlreadyLoaded = False

        For Each frm As Form In Application.OpenForms
            If frm.Name.Equals(pFormName) Then
                TheFormIsAlreadyLoaded = True
                Exit Function
            End If
        Next

    End Function

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        qualityf = New Form1
        If Not TheFormIsAlreadyLoaded("Form1") Then
            qualityf.Show()
        End If
    End Sub
End Class
