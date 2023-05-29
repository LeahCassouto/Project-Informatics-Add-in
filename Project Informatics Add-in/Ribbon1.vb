Imports System.Drawing.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.MSProject
Imports Microsoft.Office.Tools
Imports Microsoft.Office.Interop


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

        For Each frm As Form In System.Windows.Forms.Application.OpenForms

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

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        SaveFileDialog1.ShowDialog()

        'ImportPercentCompleteToExcel()
    End Sub


    'Private Sub ImportPercentCompleteToExcel()
    '    ' Specify the path to the existing Excel template file
    '    Dim templateFilePath As String = "C:\Users\Leah Cassouto\OneDrive\Documents\SCHOOL\Final Project\Project Informatics Add-in\Template.xltm"

    '    ' Specify the path for the new Excel file
    '    Dim newFilePath As String = "C:\Path\To\Your\NewFile.xlsx"

    '        ' Create an instance of the Excel application
    '        Dim excelApp As New Excel.Application()

    '        ' Open the existing template workbook
    '        Dim templateWorkbook As Excel.Workbook = excelApp.Workbooks.Open(templateFilePath)

    '        ' Create a new workbook based on the template
    '        Dim newWorkbook As Excel.Workbook = templateWorkbook.SaveAs(newFilePath)

    '        ' Get the active worksheet in the new workbook
    '        Dim worksheet As Excel.Worksheet = newWorkbook.ActiveSheet

    '        ' Create an instance of the Project application
    '        Dim projectApp As New Project.Application()

    '        ' Open the Project file
    '        Dim projectFile As Project.Project = projectApp.Projects.Open("C:\Path\To\Your\Project\File.mpp")

    '        ' Get the total percent complete of the project
    '        Dim percentComplete As Integer = projectFile.ProjectSummaryTask.PercentComplete

    '        ' Import the percent complete into the specified cell in Excel
    '        worksheet.Cells(1, 1).Value = "Total Percent Complete"
    '        worksheet.Cells(2, 1).Value = percentComplete

    '        ' Save and close the new Excel workbook
    '        newWorkbook.Save()
    '        newWorkbook.Close()

    '        ' Close the template workbook
    '        templateWorkbook.Close()

    '        ' Close the Project file
    '        projectFile.Close()

    '        ' Quit the Excel and Project applications
    '        excelApp.Quit()
    '        projectApp.Quit()

    '        ' Release the COM objects
    '        ReleaseCOMObject(worksheet)
    '        ReleaseCOMObject(newWorkbook)
    '        ReleaseCOMObject(templateWorkbook)
    '        ReleaseCOMObject(excelApp)
    '        ReleaseCOMObject(projectFile)
    '        ReleaseCOMObject(projectApp)
    '    End Sub

    ' Helper method to release COM objects
    'Private Sub ReleaseCOMObject(ByVal obj As Object)
    '    Try
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
    '        obj = Nothing
    '    Catch ex As Exception
    '        obj = Nothing
    '    Finally
    '        GC.Collect()
    '    End Try
    'End Sub




End Class
