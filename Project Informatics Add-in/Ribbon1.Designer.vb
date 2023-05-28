Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.FILENAME = Me.Factory.CreateRibbonLabel
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Group2.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Tab1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Items.Add(Me.FILENAME)
        Me.Group2.Label = "להתשמש בקובץ אקסל"
        Me.Group2.Name = "Group2"
        '
        'Button3
        '
        Me.Button3.Image = Global.Project_Informatics_Add_in.My.Resources.Resources.folder
        Me.Button3.Label = "להשתמש בקובץ קיים"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Button2
        '
        Me.Button2.Image = Global.Project_Informatics_Add_in.My.Resources.Resources.file1
        Me.Button2.Label = "לייצר מחדש"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'FILENAME
        '
        Me.FILENAME.Label = "קובץ בשיימוש NONE: "
        Me.FILENAME.Name = "FILENAME"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Label = "Quality Parameters"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.Image = Global.Project_Informatics_Add_in.My.Resources.Resources.speed
        Me.Button1.Label = "Project Quality Parameters"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "Informatics"
        Me.Tab1.Name = "Tab1"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Project.Project"
        Me.Tabs.Add(Me.Tab1)
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FILENAME As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SaveFileDialog1 As Windows.Forms.SaveFileDialog
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
