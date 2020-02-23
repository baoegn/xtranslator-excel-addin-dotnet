Partial Class ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ribbon))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.gpRibbon = Me.Factory.CreateRibbonGroup
        Me.ddFromLang = Me.Factory.CreateRibbonDropDown
        Me.ddToLang = Me.Factory.CreateRibbonDropDown
        Me.ButtonGroup1 = Me.Factory.CreateRibbonButtonGroup
        Me.btnSwitch = Me.Factory.CreateRibbonButton
        Me.btnSpace = Me.Factory.CreateRibbonButton
        Me.btnAbout = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.btnUndo = Me.Factory.CreateRibbonButton
        Me.btnTranslate = Me.Factory.CreateRibbonButton
        Me.lblProg = Me.Factory.CreateRibbonLabel
        Me.Tab1.SuspendLayout()
        Me.gpRibbon.SuspendLayout()
        Me.ButtonGroup1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.gpRibbon)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'gpRibbon
        '
        Me.gpRibbon.Items.Add(Me.ddFromLang)
        Me.gpRibbon.Items.Add(Me.ddToLang)
        Me.gpRibbon.Items.Add(Me.ButtonGroup1)
        Me.gpRibbon.Items.Add(Me.Separator1)
        Me.gpRibbon.Items.Add(Me.btnUndo)
        Me.gpRibbon.Items.Add(Me.btnTranslate)
        Me.gpRibbon.Items.Add(Me.lblProg)
        Me.gpRibbon.Label = "xTranslator"
        Me.gpRibbon.Name = "gpRibbon"
        '
        'ddFromLang
        '
        Me.ddFromLang.Label = "From"
        Me.ddFromLang.Name = "ddFromLang"
        Me.ddFromLang.SizeString = "en-us"
        '
        'ddToLang
        '
        Me.ddToLang.Label = "To"
        Me.ddToLang.Name = "ddToLang"
        Me.ddToLang.SizeString = "en-us"
        '
        'ButtonGroup1
        '
        Me.ButtonGroup1.Items.Add(Me.btnSwitch)
        Me.ButtonGroup1.Items.Add(Me.btnSpace)
        Me.ButtonGroup1.Items.Add(Me.btnAbout)
        Me.ButtonGroup1.Name = "ButtonGroup1"
        '
        'btnSwitch
        '
        Me.btnSwitch.Image = CType(resources.GetObject("btnSwitch.Image"), System.Drawing.Image)
        Me.btnSwitch.Label = "Button1"
        Me.btnSwitch.Name = "btnSwitch"
        Me.btnSwitch.ScreenTip = "Switch languages"
        Me.btnSwitch.ShowImage = True
        Me.btnSwitch.ShowLabel = False
        '
        'btnSpace
        '
        Me.btnSpace.Enabled = False
        Me.btnSpace.Label = "btnSpace"
        Me.btnSpace.Name = "btnSpace"
        Me.btnSpace.ShowLabel = False
        '
        'btnAbout
        '
        Me.btnAbout.Image = CType(resources.GetObject("btnAbout.Image"), System.Drawing.Image)
        Me.btnAbout.Label = "About"
        Me.btnAbout.Name = "btnAbout"
        Me.btnAbout.ScreenTip = "About xTranslator Add-in"
        Me.btnAbout.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'btnUndo
        '
        Me.btnUndo.Image = CType(resources.GetObject("btnUndo.Image"), System.Drawing.Image)
        Me.btnUndo.Label = "Undo"
        Me.btnUndo.Name = "btnUndo"
        Me.btnUndo.ScreenTip = "Undo translate"
        Me.btnUndo.ShowImage = True
        '
        'btnTranslate
        '
        Me.btnTranslate.Image = CType(resources.GetObject("btnTranslate.Image"), System.Drawing.Image)
        Me.btnTranslate.Label = "Translate"
        Me.btnTranslate.Name = "btnTranslate"
        Me.btnTranslate.ScreenTip = "Start translate selected cells"
        Me.btnTranslate.ShowImage = True
        '
        'lblProg
        '
        Me.lblProg.Label = "Starting..."
        Me.lblProg.Name = "lblProg"
        Me.lblProg.ShowLabel = False
        '
        'ribbon
        '
        Me.Name = "ribbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.gpRibbon.ResumeLayout(False)
        Me.gpRibbon.PerformLayout()
        Me.ButtonGroup1.ResumeLayout(False)
        Me.ButtonGroup1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents gpRibbon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ddFromLang As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ddToLang As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents btnAbout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSwitch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonGroup1 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents btnTranslate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSpace As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUndo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents lblProg As Microsoft.Office.Tools.Ribbon.RibbonLabel
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As ribbon
        Get
            Return Me.GetRibbon(Of ribbon)()
        End Get
    End Property
End Class
