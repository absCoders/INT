<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFBASE0
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.UltraGridExcelExporter1 = New Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter(Me.components)
        Me._ASFBASE1_Toolbars_Dock_Area_Left = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me.tlb = New Infragistics.Win.UltraWinToolbars.UltraToolbarsManager(Me.components)
        Me._ASFBASE1_Toolbars_Dock_Area_Right = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me._ASFBASE1_Toolbars_Dock_Area_Top = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom = New Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea()
        Me.tip = New Infragistics.Win.UltraWinToolTip.UltraToolTipManager(Me.components)
        Me.Absx1 = New ABSX.ABSX()
        Me.enhancedGrdFilter = New Infragistics.Win.SupportDialogs.FilterUIProvider.UltraGridFilterUIProvider(Me.components)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraGridExcelExporter1
        '
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._ASFBASE1_Toolbars_Dock_Area_Left.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me._ASFBASE1_Toolbars_Dock_Area_Left.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Left
        Me._ASFBASE1_Toolbars_Dock_Area_Left.ForeColor = System.Drawing.SystemColors.ControlText
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Location = New System.Drawing.Point(0, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Name = "_ASFBASE1_Toolbars_Dock_Area_Left"
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 616)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.ToolbarsManager = Me.tlb
        '
        'tlb
        '
        Me.tlb.DesignerFlags = 1
        Me.tlb.DockWithinContainer = Me
        Me.tlb.DockWithinContainerBaseType = GetType(System.Windows.Forms.Form)
        Me.tlb.MdiMergeable = False
        Me.tlb.ShowFullMenusDelay = 500
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._ASFBASE1_Toolbars_Dock_Area_Right.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me._ASFBASE1_Toolbars_Dock_Area_Right.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Right
        Me._ASFBASE1_Toolbars_Dock_Area_Right.ForeColor = System.Drawing.SystemColors.ControlText
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(992, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Name = "_ASFBASE1_Toolbars_Dock_Area_Right"
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 616)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.ToolbarsManager = Me.tlb
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._ASFBASE1_Toolbars_Dock_Area_Top.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me._ASFBASE1_Toolbars_Dock_Area_Top.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Top
        Me._ASFBASE1_Toolbars_Dock_Area_Top.ForeColor = System.Drawing.SystemColors.ControlText
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Location = New System.Drawing.Point(0, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Name = "_ASFBASE1_Toolbars_Dock_Area_Top"
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(992, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.ToolbarsManager = Me.tlb
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Bottom
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.ForeColor = System.Drawing.SystemColors.ControlText
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 616)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Name = "_ASFBASE1_Toolbars_Dock_Area_Bottom"
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(992, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.ToolbarsManager = Me.tlb
        '
        'tip
        '
        Me.tip.ContainingControl = Me
        '
        'ASFBASE0
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(992, 616)
        Me.Controls.Add(Me._ASFBASE1_Toolbars_Dock_Area_Left)
        Me.Controls.Add(Me._ASFBASE1_Toolbars_Dock_Area_Right)
        Me.Controls.Add(Me._ASFBASE1_Toolbars_Dock_Area_Bottom)
        Me.Controls.Add(Me._ASFBASE1_Toolbars_Dock_Area_Top)
        Me.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Name = "ASFBASE0"
        Me.Text = "ASFBASE0"
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Protected Friend WithEvents Absx1 As ABSX.ABSX
    Friend WithEvents UltraGridExcelExporter1 As Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter
    Protected Friend WithEvents _ASFBASE1_Toolbars_Dock_Area_Left As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Protected Friend WithEvents _ASFBASE1_Toolbars_Dock_Area_Right As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Protected Friend WithEvents _ASFBASE1_Toolbars_Dock_Area_Top As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Protected Friend WithEvents _ASFBASE1_Toolbars_Dock_Area_Bottom As Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea
    Protected Friend WithEvents tlb As Infragistics.Win.UltraWinToolbars.UltraToolbarsManager
    Protected Friend WithEvents tip As Infragistics.Win.UltraWinToolTip.UltraToolTipManager
    Friend WithEvents UltraGridFilterUIProvider1 As Infragistics.Win.SupportDialogs.FilterUIProvider.UltraGridFilterUIProvider
    Friend WithEvents enhancedGrdFilter As Infragistics.Win.SupportDialogs.FilterUIProvider.UltraGridFilterUIProvider
End Class
