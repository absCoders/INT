<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFCONV2
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE2
    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton
        Me.grpTATCONV1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraNumericEditor1 = New Infragistics.Win.UltraWinEditors.UltraNumericEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraDateTimeEditor3 = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
        Me.splLog = New System.Windows.Forms.SplitContainer
        Me.grpLog = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraTextEditor7 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.grpFollowup = New Infragistics.Win.Misc.UltraGroupBox
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.chkFollowUp = New ABSCS.ABSCheckBox
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        Me.lblCONV_FOLLOWUP_DATE = New Infragistics.Win.Misc.UltraLabel
        Me.UltraDateTimeEditor1 = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
        Me.UltraDateTimeEditor2 = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
        Me.txtCONV_FOLLOWUP_BY = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.lblCONV_SUBJECT = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor15 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.cmdAttach = New Infragistics.Win.Misc.UltraButton
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpTATCONV1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpTATCONV1.SuspendLayout()
        CType(Me.UltraNumericEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraDateTimeEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splLog.Panel1.SuspendLayout()
        Me.splLog.Panel2.SuspendLayout()
        Me.splLog.SuspendLayout()
        CType(Me.grpLog, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpLog.SuspendLayout()
        CType(Me.UltraTextEditor7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpFollowup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFollowup.SuspendLayout()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkFollowUp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraDateTimeEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraDateTimeEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCONV_FOLLOWUP_BY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor15, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(590, 360)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 360)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(590, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 360)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(590, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 360)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(590, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(515, 2)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(69, 33)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Location = New System.Drawing.Point(440, 2)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(69, 33)
        Me.cmdUpdate.TabIndex = 0
        Me.cmdUpdate.Text = "Update"
        '
        'grpTATCONV1
        '
        Me.grpTATCONV1.Controls.Add(Me.UltraLabel2)
        Me.grpTATCONV1.Controls.Add(Me.UltraNumericEditor1)
        Me.grpTATCONV1.Controls.Add(Me.UltraLabel1)
        Me.grpTATCONV1.Controls.Add(Me.UltraDateTimeEditor3)
        Me.grpTATCONV1.Controls.Add(Me.splLog)
        Me.grpTATCONV1.Controls.Add(Me.chkFollowUp)
        Me.grpTATCONV1.Controls.Add(Me.UltraLabel3)
        Me.grpTATCONV1.Controls.Add(Me.lblCONV_FOLLOWUP_DATE)
        Me.grpTATCONV1.Controls.Add(Me.UltraDateTimeEditor1)
        Me.grpTATCONV1.Controls.Add(Me.UltraDateTimeEditor2)
        Me.grpTATCONV1.Controls.Add(Me.txtCONV_FOLLOWUP_BY)
        Me.grpTATCONV1.Controls.Add(Me.lblCONV_SUBJECT)
        Me.grpTATCONV1.Controls.Add(Me.UltraTextEditor15)
        Me.grpTATCONV1.Dock = System.Windows.Forms.DockStyle.Fill
        Appearance1.ForeColor = System.Drawing.Color.Blue
        Me.grpTATCONV1.HeaderAppearance = Appearance1
        Me.grpTATCONV1.Location = New System.Drawing.Point(0, 0)
        Me.grpTATCONV1.Name = "grpTATCONV1"
        Me.grpTATCONV1.Size = New System.Drawing.Size(590, 316)
        Me.grpTATCONV1.TabIndex = 0
        '
        'UltraLabel2
        '
        Me.UltraLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(345, 261)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(67, 18)
        Me.UltraLabel2.TabIndex = 153
        Me.UltraLabel2.Text = "Promised"
        '
        'UltraNumericEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraNumericEditor1, "CONV_PROMISE_AMT")
        Me.UltraNumericEditor1.AlwaysInEditMode = True
        Me.UltraNumericEditor1.FormatString = ""
        Me.UltraNumericEditor1.Location = New System.Drawing.Point(345, 283)
        Me.UltraNumericEditor1.MaxValue = 9999999
        Me.UltraNumericEditor1.MinValue = 0
        Me.UltraNumericEditor1.Name = "UltraNumericEditor1"
        Me.UltraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.[Double]
        Me.UltraNumericEditor1.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.UltraNumericEditor1.Size = New System.Drawing.Size(108, 25)
        Me.UltraNumericEditor1.TabIndex = 152
        '
        'UltraLabel1
        '
        Me.UltraLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(462, 260)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(22, 18)
        Me.UltraLabel1.TabIndex = 132
        Me.UltraLabel1.Text = "By"
        '
        'UltraDateTimeEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraDateTimeEditor3, "CONV_PROMISE_BY")
        Me.UltraDateTimeEditor3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.UltraDateTimeEditor3.DateTime = New Date(2007, 1, 27, 0, 0, 0, 0)
        Me.UltraDateTimeEditor3.Location = New System.Drawing.Point(462, 283)
        Me.UltraDateTimeEditor3.Name = "UltraDateTimeEditor3"
        Me.UltraDateTimeEditor3.Size = New System.Drawing.Size(125, 25)
        Me.UltraDateTimeEditor3.TabIndex = 131
        Me.UltraDateTimeEditor3.Value = New Date(2007, 1, 27, 0, 0, 0, 0)
        '
        'splLog
        '
        Me.splLog.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.splLog.Location = New System.Drawing.Point(3, 65)
        Me.splLog.Name = "splLog"
        Me.splLog.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splLog.Panel1
        '
        Me.splLog.Panel1.Controls.Add(Me.grpLog)
        '
        'splLog.Panel2
        '
        Me.splLog.Panel2.Controls.Add(Me.grpFollowup)
        Me.splLog.Size = New System.Drawing.Size(587, 189)
        Me.splLog.SplitterDistance = 90
        Me.splLog.TabIndex = 130
        '
        'grpLog
        '
        Me.grpLog.Controls.Add(Me.UltraTextEditor7)
        Me.grpLog.Dock = System.Windows.Forms.DockStyle.Fill
        Appearance3.ForeColor = System.Drawing.Color.Blue
        Me.grpLog.HeaderAppearance = Appearance3
        Me.grpLog.Location = New System.Drawing.Point(0, 0)
        Me.grpLog.Name = "grpLog"
        Me.grpLog.Size = New System.Drawing.Size(587, 90)
        Me.grpLog.TabIndex = 131
        Me.grpLog.Text = "Log"
        '
        'UltraTextEditor7
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor7, "CONV_NOTES")
        Me.UltraTextEditor7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraTextEditor7.Location = New System.Drawing.Point(3, 20)
        Me.UltraTextEditor7.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraTextEditor7.Multiline = True
        Me.UltraTextEditor7.Name = "UltraTextEditor7"
        Me.UltraTextEditor7.Scrollbars = System.Windows.Forms.ScrollBars.Vertical
        Me.UltraTextEditor7.Size = New System.Drawing.Size(581, 67)
        Me.UltraTextEditor7.TabIndex = 1
        '
        'grpFollowup
        '
        Me.grpFollowup.Controls.Add(Me.UltraTextEditor1)
        Me.grpFollowup.Dock = System.Windows.Forms.DockStyle.Fill
        Appearance2.ForeColor = System.Drawing.Color.Blue
        Me.grpFollowup.HeaderAppearance = Appearance2
        Me.grpFollowup.Location = New System.Drawing.Point(0, 0)
        Me.grpFollowup.Name = "grpFollowup"
        Me.grpFollowup.Size = New System.Drawing.Size(587, 95)
        Me.grpFollowup.TabIndex = 131
        Me.grpFollowup.Text = "Follow-Up"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor1, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "CONV_FOLLOWUP_NOTES")
        Me.UltraTextEditor1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraTextEditor1.Location = New System.Drawing.Point(3, 20)
        Me.UltraTextEditor1.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraTextEditor1.Multiline = True
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.ReadOnly = True
        Me.UltraTextEditor1.Scrollbars = System.Windows.Forms.ScrollBars.Vertical
        Me.UltraTextEditor1.Size = New System.Drawing.Size(581, 72)
        Me.UltraTextEditor1.TabIndex = 2
        Me.UltraTextEditor1.TabStop = False
        '
        'chkFollowUp
        '
        Me.chkFollowUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkFollowUp.Location = New System.Drawing.Point(6, 260)
        Me.chkFollowUp.Name = "chkFollowUp"
        Me.chkFollowUp.Size = New System.Drawing.Size(162, 20)
        Me.chkFollowUp.TabIndex = 2
        Me.chkFollowUp.Text = "Needs Follow-Up by"
        '
        'UltraLabel3
        '
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(462, 12)
        Me.UltraLabel3.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(36, 18)
        Me.UltraLabel3.TabIndex = 127
        Me.UltraLabel3.Text = "Date"
        '
        'lblCONV_FOLLOWUP_DATE
        '
        Me.lblCONV_FOLLOWUP_DATE.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblCONV_FOLLOWUP_DATE.AutoSize = True
        Me.lblCONV_FOLLOWUP_DATE.Location = New System.Drawing.Point(185, 261)
        Me.lblCONV_FOLLOWUP_DATE.Margin = New System.Windows.Forms.Padding(4)
        Me.lblCONV_FOLLOWUP_DATE.Name = "lblCONV_FOLLOWUP_DATE"
        Me.lblCONV_FOLLOWUP_DATE.Size = New System.Drawing.Size(32, 18)
        Me.lblCONV_FOLLOWUP_DATE.TabIndex = 129
        Me.lblCONV_FOLLOWUP_DATE.Text = "Due"
        '
        'UltraDateTimeEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraDateTimeEditor1, "CONV_FOLLOWUP_DATE")
        Me.UltraDateTimeEditor1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.UltraDateTimeEditor1.DateTime = New Date(2007, 1, 27, 0, 0, 0, 0)
        Me.UltraDateTimeEditor1.Location = New System.Drawing.Point(185, 284)
        Me.UltraDateTimeEditor1.Name = "UltraDateTimeEditor1"
        Me.UltraDateTimeEditor1.Size = New System.Drawing.Size(125, 25)
        Me.UltraDateTimeEditor1.TabIndex = 4
        Me.UltraDateTimeEditor1.Value = New Date(2007, 1, 27, 0, 0, 0, 0)
        '
        'UltraDateTimeEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraDateTimeEditor2, "CONV_DATE")
        Me.UltraDateTimeEditor2.DateTime = New Date(2007, 1, 27, 0, 0, 0, 0)
        Me.UltraDateTimeEditor2.Location = New System.Drawing.Point(462, 31)
        Me.UltraDateTimeEditor2.Name = "UltraDateTimeEditor2"
        Me.UltraDateTimeEditor2.Size = New System.Drawing.Size(125, 25)
        Me.UltraDateTimeEditor2.TabIndex = 1
        Me.UltraDateTimeEditor2.TabStop = False
        Me.UltraDateTimeEditor2.Value = New Date(2007, 1, 27, 0, 0, 0, 0)
        '
        'txtCONV_FOLLOWUP_BY
        '
        Me.Absx1.SetABSColumnName(Me.txtCONV_FOLLOWUP_BY, "CONV_FOLLOWUP_BY")
        Me.Absx1.SetABSHasButton(Me.txtCONV_FOLLOWUP_BY, True)
        Me.Absx1.SetABSViewName(Me.txtCONV_FOLLOWUP_BY, "USER_ID")
        Me.txtCONV_FOLLOWUP_BY.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtCONV_FOLLOWUP_BY.Location = New System.Drawing.Point(6, 284)
        Me.txtCONV_FOLLOWUP_BY.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCONV_FOLLOWUP_BY.Name = "txtCONV_FOLLOWUP_BY"
        Me.txtCONV_FOLLOWUP_BY.Size = New System.Drawing.Size(162, 25)
        Me.txtCONV_FOLLOWUP_BY.TabIndex = 3
        '
        'lblCONV_SUBJECT
        '
        Me.lblCONV_SUBJECT.AutoSize = True
        Me.lblCONV_SUBJECT.Location = New System.Drawing.Point(6, 13)
        Me.lblCONV_SUBJECT.Margin = New System.Windows.Forms.Padding(4)
        Me.lblCONV_SUBJECT.Name = "lblCONV_SUBJECT"
        Me.lblCONV_SUBJECT.Size = New System.Drawing.Size(113, 18)
        Me.lblCONV_SUBJECT.TabIndex = 121
        Me.lblCONV_SUBJECT.Text = "Subject/Contact"
        '
        'UltraTextEditor15
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor15, "CONV_SUBJECT")
        Me.UltraTextEditor15.Location = New System.Drawing.Point(6, 32)
        Me.UltraTextEditor15.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraTextEditor15.Name = "UltraTextEditor15"
        Me.UltraTextEditor15.Size = New System.Drawing.Size(449, 25)
        Me.UltraTextEditor15.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grpTATCONV1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdAttach)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdUpdate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
        Me.SplitContainer1.Size = New System.Drawing.Size(590, 360)
        Me.SplitContainer1.SplitterDistance = 316
        Me.SplitContainer1.TabIndex = 10
        '
        'cmdAttach
        '
        Me.cmdAttach.Location = New System.Drawing.Point(6, 2)
        Me.cmdAttach.Name = "cmdAttach"
        Me.cmdAttach.Size = New System.Drawing.Size(69, 33)
        Me.cmdAttach.TabIndex = 2
        Me.cmdAttach.TabStop = False
        Me.cmdAttach.Text = "Attach"
        '
        'ASFCONV2
        '
        Me.Absx1.SetABSTableName(Me, "TATCONV1")
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(590, 360)
        Me.ControlBox = False
        Me.Name = "ASFCONV2"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Conversation Log"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpTATCONV1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpTATCONV1.ResumeLayout(False)
        Me.grpTATCONV1.PerformLayout()
        CType(Me.UltraNumericEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraDateTimeEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splLog.Panel1.ResumeLayout(False)
        Me.splLog.Panel2.ResumeLayout(False)
        Me.splLog.ResumeLayout(False)
        CType(Me.grpLog, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpLog.ResumeLayout(False)
        Me.grpLog.PerformLayout()
        CType(Me.UltraTextEditor7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpFollowup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFollowup.ResumeLayout(False)
        Me.grpFollowup.PerformLayout()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkFollowUp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraDateTimeEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraDateTimeEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCONV_FOLLOWUP_BY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor15, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grpTATCONV1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents lblCONV_SUBJECT As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor15 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraTextEditor7 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents chkFollowUp As ABSCS.ABSCheckBox
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraDateTimeEditor2 As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents lblCONV_FOLLOWUP_DATE As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraDateTimeEditor1 As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents txtCONV_FOLLOWUP_BY As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cmdAttach As Infragistics.Win.Misc.UltraButton
    Friend WithEvents splLog As System.Windows.Forms.SplitContainer
    Friend WithEvents grpLog As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents grpFollowup As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraDateTimeEditor3 As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraNumericEditor1 As Infragistics.Win.UltraWinEditors.UltraNumericEditor
End Class
