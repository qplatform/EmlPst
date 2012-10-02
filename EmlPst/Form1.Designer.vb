<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

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
        Me.FB = New System.Windows.Forms.FolderBrowserDialog()
        Me.lblFolder = New System.Windows.Forms.Label()
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstFiles = New System.Windows.Forms.ListBox()
        Me.lblProcess = New System.Windows.Forms.Label()
        Me.lblRec = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.PB1 = New System.Windows.Forms.ProgressBar()
        Me.chkRename = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'FB
        '
        Me.FB.ShowNewFolderButton = False
        '
        'lblFolder
        '
        Me.lblFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFolder.Location = New System.Drawing.Point(88, 10)
        Me.lblFolder.Name = "lblFolder"
        Me.lblFolder.Size = New System.Drawing.Size(544, 23)
        Me.lblFolder.TabIndex = 0
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowse.Location = New System.Drawing.Point(638, 10)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(75, 23)
        Me.cmdBrowse.TabIndex = 1
        Me.cmdBrowse.Text = "Browse ..."
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Selected Folder"
        '
        'lstFiles
        '
        Me.lstFiles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstFiles.FormattingEnabled = True
        Me.lstFiles.Location = New System.Drawing.Point(8, 41)
        Me.lstFiles.Name = "lstFiles"
        Me.lstFiles.Size = New System.Drawing.Size(705, 342)
        Me.lstFiles.TabIndex = 3
        '
        'lblProcess
        '
        Me.lblProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblProcess.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblProcess.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcess.ForeColor = System.Drawing.Color.Red
        Me.lblProcess.Location = New System.Drawing.Point(88, 399)
        Me.lblProcess.Name = "lblProcess"
        Me.lblProcess.Size = New System.Drawing.Size(139, 23)
        Me.lblProcess.TabIndex = 4
        Me.lblProcess.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblRec
        '
        Me.lblRec.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblRec.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblRec.Location = New System.Drawing.Point(88, 427)
        Me.lblRec.Name = "lblRec"
        Me.lblRec.Size = New System.Drawing.Size(435, 23)
        Me.lblRec.TabIndex = 5
        Me.lblRec.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 404)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Overall Process"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 432)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(42, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Record"
        '
        'cmdStart
        '
        Me.cmdStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStart.Location = New System.Drawing.Point(638, 399)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(75, 51)
        Me.cmdStart.TabIndex = 8
        Me.cmdStart.Text = "Start"
        Me.cmdStart.UseVisualStyleBackColor = True
        '
        'PB1
        '
        Me.PB1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PB1.Location = New System.Drawing.Point(233, 399)
        Me.PB1.Name = "PB1"
        Me.PB1.Size = New System.Drawing.Size(399, 23)
        Me.PB1.TabIndex = 9
        '
        'chkRename
        '
        Me.chkRename.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkRename.AutoSize = True
        Me.chkRename.Location = New System.Drawing.Point(529, 429)
        Me.chkRename.Name = "chkRename"
        Me.chkRename.Size = New System.Drawing.Size(110, 17)
        Me.chkRename.TabIndex = 10
        Me.chkRename.Text = "Rename Imported"
        Me.chkRename.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(719, 459)
        Me.Controls.Add(Me.chkRename)
        Me.Controls.Add(Me.PB1)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblRec)
        Me.Controls.Add(Me.lblProcess)
        Me.Controls.Add(Me.lstFiles)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.lblFolder)
        Me.Name = "Form1"
        Me.Text = "Import EML to PST"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FB As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents lblFolder As System.Windows.Forms.Label
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstFiles As System.Windows.Forms.ListBox
    Friend WithEvents lblProcess As System.Windows.Forms.Label
    Friend WithEvents lblRec As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents PB1 As System.Windows.Forms.ProgressBar
    Friend WithEvents chkRename As System.Windows.Forms.CheckBox

End Class
