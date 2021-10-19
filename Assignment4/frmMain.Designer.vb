<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lstSelect = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstDisplay = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnQuit = New System.Windows.Forms.Button()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.btnEnterFile = New System.Windows.Forms.Button()
        Me.btnSelectFile = New System.Windows.Forms.Button()
        Me.cmbFilter = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'lstSelect
        '
        Me.lstSelect.FormattingEnabled = True
        Me.lstSelect.Location = New System.Drawing.Point(10, 20)
        Me.lstSelect.Margin = New System.Windows.Forms.Padding(2)
        Me.lstSelect.Name = "lstSelect"
        Me.lstSelect.Size = New System.Drawing.Size(147, 238)
        Me.lstSelect.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 5)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Selection Box"
        '
        'lstDisplay
        '
        Me.lstDisplay.FormattingEnabled = True
        Me.lstDisplay.Location = New System.Drawing.Point(169, 20)
        Me.lstDisplay.Margin = New System.Windows.Forms.Padding(2)
        Me.lstDisplay.Name = "lstDisplay"
        Me.lstDisplay.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.lstDisplay.Size = New System.Drawing.Size(420, 264)
        Me.lstDisplay.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(166, 5)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Display Box"
        '
        'btnQuit
        '
        Me.btnQuit.Location = New System.Drawing.Point(482, 289)
        Me.btnQuit.Margin = New System.Windows.Forms.Padding(2)
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.Size = New System.Drawing.Size(105, 28)
        Me.btnQuit.TabIndex = 11
        Me.btnQuit.Text = "Quit Program"
        Me.btnQuit.UseVisualStyleBackColor = True
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(230, 294)
        Me.txtFileName.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.ReadOnly = True
        Me.txtFileName.Size = New System.Drawing.Size(248, 20)
        Me.txtFileName.TabIndex = 10
        Me.txtFileName.TabStop = False
        '
        'btnEnterFile
        '
        Me.btnEnterFile.Enabled = False
        Me.btnEnterFile.Location = New System.Drawing.Point(120, 289)
        Me.btnEnterFile.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEnterFile.Name = "btnEnterFile"
        Me.btnEnterFile.Size = New System.Drawing.Size(106, 28)
        Me.btnEnterFile.TabIndex = 9
        Me.btnEnterFile.Text = "Enter File Data"
        Me.btnEnterFile.UseVisualStyleBackColor = True
        '
        'btnSelectFile
        '
        Me.btnSelectFile.Location = New System.Drawing.Point(10, 289)
        Me.btnSelectFile.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(106, 28)
        Me.btnSelectFile.TabIndex = 8
        Me.btnSelectFile.Text = "Select Text File"
        Me.btnSelectFile.UseVisualStyleBackColor = True
        '
        'cmbFilter
        '
        Me.cmbFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFilter.FormattingEnabled = True
        Me.cmbFilter.Items.AddRange(New Object() {"Doctor Appointments", "Doctor Availability", "Patients"})
        Me.cmbFilter.Location = New System.Drawing.Point(10, 263)
        Me.cmbFilter.MaxDropDownItems = 2
        Me.cmbFilter.Name = "cmbFilter"
        Me.cmbFilter.Size = New System.Drawing.Size(147, 21)
        Me.cmbFilter.TabIndex = 12
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(600, 324)
        Me.Controls.Add(Me.cmbFilter)
        Me.Controls.Add(Me.btnQuit)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.btnEnterFile)
        Me.Controls.Add(Me.btnSelectFile)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lstDisplay)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstSelect)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmMain"
        Me.Text = "Scheduler"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstSelect As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lstDisplay As ListBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnQuit As Button
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents btnEnterFile As Button
    Friend WithEvents btnSelectFile As Button
    Friend WithEvents cmbFilter As ComboBox
End Class
