<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container()
        Me.lstSelect = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstDisplay = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnQuit = New System.Windows.Forms.Button()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.btnEnterFile = New System.Windows.Forms.Button()
        Me.btnSelectFile = New System.Windows.Forms.Button()
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmbFilter = New System.Windows.Forms.ComboBox()
        Me.btnAddPatient = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.btnQuit.Location = New System.Drawing.Point(483, 356)
        Me.btnQuit.Margin = New System.Windows.Forms.Padding(2)
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.Size = New System.Drawing.Size(105, 28)
        Me.btnQuit.TabIndex = 11
        Me.btnQuit.Text = "Quit Program"
        Me.btnQuit.UseVisualStyleBackColor = True
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(231, 361)
        Me.txtFileName.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.ReadOnly = True
        Me.txtFileName.Size = New System.Drawing.Size(248, 20)
        Me.txtFileName.TabIndex = 10
        '
        'btnEnterFile
        '
        Me.btnEnterFile.Enabled = False
        Me.btnEnterFile.Location = New System.Drawing.Point(121, 356)
        Me.btnEnterFile.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEnterFile.Name = "btnEnterFile"
        Me.btnEnterFile.Size = New System.Drawing.Size(106, 28)
        Me.btnEnterFile.TabIndex = 9
        Me.btnEnterFile.Text = "Enter File Data"
        Me.btnEnterFile.UseVisualStyleBackColor = True
        '
        'btnSelectFile
        '
        Me.btnSelectFile.Location = New System.Drawing.Point(11, 356)
        Me.btnSelectFile.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(106, 28)
        Me.btnSelectFile.TabIndex = 8
        Me.btnSelectFile.Text = "Select Text File"
        Me.btnSelectFile.UseVisualStyleBackColor = True
        '
        'ErrorProvider
        '
        Me.ErrorProvider.ContainerControl = Me
        '
        'cmbFilter
        '
        Me.cmbFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFilter.FormattingEnabled = True
        Me.cmbFilter.Items.AddRange(New Object() {"Doctors", "Patients"})
        Me.cmbFilter.Location = New System.Drawing.Point(10, 263)
        Me.cmbFilter.MaxDropDownItems = 2
        Me.cmbFilter.Name = "cmbFilter"
        Me.cmbFilter.Size = New System.Drawing.Size(147, 21)
        Me.cmbFilter.TabIndex = 12
        '
        'btnAddPatient
        '
        Me.btnAddPatient.Location = New System.Drawing.Point(10, 297)
        Me.btnAddPatient.Margin = New System.Windows.Forms.Padding(2)
        Me.btnAddPatient.Name = "btnAddPatient"
        Me.btnAddPatient.Size = New System.Drawing.Size(175, 55)
        Me.btnAddPatient.TabIndex = 13
        Me.btnAddPatient.Text = "Add New Patient"
        Me.btnAddPatient.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(211, 297)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(175, 55)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Cancel Appointment"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(413, 297)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(175, 55)
        Me.Button2.TabIndex = 15
        Me.Button2.Text = "Change Appointment"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(600, 394)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnAddPatient)
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
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents ErrorProvider As ErrorProvider
    Friend WithEvents cmbFilter As ComboBox
    Friend WithEvents btnAddPatient As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
End Class
