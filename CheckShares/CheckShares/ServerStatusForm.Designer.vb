<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ServerStatusForm
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ServerStatusForm))
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ScrollPanel = New System.Windows.Forms.Panel
        Me.StatusPanel = New System.Windows.Forms.FlowLayoutPanel
        Me.StatusToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.AfterLoadTimer = New System.Windows.Forms.Timer(Me.components)
        Me.ScrollPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Network Path"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(211, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Free Space"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(421, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Last Time"
        Me.Label4.Visible = False
        '
        'ScrollPanel
        '
        Me.ScrollPanel.AutoScroll = True
        Me.ScrollPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ScrollPanel.Controls.Add(Me.StatusPanel)
        Me.ScrollPanel.Location = New System.Drawing.Point(12, 25)
        Me.ScrollPanel.Name = "ScrollPanel"
        Me.ScrollPanel.Size = New System.Drawing.Size(620, 302)
        Me.ScrollPanel.TabIndex = 4
        '
        'StatusPanel
        '
        Me.StatusPanel.BackColor = System.Drawing.SystemColors.Window
        Me.StatusPanel.Location = New System.Drawing.Point(0, 0)
        Me.StatusPanel.Margin = New System.Windows.Forms.Padding(0)
        Me.StatusPanel.Name = "StatusPanel"
        Me.StatusPanel.Size = New System.Drawing.Size(601, 301)
        Me.StatusPanel.TabIndex = 0
        '
        'AfterLoadTimer
        '
        '
        'ServerStatusForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(645, 340)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ScrollPanel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ServerStatusForm"
        Me.Text = "Server Status"
        Me.ScrollPanel.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ScrollPanel As System.Windows.Forms.Panel
    Friend WithEvents StatusToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents AfterLoadTimer As System.Windows.Forms.Timer
    Friend WithEvents StatusPanel As System.Windows.Forms.FlowLayoutPanel

End Class
