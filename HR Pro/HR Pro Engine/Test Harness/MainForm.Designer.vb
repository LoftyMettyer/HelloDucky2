﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
      Me.pnlScripting = New System.Windows.Forms.Panel()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.chkDebugMode = New System.Windows.Forms.CheckBox()
      Me.txtServer = New System.Windows.Forms.TextBox()
      Me.txtDatabase = New System.Windows.Forms.TextBox()
      Me.butScriptDB = New System.Windows.Forms.Button()
      Me.Panel1 = New System.Windows.Forms.Panel()
      Me.Button1 = New System.Windows.Forms.Button()
      Me.butImport = New System.Windows.Forms.Button()
      Me.txtServer2 = New System.Windows.Forms.TextBox()
      Me.txtDatabase2 = New System.Windows.Forms.TextBox()
      Me.txtPassword2 = New System.Windows.Forms.TextBox()
      Me.txtUser2 = New System.Windows.Forms.TextBox()
      Me.butViewObjects = New System.Windows.Forms.Button()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.pnlScripting.SuspendLayout()
      Me.Panel1.SuspendLayout()
      Me.SuspendLayout()
      '
      'pnlScripting
      '
      Me.pnlScripting.Controls.Add(Me.Label3)
      Me.pnlScripting.Controls.Add(Me.Label2)
      Me.pnlScripting.Controls.Add(Me.chkDebugMode)
      Me.pnlScripting.Controls.Add(Me.txtServer)
      Me.pnlScripting.Controls.Add(Me.txtDatabase)
      Me.pnlScripting.Controls.Add(Me.butScriptDB)
      Me.pnlScripting.Location = New System.Drawing.Point(15, 39)
      Me.pnlScripting.Name = "pnlScripting"
      Me.pnlScripting.Size = New System.Drawing.Size(308, 163)
      Me.pnlScripting.TabIndex = 21
      '
      'Label3
      '
      Me.Label3.AutoSize = True
      Me.Label3.Location = New System.Drawing.Point(16, 52)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(53, 13)
      Me.Label3.TabIndex = 26
      Me.Label3.Text = "Database"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(22, 27)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(38, 13)
      Me.Label2.TabIndex = 25
      Me.Label2.Text = "Server"
      '
      'chkDebugMode
      '
      Me.chkDebugMode.AutoSize = True
      Me.chkDebugMode.Checked = True
      Me.chkDebugMode.CheckState = System.Windows.Forms.CheckState.Checked
      Me.chkDebugMode.Location = New System.Drawing.Point(192, 23)
      Me.chkDebugMode.Name = "chkDebugMode"
      Me.chkDebugMode.Size = New System.Drawing.Size(88, 17)
      Me.chkDebugMode.TabIndex = 24
      Me.chkDebugMode.Text = "Debug Mode"
      Me.chkDebugMode.UseVisualStyleBackColor = True
      '
      'txtServer
      '
      Me.txtServer.Location = New System.Drawing.Point(75, 23)
      Me.txtServer.Name = "txtServer"
      Me.txtServer.Size = New System.Drawing.Size(100, 20)
      Me.txtServer.TabIndex = 23
      Me.txtServer.Text = "COA14270"
      '
      'txtDatabase
      '
      Me.txtDatabase.Location = New System.Drawing.Point(75, 49)
      Me.txtDatabase.Name = "txtDatabase"
      Me.txtDatabase.Size = New System.Drawing.Size(100, 20)
      Me.txtDatabase.TabIndex = 22
      Me.txtDatabase.Text = "openhr5"
      '
      'butScriptDB
      '
      Me.butScriptDB.Location = New System.Drawing.Point(71, 84)
      Me.butScriptDB.Name = "butScriptDB"
      Me.butScriptDB.Size = New System.Drawing.Size(104, 45)
      Me.butScriptDB.TabIndex = 21
      Me.butScriptDB.Text = "Go Script."
      Me.butScriptDB.UseVisualStyleBackColor = True
      '
      'Panel1
      '
      Me.Panel1.Controls.Add(Me.Button1)
      Me.Panel1.Controls.Add(Me.butImport)
      Me.Panel1.Controls.Add(Me.txtServer2)
      Me.Panel1.Controls.Add(Me.txtDatabase2)
      Me.Panel1.Controls.Add(Me.txtPassword2)
      Me.Panel1.Controls.Add(Me.txtUser2)
      Me.Panel1.Controls.Add(Me.butViewObjects)
      Me.Panel1.Location = New System.Drawing.Point(348, 39)
      Me.Panel1.Name = "Panel1"
      Me.Panel1.Size = New System.Drawing.Size(423, 163)
      Me.Panel1.TabIndex = 22
      '
      'Button1
      '
      Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Button1.Location = New System.Drawing.Point(294, 70)
      Me.Button1.Name = "Button1"
      Me.Button1.Size = New System.Drawing.Size(117, 59)
      Me.Button1.TabIndex = 26
      Me.Button1.Text = "THIS ONE - Audit" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
      Me.Button1.UseVisualStyleBackColor = True
      '
      'butImport
      '
      Me.butImport.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.butImport.Location = New System.Drawing.Point(161, 70)
      Me.butImport.Name = "butImport"
      Me.butImport.Size = New System.Drawing.Size(127, 40)
      Me.butImport.TabIndex = 5
      Me.butImport.Text = "Import"
      Me.butImport.UseVisualStyleBackColor = True
      Me.butImport.Visible = False
      '
      'txtServer2
      '
      Me.txtServer2.Location = New System.Drawing.Point(121, 40)
      Me.txtServer2.Name = "txtServer2"
      Me.txtServer2.Size = New System.Drawing.Size(100, 20)
      Me.txtServer2.TabIndex = 4
      Me.txtServer2.Text = "harpdev02"
      '
      'txtDatabase2
      '
      Me.txtDatabase2.Location = New System.Drawing.Point(14, 40)
      Me.txtDatabase2.Name = "txtDatabase2"
      Me.txtDatabase2.Size = New System.Drawing.Size(100, 20)
      Me.txtDatabase2.TabIndex = 3
      Me.txtDatabase2.Text = "std41"
      '
      'txtPassword2
      '
      Me.txtPassword2.Location = New System.Drawing.Point(121, 13)
      Me.txtPassword2.Name = "txtPassword2"
      Me.txtPassword2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
      Me.txtPassword2.Size = New System.Drawing.Size(100, 20)
      Me.txtPassword2.TabIndex = 2
      Me.txtPassword2.Text = "asr"
      '
      'txtUser2
      '
      Me.txtUser2.Location = New System.Drawing.Point(14, 13)
      Me.txtUser2.Name = "txtUser2"
      Me.txtUser2.Size = New System.Drawing.Size(100, 20)
      Me.txtUser2.TabIndex = 1
      Me.txtUser2.Text = "sa"
      '
      'butViewObjects
      '
      Me.butViewObjects.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.butViewObjects.Location = New System.Drawing.Point(25, 70)
      Me.butViewObjects.Name = "butViewObjects"
      Me.butViewObjects.Size = New System.Drawing.Size(129, 40)
      Me.butViewObjects.TabIndex = 0
      Me.butViewObjects.Text = "Export"
      Me.butViewObjects.UseVisualStyleBackColor = True
      Me.butViewObjects.Visible = False
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(346, 9)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(116, 18)
      Me.Label1.TabIndex = 23
      Me.Label1.Text = "Export Selection"
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label4.Location = New System.Drawing.Point(12, 8)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(59, 18)
      Me.Label4.TabIndex = 24
      Me.Label4.Text = "Scripter"
      '
      'MainForm
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(806, 426)
      Me.Controls.Add(Me.Label4)
      Me.Controls.Add(Me.Label1)
      Me.Controls.Add(Me.Panel1)
      Me.Controls.Add(Me.pnlScripting)
      Me.Name = "MainForm"
      Me.Text = "DB"
      Me.pnlScripting.ResumeLayout(False)
      Me.pnlScripting.PerformLayout()
      Me.Panel1.ResumeLayout(False)
      Me.Panel1.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
    Friend WithEvents pnlScripting As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkDebugMode As System.Windows.Forms.CheckBox
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents butScriptDB As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents butViewObjects As System.Windows.Forms.Button
    Friend WithEvents txtServer2 As System.Windows.Forms.TextBox
    Friend WithEvents txtDatabase2 As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword2 As System.Windows.Forms.TextBox
    Friend WithEvents txtUser2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents butImport As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
