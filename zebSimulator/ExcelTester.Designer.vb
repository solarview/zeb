﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExcelTester
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtDisplay = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdImportExcel = New System.Windows.Forms.Button()
        Me.FillWithStrings = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'txtDisplay
        '
        Me.txtDisplay.Location = New System.Drawing.Point(132, 90)
        Me.txtDisplay.Multiline = True
        Me.txtDisplay.Name = "txtDisplay"
        Me.txtDisplay.Size = New System.Drawing.Size(377, 237)
        Me.txtDisplay.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(300, 40)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cmdImportExcel
        '
        Me.cmdImportExcel.Location = New System.Drawing.Point(403, 40)
        Me.cmdImportExcel.Name = "cmdImportExcel"
        Me.cmdImportExcel.Size = New System.Drawing.Size(121, 23)
        Me.cmdImportExcel.TabIndex = 2
        Me.cmdImportExcel.Text = "Import Excel File ..."
        Me.cmdImportExcel.UseVisualStyleBackColor = True
        '
        'FillWithStrings
        '
        Me.FillWithStrings.AutoSize = True
        Me.FillWithStrings.Location = New System.Drawing.Point(603, 90)
        Me.FillWithStrings.Name = "FillWithStrings"
        Me.FillWithStrings.Size = New System.Drawing.Size(81, 17)
        Me.FillWithStrings.TabIndex = 3
        Me.FillWithStrings.Text = "CheckBox1"
        Me.FillWithStrings.UseVisualStyleBackColor = True
        '
        'ExcelTester
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.FillWithStrings)
        Me.Controls.Add(Me.cmdImportExcel)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtDisplay)
        Me.Name = "ExcelTester"
        Me.Text = "MainForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtDisplay As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents cmdImportExcel As Button
    Friend WithEvents FillWithStrings As CheckBox
End Class
