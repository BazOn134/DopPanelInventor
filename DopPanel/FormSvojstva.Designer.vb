<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSvojstva
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
        Me.ButonZakryt = New System.Windows.Forms.Button()
        Me.TB_Name = New System.Windows.Forms.TextBox()
        Me.TB_PartNumber = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TB_Title = New System.Windows.Forms.TextBox()
        Me.TB_Description = New System.Windows.Forms.TextBox()
        Me.TB_Material = New System.Windows.Forms.TextBox()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.B_Sledujushij = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ButonZakryt
        '
        Me.ButonZakryt.Location = New System.Drawing.Point(365, 298)
        Me.ButonZakryt.Name = "ButonZakryt"
        Me.ButonZakryt.Size = New System.Drawing.Size(325, 34)
        Me.ButonZakryt.TabIndex = 0
        Me.ButonZakryt.Text = "Закрыть"
        Me.ButonZakryt.UseVisualStyleBackColor = True
        '
        'TB_Name
        '
        Me.TB_Name.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TB_Name.Location = New System.Drawing.Point(16, 12)
        Me.TB_Name.Name = "TB_Name"
        Me.TB_Name.Size = New System.Drawing.Size(674, 29)
        Me.TB_Name.TabIndex = 1
        '
        'TB_PartNumber
        '
        Me.TB_PartNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TB_PartNumber.Location = New System.Drawing.Point(34, 60)
        Me.TB_PartNumber.Name = "TB_PartNumber"
        Me.TB_PartNumber.Size = New System.Drawing.Size(549, 26)
        Me.TB_PartNumber.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(589, 66)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Обозначение"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(77, 98)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Наименование"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.Location = New System.Drawing.Point(77, 132)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Наименование"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label5.Location = New System.Drawing.Point(13, 165)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 16)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Материал"
        '
        'TB_Title
        '
        Me.TB_Title.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TB_Title.Location = New System.Drawing.Point(190, 92)
        Me.TB_Title.Name = "TB_Title"
        Me.TB_Title.Size = New System.Drawing.Size(500, 26)
        Me.TB_Title.TabIndex = 8
        '
        'TB_Description
        '
        Me.TB_Description.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TB_Description.Location = New System.Drawing.Point(190, 126)
        Me.TB_Description.Name = "TB_Description"
        Me.TB_Description.Size = New System.Drawing.Size(500, 26)
        Me.TB_Description.TabIndex = 9
        '
        'TB_Material
        '
        Me.TB_Material.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TB_Material.Location = New System.Drawing.Point(93, 160)
        Me.TB_Material.Name = "TB_Material"
        Me.TB_Material.Size = New System.Drawing.Size(597, 26)
        Me.TB_Material.TabIndex = 11
        '
        'ListBox1
        '
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Location = New System.Drawing.Point(120, 236)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(207, 96)
        Me.ListBox1.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label4.Location = New System.Drawing.Point(120, 202)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(229, 16)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Фамилии в основной надписи"
        '
        'B_Sledujushij
        '
        Me.B_Sledujushij.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.B_Sledujushij.Location = New System.Drawing.Point(365, 236)
        Me.B_Sledujushij.Name = "B_Sledujushij"
        Me.B_Sledujushij.Size = New System.Drawing.Size(325, 55)
        Me.B_Sledujushij.TabIndex = 13
        Me.B_Sledujushij.Text = "Следующий элемент"
        Me.B_Sledujushij.UseVisualStyleBackColor = True
        '
        'FormSvojstva
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(702, 358)
        Me.Controls.Add(Me.B_Sledujushij)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.TB_Material)
        Me.Controls.Add(Me.TB_Description)
        Me.Controls.Add(Me.TB_Title)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TB_PartNumber)
        Me.Controls.Add(Me.TB_Name)
        Me.Controls.Add(Me.ButonZakryt)
        Me.Name = "FormSvojstva"
        Me.Text = "Свойства"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButonZakryt As System.Windows.Forms.Button
    Friend WithEvents TB_Name As System.Windows.Forms.TextBox
    Friend WithEvents TB_PartNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TB_Title As System.Windows.Forms.TextBox
    Friend WithEvents TB_Description As System.Windows.Forms.TextBox
    Friend WithEvents TB_Material As System.Windows.Forms.TextBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents B_Sledujushij As System.Windows.Forms.Button
End Class
