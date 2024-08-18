<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmXL2Csv : Inherits Form

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Friend WithEvents sbStatusBar As System.Windows.Forms.StatusBar
    Friend WithEvents cmdConv As System.Windows.Forms.Button
    Friend WithEvents cmdAnnuler As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmXL2Csv))
        Me.sbStatusBar = New System.Windows.Forms.StatusBar()
        Me.cmdConv = New System.Windows.Forms.Button()
        Me.cmdAnnuler = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdAjouterMenuCtx = New System.Windows.Forms.Button()
        Me.cmdEnleverMenuCtx = New System.Windows.Forms.Button()
        Me.chkFusionCsv = New System.Windows.Forms.CheckBox()
        Me.chkTexte = New System.Windows.Forms.CheckBox()
        Me.chkODBC = New System.Windows.Forms.CheckBox()
        Me.chkAutomation = New System.Windows.Forms.CheckBox()
        Me.chkXL2Csv = New System.Windows.Forms.CheckBox()
        Me.chkXL2CsvNPOI = New System.Windows.Forms.CheckBox()
        Me.chkXL2CsvSSG = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'sbStatusBar
        '
        Me.sbStatusBar.Location = New System.Drawing.Point(0, 107)
        Me.sbStatusBar.Name = "sbStatusBar"
        Me.sbStatusBar.Size = New System.Drawing.Size(604, 22)
        Me.sbStatusBar.TabIndex = 0
        '
        'cmdConv
        '
        Me.cmdConv.Location = New System.Drawing.Point(173, 35)
        Me.cmdConv.Name = "cmdConv"
        Me.cmdConv.Size = New System.Drawing.Size(103, 32)
        Me.cmdConv.TabIndex = 1
        Me.cmdConv.Text = "Convertir"
        Me.ToolTip1.SetToolTip(Me.cmdConv, "Convertir le fichier Excel")
        '
        'cmdAnnuler
        '
        Me.cmdAnnuler.Enabled = False
        Me.cmdAnnuler.Location = New System.Drawing.Point(311, 35)
        Me.cmdAnnuler.Name = "cmdAnnuler"
        Me.cmdAnnuler.Size = New System.Drawing.Size(103, 32)
        Me.cmdAnnuler.TabIndex = 2
        Me.cmdAnnuler.Text = "Annuler"
        Me.ToolTip1.SetToolTip(Me.cmdAnnuler, "Interrompre la requête en cours, et renvoyer les données déjà récupérées")
        '
        'cmdAjouterMenuCtx
        '
        Me.cmdAjouterMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAjouterMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAjouterMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAjouterMenuCtx.Location = New System.Drawing.Point(173, 73)
        Me.cmdAjouterMenuCtx.Name = "cmdAjouterMenuCtx"
        Me.cmdAjouterMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAjouterMenuCtx.Size = New System.Drawing.Size(103, 25)
        Me.cmdAjouterMenuCtx.TabIndex = 35
        Me.cmdAjouterMenuCtx.Text = "Ajouter menu ctx."
        Me.ToolTip1.SetToolTip(Me.cmdAjouterMenuCtx, "Ajouter les menus contextuels pour convertir directement un fichier Excel depuis " & _
        "l'explorateur de fichiers (sous Windows Vista, il faut préalablement lancer l'ap" & _
        "plication en tant qu'admin.)")
        Me.cmdAjouterMenuCtx.UseVisualStyleBackColor = False
        '
        'cmdEnleverMenuCtx
        '
        Me.cmdEnleverMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnleverMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEnleverMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnleverMenuCtx.Location = New System.Drawing.Point(311, 73)
        Me.cmdEnleverMenuCtx.Name = "cmdEnleverMenuCtx"
        Me.cmdEnleverMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEnleverMenuCtx.Size = New System.Drawing.Size(103, 25)
        Me.cmdEnleverMenuCtx.TabIndex = 36
        Me.cmdEnleverMenuCtx.Text = "Enlever menu ctx."
        Me.ToolTip1.SetToolTip(Me.cmdEnleverMenuCtx, "Enlever les menus contextuels de XL2Csv (sous Windows Vista, il faut préalablemen" & _
        "t lancer l'application en tant qu'admin.)")
        Me.cmdEnleverMenuCtx.UseVisualStyleBackColor = False
        '
        'chkFusionCsv
        '
        Me.chkFusionCsv.AutoSize = True
        Me.chkFusionCsv.Location = New System.Drawing.Point(23, 55)
        Me.chkFusionCsv.Name = "chkFusionCsv"
        Me.chkFusionCsv.Size = New System.Drawing.Size(77, 17)
        Me.chkFusionCsv.TabIndex = 37
        Me.chkFusionCsv.Text = "Fusion csv"
        Me.chkFusionCsv.UseVisualStyleBackColor = True
        Me.chkFusionCsv.Visible = False
        '
        'chkTexte
        '
        Me.chkTexte.AutoSize = True
        Me.chkTexte.Location = New System.Drawing.Point(23, 75)
        Me.chkTexte.Name = "chkTexte"
        Me.chkTexte.Size = New System.Drawing.Size(53, 17)
        Me.chkTexte.TabIndex = 38
        Me.chkTexte.Text = "Texte"
        Me.chkTexte.UseVisualStyleBackColor = True
        Me.chkTexte.Visible = False
        '
        'chkODBC
        '
        Me.chkODBC.AutoSize = True
        Me.chkODBC.Location = New System.Drawing.Point(23, 15)
        Me.chkODBC.Name = "chkODBC"
        Me.chkODBC.Size = New System.Drawing.Size(56, 17)
        Me.chkODBC.TabIndex = 39
        Me.chkODBC.Text = "ODBC"
        Me.chkODBC.UseVisualStyleBackColor = True
        Me.chkODBC.Visible = False
        '
        'chkAutomation
        '
        Me.chkAutomation.AutoSize = True
        Me.chkAutomation.Location = New System.Drawing.Point(23, 35)
        Me.chkAutomation.Name = "chkAutomation"
        Me.chkAutomation.Size = New System.Drawing.Size(79, 17)
        Me.chkAutomation.TabIndex = 40
        Me.chkAutomation.Text = "Automation"
        Me.chkAutomation.UseVisualStyleBackColor = True
        Me.chkAutomation.Visible = False
        '
        'chkXL2Csv
        '
        Me.chkXL2Csv.AutoSize = True
        Me.chkXL2Csv.Location = New System.Drawing.Point(475, 15)
        Me.chkXL2Csv.Name = "chkXL2Csv"
        Me.chkXL2Csv.Size = New System.Drawing.Size(63, 17)
        Me.chkXL2Csv.TabIndex = 41
        Me.chkXL2Csv.Text = "XL2Csv"
        Me.chkXL2Csv.UseVisualStyleBackColor = True
        Me.chkXL2Csv.Visible = False
        '
        'chkXL2CsvNPOI
        '
        Me.chkXL2CsvNPOI.AutoSize = True
        Me.chkXL2CsvNPOI.Checked = True
        Me.chkXL2CsvNPOI.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkXL2CsvNPOI.Location = New System.Drawing.Point(475, 38)
        Me.chkXL2CsvNPOI.Name = "chkXL2CsvNPOI"
        Me.chkXL2CsvNPOI.Size = New System.Drawing.Size(92, 17)
        Me.chkXL2CsvNPOI.TabIndex = 42
        Me.chkXL2CsvNPOI.Text = "XL2Csv NPOI"
        Me.chkXL2CsvNPOI.UseVisualStyleBackColor = True
        Me.chkXL2CsvNPOI.Visible = False
        '
        'chkXL2CsvSSG
        '
        Me.chkXL2CsvSSG.AutoSize = True
        Me.chkXL2CsvSSG.Location = New System.Drawing.Point(475, 61)
        Me.chkXL2CsvSSG.Name = "chkXL2CsvSSG"
        Me.chkXL2CsvSSG.Size = New System.Drawing.Size(88, 17)
        Me.chkXL2CsvSSG.TabIndex = 43
        Me.chkXL2CsvSSG.Text = "XL2Csv SSG"
        Me.chkXL2CsvSSG.UseVisualStyleBackColor = True
        Me.chkXL2CsvSSG.Visible = False
        '
        'frmXL2Csv
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(604, 129)
        Me.Controls.Add(Me.chkXL2CsvSSG)
        Me.Controls.Add(Me.chkXL2CsvNPOI)
        Me.Controls.Add(Me.chkXL2Csv)
        Me.Controls.Add(Me.chkAutomation)
        Me.Controls.Add(Me.chkODBC)
        Me.Controls.Add(Me.chkTexte)
        Me.Controls.Add(Me.chkFusionCsv)
        Me.Controls.Add(Me.cmdAjouterMenuCtx)
        Me.Controls.Add(Me.cmdEnleverMenuCtx)
        Me.Controls.Add(Me.cmdAnnuler)
        Me.Controls.Add(Me.cmdConv)
        Me.Controls.Add(Me.sbStatusBar)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmXL2Csv"
        Me.Text = "XL2Csv : Convertir un fichier Excel en fichiers Csv"
        Me.ResumeLayout(False)
        Me.PerformLayout()

End Sub
    Public WithEvents cmdAjouterMenuCtx As System.Windows.Forms.Button
    Public WithEvents cmdEnleverMenuCtx As System.Windows.Forms.Button
    Friend WithEvents chkFusionCsv As System.Windows.Forms.CheckBox
    Friend WithEvents chkTexte As System.Windows.Forms.CheckBox
    Friend WithEvents chkODBC As System.Windows.Forms.CheckBox
    Friend WithEvents chkAutomation As System.Windows.Forms.CheckBox
    Friend WithEvents chkXL2Csv As System.Windows.Forms.CheckBox
    Friend WithEvents chkXL2CsvNPOI As System.Windows.Forms.CheckBox
    Friend WithEvents chkXL2CsvSSG As System.Windows.Forms.CheckBox

End Class
