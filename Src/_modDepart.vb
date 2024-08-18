
' Fichier modDepart.vb
' --------------------

' XL2Csv : Convertir un fichier Excel en fichiers Csv

Module _modDepart

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
        Public Const bDebug As Boolean = False
        Public Const bRelease As Boolean = True
#End If

    Public sNomAppli$ = My.Application.Info.Title
    Public m_sTitreMsg$ = sNomAppli
    Public Const sTitreMsgDescription$ = " : Convertir un fichier Excel en fichiers Csv"

    Private Const sDateVersionXL2Csv$ = "18/08/2024"
    Public Const sDateVersionAppli$ = sDateVersionXL2Csv

    'Public Const bExcel2007SupportNPOI As Boolean = False
    Public Const bExcel2007SupportSSG As Boolean = True
    Public Const dDateNulle As Date = #12:00:00 AM#

    Public Const sExtXlsx$ = ".xlsx"

    Public Sub Main()

        If bDebug Then Depart() : Exit Sub

        Try
            Depart()
        Catch ex As Exception
            AfficherMsgErreur2(ex, "Main " & m_sTitreMsg)
        End Try

    End Sub

    Public Sub Depart()

        ' On peut démarrer l'application sur la feuille, ou bien sur la procédure 
        '  Main() si on veut pouvoir détecter l'absence de la dll sans plantage

        ' Extraire les options passées en argument de la ligne de commande
        ' Cette fct ne marche pas avec des chemins contenant des espaces, même entre guillemets
        'Dim asArgs$() = Environment.GetCommandLineArgs()
        Dim sArg0$ = Microsoft.VisualBasic.Interaction.Command
        Dim sCheminFichier$ = ""
        Dim iTypeConv As frmXL2Csv.TypeConv = frmXL2Csv.TypeConv.XL2Csv
        Dim bSyntaxeOk As Boolean = False
        Dim iNbArguments% = 0

        If sArg0 <> "" Then
            Dim asArgs$() = asArgLigneCmd(sArg0)
            iNbArguments = UBound(asArgs) + 1
            'If iNbArguments <= 2 Then bSyntaxeOk = True
            If iNbArguments = 1 Or iNbArguments = 2 Then bSyntaxeOk = True
            If Not bSyntaxeOk Then GoTo Suite
            If iNbArguments = 1 Then
                sCheminFichier = asArgs(0)
                If Not bFichierExiste(sCheminFichier, bPrompt:=True) Then _
                    bSyntaxeOk = False

                ' S'il n'y a qu'un seul argument et que l'extension du fichier est .xlsx alors SSG
                If sCheminFichier.ToLower.EndsWith(sExtXlsx) Then
                    'iTypeConv = frmXL2Csv.TypeConv.XL2CsvNPOI Pas encore implémenté
                    If bSpreadSheetGear Then iTypeConv = frmXL2Csv.TypeConv.XL2CsvSSG ' 09/08/2012
                End If

                GoTo Suite
            End If
            Dim sCmd$ = asArgs(1)
            'If bDebug Then MsgBox("Commande : " & sCmd, MsgBoxStyle.Information, m_sTitreMsg)
            If sCmd = frmXL2Csv.sXL2Csv Then
                iTypeConv = frmXL2Csv.TypeConv.XL2Csv
            ElseIf sCmd = frmXL2Csv.sXL2CsvNPOI Then
                iTypeConv = frmXL2Csv.TypeConv.XL2CsvNPOI
            ElseIf sCmd = frmXL2Csv.sXL2CsvSSG Then
                iTypeConv = frmXL2Csv.TypeConv.XL2CsvSSG ' 09/08/2012
            ElseIf sCmd = frmXL2Csv.sXL2CsvAutomation Then
                iTypeConv = frmXL2Csv.TypeConv.XL2CsvAutomation
            ElseIf sCmd = frmXL2Csv.sXL2CsvODBC Then
                iTypeConv = frmXL2Csv.TypeConv.XL2CsvODBC
            ElseIf sCmd = frmXL2Csv.sXL2Txt Then
                iTypeConv = frmXL2Csv.TypeConv.XL2Txt
            ElseIf sCmd = frmXL2Csv.sXL2CsvGroup Then
                iTypeConv = frmXL2Csv.TypeConv.XL2CsvGroup
            Else
                MsgBox("Commande non reconnue : " & sCmd,
                    MsgBoxStyle.Information, m_sTitreMsg & sTitreMsgDescription)
                bSyntaxeOk = False
            End If
            sCheminFichier = asArgs(0)
            If Not bFichierExiste(sCheminFichier, bPrompt:=True) Then _
                bSyntaxeOk = False
        End If
Suite:
        If Not bSyntaxeOk Then
            Dim sTxt$ =
                "Syntaxe : Chemin du fichier Excel à convertir" & vbCrLf &
                "en autant de fichiers Csv qu'il y a de feuille Excel" & vbCrLf &
                "Options possibles :" & vbCrLf &
                "XL2CsvNPOI : utiliser la librairie NPOI au lieu d'ExcelLibrary" & vbCrLf
            If bSpreadSheetGear Then sTxt &=
                "XL2CsvSSG : utiliser la librairie SpreadSheetGear au lieu d'ExcelLibrary" & vbCrLf
            sTxt &=
                "XL2Txt : pour convertir en un seul fichier Texte (via ExcelLibrary)" & vbCrLf &
                "XL2CsvGroup : pour convertir en un seul fichier Csv fusionné" & vbCrLf &
                "XL2CsvODBC : comme XL2Csv mais via ODBC : colonnes homogènes" & vbCrLf &
                "XL2CsvAutomation : comme XL2Csv mais via Automation Excel" & vbCrLf &
                "Exemples : " & vbCrLf &
                "XL2Csv.exe C:\Tmp\MonFichierExcel" & vbCrLf &
                "XL2Csv.exe C:\Tmp\MonFichierExcel XL2CsvNPOI" & vbCrLf
            If bSpreadSheetGear Then sTxt &=
                "XL2Csv.exe C:\Tmp\MonFichierExcel XL2CsvSSG" & vbCrLf
            sTxt &=
                "XL2Csv.exe C:\Tmp\MonFichierExcel XL2CsvODBC" & vbCrLf &
                "XL2Csv.exe C:\Tmp\MonFichierExcel XL2CsvAutomation" & vbCrLf &
                "XL2Csv.exe C:\Tmp\MonFichierExcel XL2Txt" & vbCrLf &
                "XL2Csv.exe C:\Tmp\MonFichierExcel XL2CsvGroup" & vbCrLf &
                "Sinon ajouter les menus contextuels via le menu dédié" & vbCrLf &
                "(utilisation des menus contextuels avec le bouton droit" & vbCrLf &
                " de la souris dans l'explorateur de fichier, en mode admin.)"
            MsgBox(sTxt, MsgBoxStyle.Information, m_sTitreMsg & sTitreMsgDescription)
            If iNbArguments > 0 Then Exit Sub
        End If

        ' Cette dll ne figure pas dans le Framework .NET, elle se trouve ici :
        '  C:\Program Files\Microsoft.NET\Primary Interop Assemblies\adodb.dll
        ' Il faut donc installer les PIA, ou sinon, il suffit de copier la dll
        '  dans le répertoire de l'application
        If iTypeConv = frmXL2Csv.TypeConv.XL2CsvGroup Then
            If Not bFichierExiste(Application.StartupPath & "\ADODB.dll",
                bPrompt:=True) Then Exit Sub
        End If
        If iTypeConv = frmXL2Csv.TypeConv.XL2Csv Or
           iTypeConv = frmXL2Csv.TypeConv.XL2Txt Then
            If Not bFichierExiste(Application.StartupPath & "\ExcelLibrary.dll",
                bPrompt:=True) Then Exit Sub
        End If
        If iTypeConv = frmXL2Csv.TypeConv.XL2CsvNPOI Then
            ' 17/08/2024 NPOI.dll -> NPOI.Core.dll
            'If Not bFichierExiste(Application.StartupPath & "\NPOI.dll", bPrompt:=True) Then Exit Sub
            If Not bFichierExiste(Application.StartupPath & "\NPOI.Core.dll", bPrompt:=True) Then Exit Sub
        End If
        If iTypeConv = frmXL2Csv.TypeConv.XL2CsvSSG Then ' 09/08/2012
            If Not bFichierExiste(Application.StartupPath & "\SpreadSheetGear.dll",
                bPrompt:=True) Then Exit Sub
        End If

        Dim oFrm As New frmXL2Csv
        oFrm.m_sCheminFichierXL = sCheminFichier
        oFrm.m_iTypeConv = iTypeConv
        Application.Run(oFrm)

    End Sub

End Module