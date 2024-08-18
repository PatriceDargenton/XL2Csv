
' Fichier frmXL2Csv.vb
' --------------------

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' a pour Array (tableau) : ()
' o pour Object : objet instancié localement
' refX pour reference à un objet X préexistant qui n'est pas sensé être fermé
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

Imports System.Text ' Pour StringBuilder

Public Class frmXL2Csv

#Region "Configuration"

    ' Mettre False pour éviter de remplacer les , par des . pour les nombres réels
    ' (cela permet de réouvrir correctement les fichiers csv sous Excel avec la ,)
    Private Const bRemplacerSepDec As Boolean = True

    ' Appliquer un TrimEnd = RTrim (enlever les éventuels espaces à la fin des champs textes)
    Private Const bEnleverEspacesFin As Boolean = True

    ' Remplacer les booléens vrai et faux par 1 et vide, respectivement
    Private Const bRemplacerVraiFaux As Boolean = True
    Private Const sValFaux$ = ""
    Private Const sValVrai$ = "1"

#End Region

#Region "Interface"

    Public m_sCheminFichierXL$ = ""
    Public m_iTypeConv As TypeConv

#End Region

#Region "Déclarations"

    Public Enum TypeConv
        XL2Txt      ' Via ExcelLibrary
        XL2CsvGroup ' Toujours via ODBC
        XL2Csv      ' Méthode rapide via ExcelLibrary (et non plus via ODBC)
        XL2CsvNPOI  ' Méthode rapide via NPOI
        XL2CsvSSG   ' Méthode rapide via SpreadSheetGear 09/08/2012
        ' Ancienne méthode via ODBC
        ' (pour classeur Excel correctement formaté : colonnes homogènes) 
        XL2CsvODBC
        XL2CsvAutomation ' Méthode via Automation VBA Excel (lent car via Excel)
    End Enum

    ' Types de conversion
    Public Const sXL2Txt$ = "XL2Txt"           ' Un seul fichier texte via ExcelLibrary
    Public Const sXL2CsvGroup$ = "XL2CsvGroup" ' Un seul fichier csv groupé : fusion csv, via ODBC

    ' Autant de fichiers csv que de feuilles Excel :
    ' --------------------------------------------
    Public Const sXL2Csv$ = "XL2Csv" ' Option par défaut : XL2Csv via ExcelLibrary
    Public Const sXL2CsvNPOI$ = "XL2CsvNPOI" ' XL2Csv via NPOI
    Public Const sXL2CsvSSG$ = "XL2CsvSSG" ' XL2Csv via SpreadSheetGear 09/08/2012
    Public Const sXL2CsvAutomation$ = "XL2CsvAutomation" ' XL2Csv via automation Excel
    Public Const sXL2CsvODBC$ = "XL2CsvODBC"   ' XL2Csv via ODBC
    ' --------------------------------------------

    Private WithEvents m_oODBC As New clsODBC

    Private WithEvents m_msgDelegue As clsMsgDelegue = New clsMsgDelegue

    ' Menus contextuels

    ' 17/09/2017 Excel.Sheet.8 -> *, car sinon ne fonctionne plus ?
    Private Const sMenuCtx_TypeFichierExcel$ = "Excel.Sheet.8"
    Private Const sMenuCtx_TypeFichierTous$ = "*" ' Tous les fichiers
    'Private Const sMenuCtx_TypeFichierSelect$ = sMenuCtx_TypeFichierExcel
    'Private Const sMenuCtx_TypeFichierSelect$ = sMenuCtx_TypeFichierExcel2007
    Private Const sMenuCtx_TypeFichierSelect$ = sMenuCtx_TypeFichierTous

    Private Const sMenuCtx_TypeFichierExcel2007$ = "Excel.Sheet.12"
    ' 03/09/2017 ConvertirEnCsv -> XL2Csv.ConvertirEnCsv (et pour les suivants aussi)
    Private Const sMenuCtx_CleCmdConvertirEnCsv$ = "XL2Csv.ConvertirEnCsv"
    Private Const sMenuCtx_CleCmdConvertirEnCsvDescription$ =
        "Convertir en fichiers Csv (via XLLib.)"
    Private Const sMenuCtx_CleCmdConvertirEnCsvNPOI$ = "XL2Csv.ConvertirEnCsvNPOI"
    Private Const sMenuCtx_CleCmdConvertirEnCsvNPOIDescription$ =
        "Convertir en fichiers Csv (via NPOI)"
    Private Const sMenuCtx_CleCmdConvertirEnCsvSSG$ = "XL2Csv.ConvertirEnCsvSSG" ' 09/08/2012
    Private Const sMenuCtx_CleCmdConvertirEnCsvSSGDescription$ =
        "Convertir en fichiers Csv (via SSG)"
    Private Const sMenuCtx_CleCmdConvertirEn1Csv$ = "XL2Csv.ConvertirEn1Csv" ' XL2CsvGroup
    Private Const sMenuCtx_CleCmdConvertirEn1CsvDescription$ = "Convertir en un fichier Csv fusionné"
    Private Const sMenuCtx_CleCmdConvertirEnTxt$ = "XL2Csv.ConvertirEnTxt"
    Private Const sMenuCtx_CleCmdConvertirEnTxtDescription$ = "Convertir en un fichier Texte"

    Private Const sMenuCtx_CleCmdConvertirEnCsvAutomation$ = "XL2Csv.ConvertirEnCsvAutomation"
    Private Const sMenuCtx_CleCmdConvertirEnCsvAutomationDescription$ =
        "Convertir en fichiers Csv (automation)"

    Private Const sMenuCtx_CleCmdConvertirEnCsvODBC$ = "XL2Csv.ConvertirEnCsvODBC"
    Private Const sMenuCtx_CleCmdConvertirEnCsvODBCDescription$ =
        "Convertir en fichiers Csv (ODBC)"

#End Region

#Region "Initialisations"

    Private Sub frmXL2Csv_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Dim sVersionAppli$ = My.Application.Info.Version.Major &
            "." & My.Application.Info.Version.Minor &
            My.Application.Info.Version.Build
        Dim sVersion$ = " - V" & sVersionAppli & " (" & sDateVersionAppli & ")"
        Dim sDebug$ = " - Debug"
        Dim sTxt$ = Me.Text & sVersion
        If bDebug Then sTxt &= sDebug
        Me.Text = sTxt

        ' 17/09/2017 On ne peut plus associer un menu avec .xls ni .xlsx ?!
        ' On associe tous les fichiers (*) avec XL2Csv alors
        Dim sType$ = sMenuCtx_TypeFichierSelect
        VerifierMenuCtx(sType)

        If Me.m_iTypeConv = TypeConv.XL2CsvODBC Or
           Me.m_iTypeConv = TypeConv.XL2CsvGroup Then
            clsODBC.VerifierConfigODBCExcel()
        End If

        Dim bModeConfig As Boolean = False
        If Me.m_sCheminFichierXL.Length = 0 Then
            bModeConfig = True
        Else
            If Not bFichierExiste(Me.m_sCheminFichierXL,
                bPrompt:=True) Then bModeConfig = True
        End If

        If bModeConfig Then
            Me.cmdAjouterMenuCtx.Visible = True
            Me.cmdEnleverMenuCtx.Visible = True
            Me.chkXL2Csv.Visible = True
            Me.chkXL2CsvNPOI.Visible = True
            If bSpreadSheetGear Then Me.chkXL2CsvSSG.Visible = True
            Me.chkFusionCsv.Visible = True
            Me.chkODBC.Visible = True
            Me.chkTexte.Visible = True
            Me.chkAutomation.Visible = True
            Const sCmd$ = "Ajouter/Retirer le menu contextuel "
            Me.ToolTip1.SetToolTip(Me.chkXL2Csv,
                sCmd & "XL2Csv : Convertir un fichier Excel en fichiers Csv via ExcelLibrary")
            Me.ToolTip1.SetToolTip(Me.chkXL2CsvNPOI,
                sCmd & "XL2CsvNPOI : Convertir un fichier Excel en fichiers Csv via NPOI")
            Me.ToolTip1.SetToolTip(Me.chkXL2CsvSSG,
                sCmd & "XL2CsvSSG : Convertir un fichier Excel en fichiers Csv via SpreadSheetGear") ' 09/08/2012
            Me.ToolTip1.SetToolTip(Me.chkFusionCsv,
                sCmd & "XL2CsvGroup : Convertir un fichier Excel en un fichier Csv via ODBC")
            Me.ToolTip1.SetToolTip(Me.chkAutomation,
                sCmd & "XL2CsvAutomation : Convertir un fichier Excel en fichiers Csv via Automation Excel")
            Me.ToolTip1.SetToolTip(Me.chkODBC,
                sCmd & "XL2CsvODBC : Convertir un fichier Excel en fichiers Csv via ODBC")
            Me.ToolTip1.SetToolTip(Me.chkTexte,
                sCmd & "XL2Txt : Convertir un fichier Excel en un fichier Texte via ExcelLibrary")
            Me.cmdConv.Visible = False
            Me.cmdAnnuler.Visible = False
            Exit Sub
        End If

        Select Case Me.m_iTypeConv
            Case TypeConv.XL2CsvGroup : Me.Text =
            "XL2CsvGroup" & sVersion & " : Convertir un fichier Excel en fichier Csv via ODBC"
            Case TypeConv.XL2Csv : Me.Text =
            "XL2Csv" & sVersion & " : Convertir un fichier Excel en fichiers Csv via ExcelLibrary"
            Case TypeConv.XL2CsvNPOI : Me.Text =
            "XL2Csv" & sVersion & " : Convertir un fichier Excel en fichiers Csv via NPOI"
            Case TypeConv.XL2CsvSSG : Me.Text =
            "XL2Csv" & sVersion & " : Convertir un fichier Excel en fichiers Csv via SpreaSheetGear" ' 09/08/2012
            Case TypeConv.XL2CsvAutomation : Me.Text =
            "XL2Csv" & sVersion & " : Convertir un fichier Excel en fichiers Csv via automation Excel"
            Case TypeConv.XL2CsvODBC : Me.Text =
            "XL2Csv" & sVersion & " : Convertir un fichier Excel en fichiers Csv via ODBC"
            Case TypeConv.XL2Txt : Me.Text =
            "XL2Txt" & sVersion & " : Convertir un fichier Excel en fichier Texte via ExcelLibrary"
        End Select
        If bDebug Then Me.Text &= sDebug

        Me.cmdAjouterMenuCtx.Visible = False
        Me.cmdEnleverMenuCtx.Visible = False

        Me.cmdConv.Visible = True
        Me.cmdAnnuler.Visible = True

        Me.cmdAnnuler.Enabled = True
        Me.cmdConv.Enabled = False
        Dim bOk As Boolean

        If Me.m_iTypeConv = TypeConv.XL2Csv Then
            bOk = bConvertirXLRapide(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        ElseIf Me.m_iTypeConv = TypeConv.XL2CsvNPOI Then
            bOk = bConvertirXLRapideNPOI(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        ElseIf Me.m_iTypeConv = TypeConv.XL2CsvSSG Then
            bOk = bConvertirXLRapideSSG(Me.m_sCheminFichierXL, Me.m_msgDelegue) ' 09/08/2012
        ElseIf Me.m_iTypeConv = TypeConv.XL2CsvAutomation Then
            bOk = bConvertirXLAutomation(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        ElseIf Me.m_iTypeConv = TypeConv.XL2Txt Then
            bOk = bConvertirXL2Txt(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        Else
            bOk = bConvertirXLODBC(bAuto:=True)
        End If

        Me.cmdConv.Enabled = True
        Me.cmdAnnuler.Enabled = False
        If bOk And bRelease Then Me.Close()

    End Sub

    Private Sub AfficherMessage(sMsg$)
        Me.sbStatusBar.Text = sMsg
        Application.DoEvents()
    End Sub

    Private Sub m_oODBC_EvAfficherMessage(sender As Object, e As clsMsgEventArgs) _
            Handles m_oODBC.EvAfficherMessage
        AfficherMessage(e.sMessage)
    End Sub

    Private Sub chkODBC_CheckedChanged(sender As Object, e As EventArgs) _
            Handles chkODBC.CheckedChanged

        If Me.chkODBC.Checked Then
            clsODBC.VerifierConfigODBCExcel()
        End If

    End Sub

    Private Sub AfficherMessageEv(sender As Object, e As clsMsgEventArgs) _
            Handles m_msgDelegue.EvAfficherMessage

        Me.AfficherMessage(e.sMessage)

        ' Autre solution :
        'AddHandler glb.msgDelegue.EvAfficherMessage, _
        '    AddressOf AfficherMessageEv

    End Sub

#End Region

#Region "Conversion"

    Private Sub cmdConv_Click(sender As Object, e As EventArgs) Handles cmdConv.Click

        Me.cmdAnnuler.Enabled = True
        Me.cmdConv.Enabled = False

        If Me.m_iTypeConv = TypeConv.XL2CsvAutomation Then
            bConvertirXLAutomation(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        ElseIf Me.m_iTypeConv = TypeConv.XL2Csv Then
            bConvertirXLRapide(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        ElseIf Me.m_iTypeConv = TypeConv.XL2CsvNPOI Then
            bConvertirXLRapideNPOI(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        ElseIf Me.m_iTypeConv = TypeConv.XL2CsvSSG Then
            bConvertirXLRapideSSG(Me.m_sCheminFichierXL, Me.m_msgDelegue) ' 09/08/2012
        ElseIf Me.m_iTypeConv = TypeConv.XL2Txt Then
            bConvertirXL2Txt(Me.m_sCheminFichierXL, Me.m_msgDelegue)
        Else
            bConvertirXLODBC(bAuto:=False)
        End If

        Me.cmdConv.Enabled = True
        Me.cmdAnnuler.Enabled = False

    End Sub

    Private Sub cmdAnnuler_Click(sender As Object, e As EventArgs) Handles cmdAnnuler.Click
        Me.m_oODBC.Annuler()
        Me.m_msgDelegue.m_bAnnuler = True
    End Sub

    Private Function bConvertirXLODBC(bAuto As Boolean) As Boolean

        ' 16/04/2011 Il faut vraiment vérifier, sinon c'est trop long !
        ' 11/03/2012 bEcriture:=False : Non ! si on veut aller vite
        '  ne pas permettre qu'Excel soit ouvert, même ici
        If Not bFichierAccessibleMultiTest(Me.m_sCheminFichierXL, Me.m_msgDelegue) Then _
            Return False

        Me.m_oODBC.m_bAfficherMsg = False

        Me.m_oODBC.m_sChaineConnexionDirecte =
            "Provider=Microsoft.Jet.OLEDB.4.0;" &
            "Data Source=" & Me.m_sCheminFichierXL & ";" &
            "Extended Properties=""Excel 8.0;"";"

        AfficherMessage("Analyse du fichier Excel " &
            IO.Path.GetFileName(Me.m_sCheminFichierXL) & " en cours...")

        Me.m_oODBC.LibererRessources()
        Me.m_oODBC.m_bPrompt = False
        Me.m_oODBC.m_bVerifierConfigODBCExcel = False ' Déjà fait une fois
        Me.m_oODBC.m_bCopierDonneesPressePapier = False ' Sauf en mode Debug
        Me.m_oODBC.m_bLireToutDUnBloc = True

        ' Mettre False pour éviter de remplacer les , par des . pour les nombres réels
        ' (cela permet de réouvrir correctement les fichiers csv sous Excel avec la ,)
        Me.m_oODBC.m_bRemplacerSepDec = bRemplacerSepDec

        Me.m_oODBC.m_bEnleverEspacesFin = bEnleverEspacesFin ' Appliquer un TrimEnd = RTrim

        Me.m_oODBC.m_bRemplacerVraiFaux = bRemplacerVraiFaux
        Me.m_oODBC.m_sValFaux = sValFaux
        Me.m_oODBC.m_sValVrai = sValVrai

        If Not Me.m_oODBC.bExplorerSourceODBC(
            bExplorerChamps:=True, bRenvoyerContenu:=bDebug) Then Return False
        Dim sTable$
        Dim iNumTable% = 0
        Dim sbContenu As New StringBuilder
        ' Mémoriser les champs
        Dim asChamps$(,) = DirectCast(Me.m_oODBC.m_asChamps.Clone, String(,))
        Dim iNbChpsMax% = asChamps.GetUpperBound(1)
        ' Dimension : 1 : NbTables, 2 : NbChamps
        Dim iNbChamps% = UBound(asChamps, 2) + 1

        If Me.m_iTypeConv = TypeConv.XL2CsvGroup Then
            ' Mode tableau : mettre ligne entete comme premier classeur
            ' D'abord le nom de la table
            sbContenu.Append("Table;")
            AjouterEnteteTable(sbContenu, Me.m_oODBC.m_iNumTableMaxChamps, asChamps, iNbChamps)
        ElseIf Me.m_iTypeConv = TypeConv.XL2Txt Then
            sbContenu.Append("Fichier source : " & Me.m_sCheminFichierXL & vbCrLf)
            Dim fi As New IO.FileInfo(Me.m_sCheminFichierXL)
            Dim lTailleFichier& = fi.Length
            Dim sTailleFichier$ = sFormaterTailleOctets(lTailleFichier)
            ' fi.LastWriteTime affiche toujours la bonne heure (et la même heure)
            sbContenu.Append("Taille : " & sTailleFichier &
                ", Date : " & fi.LastWriteTime & vbCrLf & vbCrLf)
        End If

        ' Traiter d'abord les tables ayant le plus de champs afin d'avoir toutes les entetes
        ' 08/08/2012 Ssi 2 passes ! (XL2CsvGroup)
        Dim iNbPasses% = 1
        Dim b2Passes As Boolean = False
        If Me.m_iTypeConv = TypeConv.XL2CsvGroup Then b2Passes = True : iNbPasses = 2

        Dim sCheminFichier$ = ""
        Dim iNbTables% = Me.m_oODBC.m_alTables.Count
        Dim iPasse%
        For iPasse = 0 To iNbPasses - 1 ' 2 Passes
            iNumTable = 0
            For Each sTable In Me.m_oODBC.m_alTables

                If b2Passes Then
                    Dim bTableMax As Boolean = False
                    If Me.m_oODBC.m_aiNbChamps(iNumTable) = iNbChpsMax Then bTableMax = True
                    If iPasse = 0 And Not bTableMax Then GoTo TableSuivante
                    If iPasse = 1 And bTableMax Then GoTo TableSuivante
                End If

                AfficherMessage("Lecture de la feuille [" & sTable.Replace("$", "") &
                    "] en cours... (" & iNumTable + 1 & "/" & iNbTables & ")")
                If Me.m_oODBC.bAnnuler Then Exit For

                Me.m_oODBC.m_sListeSQL = "Select * From [" & sTable & "]"
                If Not Me.m_oODBC.bLireSourceODBC(bRenvoyerContenu:=bDebug,
                    bNePasFermerConnexion:=True) Then Return False

                ' Analyse du ou des tableaux résultats

                ' Enlever le $ à la fin de la table
                Dim sNomTable$ = sNomTableExcel(sTable)

                If Me.m_iTypeConv = TypeConv.XL2Txt Then
                    sbContenu.Append("Table : [" & sNomTable & "]" & vbCrLf & vbCrLf)
                ElseIf Me.m_iTypeConv = TypeConv.XL2CsvGroup Then
                    ' Rien à faire
                Else
                    sbContenu = New StringBuilder
                End If

                Dim asTableau$(,) = CType(Me.m_oODBC.m_aoMetaTableau(0), String(,))
                If IsNothing(asTableau) Then
                    ' 21/02/2009 Aucun enregistrement
                    Dim iNbColonnes0% = asChamps.GetUpperBound(1) + 1
                    ' Feuille vide et sans entête (sauf si l'entête est justement F1 !)
                    If iNbColonnes0 = 1 And asChamps(0, 0) = "F1" Then GoTo TableSuivante
                    ' Feuille vide et avec entête
                    If Me.m_iTypeConv = TypeConv.XL2Txt Or
                       Me.m_iTypeConv = TypeConv.XL2CsvODBC Then
                        AjouterEnteteTable(sbContenu, iNumTable, asChamps, iNbColonnes0)
                    End If
                    If Me.m_iTypeConv = TypeConv.XL2CsvODBC Then
                        Dim sCheminFichier0$ = IO.Path.GetDirectoryName(
                            Me.m_sCheminFichierXL) & "\" & sNomTable & ".csv"
                        ' 03/09/2017 Encodage UTF8
                        If Not bEcrireFichier(sCheminFichier0, sbContenu, bEncodageUTF8:=True) Then GoTo Erreur
                        sCheminFichier = sCheminFichier0
                    End If
                    GoTo TableSuivante
                End If

                If Me.m_iTypeConv = TypeConv.XL2Txt Then _
                    sbContenu.Append("Table;")

                Dim iNbColonnes% = asTableau.GetUpperBound(0) + 1
                If Me.m_iTypeConv = TypeConv.XL2Txt Or
                   Me.m_iTypeConv = TypeConv.XL2CsvODBC Then
                    AjouterEnteteTable(sbContenu, iNumTable, asChamps, iNbColonnes)
                End If

                Dim iNbLignes% = asTableau.GetUpperBound(1) + 1
                Dim i%, j%
                For j = 0 To iNbLignes - 1
                    If Me.m_iTypeConv = TypeConv.XL2CsvGroup Or
                       Me.m_iTypeConv = TypeConv.XL2Txt Then _
                        sbContenu.Append(sNomTable & ";")
                    For i = 0 To iNbColonnes - 1
                        Dim sVal$ = asTableau(i, j)
                        ' Rq sans enreg. : 1 ligne en fait
                        If IsNothing(sVal) Then Exit For
                        sbContenu.Append(sVal)
                        ' Ne pas ajouter le dernier ; pour faire comme Excel
                        If i < iNbColonnes - 1 Then sbContenu.Append(";")
                    Next i
                    sbContenu.Append(vbCrLf)
                Next j
                If Me.m_iTypeConv = TypeConv.XL2Txt Then
                    sbContenu.Append(vbCrLf & vbCrLf)
                ElseIf Me.m_iTypeConv = TypeConv.XL2CsvGroup Then
                    ' Ne rien ajouter
                Else
                    Dim sCheminFichier0$ = IO.Path.GetDirectoryName(
                        Me.m_sCheminFichierXL) & "\" & sNomTable & ".csv"
                    ' 03/09/2017 Encodage UTF8
                    If Not bEcrireFichier(sCheminFichier0, sbContenu, bEncodageUTF8:=True) Then GoTo Erreur
                    sCheminFichier = sCheminFichier0 ' 15/12/2007 Le noter s'il n'y en a qu'un
                End If

TableSuivante:
                iNumTable += 1
            Next sTable
        Next iPasse

        If bDebug And Not IsNothing(Me.m_oODBC.m_sbContenuRetour) Then _
            CopierPressePapier(Me.m_oODBC.m_sbContenuRetour.ToString)

        Me.m_oODBC.LibererRessources()
        AfficherMessage("Opération terminée.")

        If Me.m_oODBC.bAnnuler Then Return False

        Dim sTypeConv$ = "en fichiers csv"
        Dim sExt$ = ".csv"
        If Me.m_iTypeConv = TypeConv.XL2CsvGroup Then
            sTypeConv = "en fichier csv"
        ElseIf Me.m_iTypeConv = TypeConv.XL2Txt Then
            sTypeConv = "en fichier texte"
            sExt = ".txt"
        End If

        ' 15/12/2007 Non car déjà fait : Or (Me.m_iTypeConv = TypeConv.XL2CsvODBC And iNumTable = 1)
        If Me.m_iTypeConv = TypeConv.XL2CsvGroup Or
           Me.m_iTypeConv = TypeConv.XL2Txt Then
            sCheminFichier = IO.Path.GetDirectoryName(
                Me.m_sCheminFichierXL) & "\" &
                IO.Path.GetFileNameWithoutExtension(
                Me.m_sCheminFichierXL) & sExt
            ' 03/09/2017 Encodage UTF8
            If Not bEcrireFichier(sCheminFichier, sbContenu, bEncodageUTF8:=True) Then GoTo Erreur
        End If

        bConvertirXLODBC = True
        Me.cmdAnnuler.Enabled = False
        If Me.m_iTypeConv = TypeConv.XL2CsvODBC And iNumTable > 1 Then
            ' Plusieurs fichiers possibles
            Dim sInfo$ = "Le classeur :" & vbCrLf & Me.m_sCheminFichierXL & vbCrLf &
                "a été converti " & sTypeConv & " avec succès !" & vbCrLf & "(via ODBC)"
            MsgBox(sInfo, MsgBoxStyle.Information, m_sTitreMsg)
        Else
            If sCheminFichier.Length = 0 Then ' 21/02/2009
                Dim sInfo$ = "Le classeur est vide !" & vbCrLf & Me.m_sCheminFichierXL
                MsgBox(sInfo, MsgBoxStyle.Information, m_sTitreMsg)
            Else
                ProposerOuvrirFichier(sCheminFichier)
            End If
        End If
        Exit Function

Erreur:
        AfficherMessage("Erreur !")
        Return False

    End Function

    Private Sub AjouterEnteteTable(ByRef sbContenu As StringBuilder,
            iNumTable%, asChamps$(,), iNbChamps%)

        Dim i%
        ' Dimension : 1 : NbTables, 2 : NbChamps
        'Dim iNbChamps% = UBound(asChamps, 2) + 1
        For i = 0 To iNbChamps - 1
            Dim sChamp$ = asChamps(iNumTable, i)
            sbContenu.Append(sChamp)
            ' Ne pas ajouter le dernier ; pour faire comme Excel
            If i < iNbChamps - 1 Then sbContenu.Append(";")
        Next i
        sbContenu.Append(vbCrLf)

    End Sub

    Private Function sNomTableExcel$(sTable$)

        ' Enlever le $ à la fin de la table
        Dim iLen% = sTable.Length
        sNomTableExcel = sTable.Substring(0, iLen - 1)
        If sTable.Chars(0) = "'"c And
            sTable.Chars(iLen - 1) = "'"c Then
            sNomTableExcel = sTable.Substring(1, iLen - 3)
        End If

        ' 12/03/2010
        sNomTableExcel = sConvNomDos(sNomTableExcel)

    End Function

#End Region

#Region "Gestion des menus contextuels"

    Private Sub cmdAjouterMenuCtx_Click(sender As Object, e As EventArgs) _
        Handles cmdAjouterMenuCtx.Click

        ' On ne peut plus associer un menu avec .xls ni .xlsx ?!
        ' On associe tous les fichiers (*) avec XL2Csv alors
        Dim sType$ = sMenuCtx_TypeFichierSelect
        AjouterMenuCtx(sType)
        If sType <> sMenuCtx_TypeFichierTous Then
            'If bExcel2007SupportNPOI Then _
            If bExcel2007SupportSSG AndAlso Me.chkXL2CsvSSG.Checked Then _
                AjouterMenuCtx(sMenuCtx_TypeFichierExcel2007)
        End If

    End Sub

    Private Sub cmdEnleverMenuCtx_Click(sender As Object, e As EventArgs) _
            Handles cmdEnleverMenuCtx.Click

        Dim sType$ = sMenuCtx_TypeFichierSelect
        EnleverMenuCtx(sType)
        If sType <> sMenuCtx_TypeFichierTous Then
            'If bExcel2007SupportNPOI Then _
            If bExcel2007SupportSSG AndAlso Me.chkXL2CsvSSG.Checked Then _
                EnleverMenuCtx(sMenuCtx_TypeFichierExcel2007)
        End If

    End Sub

    Private Sub AjouterMenuCtx(sTypeFichierExcel$)
        'Optional bExcel2007 As Boolean = False)

        Dim sCheminExe$ = Application.ExecutablePath
        Const bPrompt As Boolean = False
        Const sChemin$ = """%1"""

        ' Ajouter un pointeur HKCR\.xls vers HKCR\XL2Csv
        'bAjouterTypeFichier(sMenuCtx_ExtFichierIdx, sMenuCtx_TypeFichierIdx, _
        '    sMenuCtx_ExtFichierIdxDescription)

        If Me.chkXL2CsvSSG.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvSSG,
            bPrompt, , sMenuCtx_CleCmdConvertirEnCsvSSGDescription, sCheminExe,
            sChemin & " " & sXL2CsvSSG)

        'If bExcel2007 Then Exit Sub

        If Me.chkXL2CsvNPOI.Checked Then _
            bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvNPOI,
                bPrompt, , sMenuCtx_CleCmdConvertirEnCsvNPOIDescription, sCheminExe,
                sChemin & " " & sXL2CsvNPOI)

        If Me.chkXL2Csv.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsv,
            bPrompt, , sMenuCtx_CleCmdConvertirEnCsvDescription, sCheminExe,
            sChemin) ' & " " & sXL2Csv : Par défaut

        If Me.chkFusionCsv.Checked Then
            If bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEn1Csv,
                bPrompt, , sMenuCtx_CleCmdConvertirEn1CsvDescription, sCheminExe,
                sChemin & " " & sXL2CsvGroup) Then
                clsODBC.VerifierConfigODBCExcel() ' 07/01/2012 V1.08
            End If
        End If

        If Me.chkTexte.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnTxt,
            bPrompt, , sMenuCtx_CleCmdConvertirEnTxtDescription, sCheminExe,
            sChemin & " " & sXL2Txt)

        If Me.chkODBC.Checked Then
            If bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvODBC,
                bPrompt, , sMenuCtx_CleCmdConvertirEnCsvODBCDescription, sCheminExe,
                sChemin & " " & sXL2CsvODBC) Then
                clsODBC.VerifierConfigODBCExcel() ' 07/01/2012 V1.08
            End If
        End If

        If Me.chkAutomation.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvAutomation,
            bPrompt, , sMenuCtx_CleCmdConvertirEnCsvAutomationDescription, sCheminExe,
            sChemin & " " & sXL2CsvAutomation)

        VerifierMenuCtx(sTypeFichierExcel)

    End Sub

    Private Sub EnleverMenuCtx(sTypeFichierExcel$)
        'Optional bExcel2007 As Boolean = False)

        If Me.chkXL2CsvSSG.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvSSG,
            bEnlever:=True, bPrompt:=False)

        'If bExcel2007 Then Exit Sub

        If Me.chkXL2CsvNPOI.Checked Then _
            bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvNPOI,
                bEnlever:=True, bPrompt:=False)

        If Me.chkXL2Csv.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsv,
            bEnlever:=True, bPrompt:=False)

        If Me.chkFusionCsv.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEn1Csv,
            bEnlever:=True, bPrompt:=False)
        If Me.chkTexte.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnTxt,
            bEnlever:=True, bPrompt:=False)
        If Me.chkODBC.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvODBC,
            bEnlever:=True, bPrompt:=False)
        If Me.chkAutomation.Checked Then _
        bAjouterMenuContextuel(sTypeFichierExcel, sMenuCtx_CleCmdConvertirEnCsvAutomation,
            bEnlever:=True, bPrompt:=False)

        VerifierMenuCtx(sTypeFichierExcel)

    End Sub

    Private Sub VerifierMenuCtx(sTypeFichierExcel$)

        Const sShell$ = "\shell\"

        Dim sCleDescriptionCmd$ = sTypeFichierExcel & sShell &
            sMenuCtx_CleCmdConvertirEnCsv
        Dim bCleXL2Csv As Boolean = bCleRegistreCRExiste(sCleDescriptionCmd)

        Dim sCleDescriptionCmdNPOI$ = sTypeFichierExcel & sShell &
            sMenuCtx_CleCmdConvertirEnCsvNPOI
        Dim bCleXL2CsvNPOI As Boolean = bCleRegistreCRExiste(sCleDescriptionCmdNPOI)

        Dim sCleDescriptionCmdSSG$ = sTypeFichierExcel & sShell &
            sMenuCtx_CleCmdConvertirEnCsvSSG
        Dim bCleXL2CsvSSG As Boolean = bCleRegistreCRExiste(sCleDescriptionCmdSSG)

        Dim sCleDescriptionCmdFusion$ = sTypeFichierExcel & sShell &
            sMenuCtx_CleCmdConvertirEn1Csv
        Dim bCleFusion As Boolean = bCleRegistreCRExiste(sCleDescriptionCmdFusion)

        Dim sCleDescriptionCmdAutomation$ = sTypeFichierExcel & sShell &
            sMenuCtx_CleCmdConvertirEnCsvAutomation
        Dim bCleAutom As Boolean = bCleRegistreCRExiste(sCleDescriptionCmdAutomation)

        Dim sCleDescriptionCmdODBC$ = sTypeFichierExcel & sShell &
            sMenuCtx_CleCmdConvertirEnCsvODBC
        Dim bCleODBC As Boolean = bCleRegistreCRExiste(sCleDescriptionCmdODBC)

        Dim sCleDescriptionCmdTxt$ = sTypeFichierExcel & sShell &
            sMenuCtx_CleCmdConvertirEnTxt
        Dim bCleTxt As Boolean = bCleRegistreCRExiste(sCleDescriptionCmdTxt)

        If bCleXL2Csv OrElse bCleXL2CsvNPOI OrElse bCleXL2CsvSSG OrElse
           bCleFusion OrElse bCleAutom OrElse bCleODBC OrElse bCleTxt Then

            Me.cmdAjouterMenuCtx.Enabled = False
            Me.cmdEnleverMenuCtx.Enabled = True

            Me.chkXL2Csv.Checked = bCleXL2Csv
            Me.chkXL2CsvNPOI.Checked = bCleXL2CsvNPOI
            Me.chkXL2CsvSSG.Checked = bCleXL2CsvSSG
            Me.chkFusionCsv.Checked = bCleFusion
            Me.chkAutomation.Checked = bCleAutom
            Me.chkODBC.Checked = bCleODBC
            Me.chkTexte.Checked = bCleTxt

            ' Interdire de décocher
            Me.chkXL2Csv.Enabled = False
            Me.chkXL2CsvNPOI.Enabled = False
            Me.chkXL2CsvSSG.Enabled = False
            Me.chkFusionCsv.Enabled = False
            Me.chkAutomation.Enabled = False
            Me.chkODBC.Enabled = False
            Me.chkTexte.Enabled = False

        Else

            Me.cmdAjouterMenuCtx.Enabled = True
            Me.cmdEnleverMenuCtx.Enabled = False

            ' Autoriser à cocher
            Me.chkXL2Csv.Enabled = True
            Me.chkXL2CsvNPOI.Enabled = True
            Me.chkXL2CsvSSG.Enabled = True
            Me.chkFusionCsv.Enabled = True
            Me.chkODBC.Enabled = True
            Me.chkTexte.Enabled = True
            Me.chkAutomation.Enabled = True

        End If

    End Sub

#End Region

End Class