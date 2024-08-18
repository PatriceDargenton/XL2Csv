
' Gestion Excel en liaison tardive

Option Strict Off
Option Infer On

Imports System.Text ' Pour StringBuilder

Module modExcel

    Dim oXLH As clsHebExcel = Nothing

#Region "Constantes"

    Public Const sMsgLancementExcel$ = "Lancement d'Excel..."
    Public Const sMsgPbExcel$ = "Excel n'est pas installé !"

    Private Const iIdxCouleurAuto% = -4105 ' xlAutomatic
    Private Const iIdxCoulGrise% = 56 ' Gris très clair (couleur perso)
    Private Const xlNone% = -4142 '(&HFFFFEFD2)

    Private Const iBordGauche% = 7 ' xlEdgeLeft
    Private Const iBordHaut% = 8 ' xlEdgeTop
    Private Const iBordBas% = 9 ' xlEdgeBottom
    Private Const iBordDroite% = 10 ' xlEdgeRight

    Private Const sErrExcelManip$ = "0x800A01A8"

    'Private Const sMsgErr$ = "Impossible d'exporter le document sous Excel !"
    'Private Const sMsgErrExcel$ = "Impossible de créer le document sous Excel !"
    Public Const sMsgErrCausePoss$ =
        "Cause possible : Excel est actuellement en cours d'édition d'un document"

#End Region

#Region "XL2Csv"

    'Excel.xlCVErr  Range.Value  Coerced to .NET
    '-------------  -----------  ---------------
    '    2000 00      #NULL!       -2146826288
    '    2007 07      #DIV/0!      -2146826281
    '    2015 0F      #VALUE!      -2146826273
    '    2023 17      #REF!        -2146826265
    '    2029 1D      #NAME?       -2146826259
    '    2036 24      #NUM!        -2146826252
    '    2042 2A      #N/A!        -2146826246
    Private Const sValErrNULL$ = "-2146826288"
    Private Const sValErrDIV0$ = "-2146826281"
    Private Const sValErrVALUE$ = "-2146826273"
    Private Const sValErrREF$ = "-2146826265"
    Private Const sValErrNAME$ = "-2146826259"
    Private Const sValErrNUM$ = "-2146826252"
    Private Const sValErrNA$ = "-2146826246"
    Private Const sErrNULL$ = "#NULL!"
    Private Const sErrDIV0$ = "#DIV/0!"
    Private Const sErrVALUE$ = "#VALUE!"
    Private Const sErrREF$ = "#REF!"
    Private Const sErrNAME$ = "#NAME?"
    Private Const sErrNUM$ = "#NUM!"
    Private Const sErrNA$ = "#N/A" ' Pas de ! pour celui-là !

    Public Function bConvertirXLAutomation(sCheminFichierXL$, msgDelegue As clsMsgDelegue) As Boolean

        ' Convertir un classeur Excel en fichiers csv de manière sécurisée
        '  (cellule par cellule, ce qui est plus lent que par ODBC)

        bConvertirXLAutomation = False

        If Not bFichierAccessibleMultiTest(sCheminFichierXL, msgDelegue) Then _
            Return False

        Dim sCheminDossierXL$ = IO.Path.GetDirectoryName(sCheminFichierXL)

        ' Optimisation : conserver la précédante instance d'Excel
        'Dim oXLH As clsHebExcel
        Dim oWkb As Object = Nothing
        Dim oSht As Object = Nothing

        msgDelegue.m_bAnnuler = False

        Try

            Sablier()
            msgDelegue.AfficherMsg(sMsgLancementExcel)
            If IsNothing(oXLH) Then oXLH = New clsHebExcel(bInterdireAppliAvant:=False)
            If IsNothing(oXLH.oXL) Then
                msgDelegue.AfficherMsg(sMsgPbExcel)
                GoTo Fin
            End If

            'oXLH.oXL.Visible = True ' Mode Debug
            oXLH.oXL.Visible = False

            msgDelegue.AfficherMsg(sMsgOuvertureClasseur)
            oWkb = oXLH.oXL.Workbooks.Open(sCheminFichierXL)

            Dim iFeuille% = 0
            Dim iNbFeuilles% = oWkb.Worksheets.Count
            For Each oSht In oWkb.Worksheets
                iFeuille += 1
                'Dim sPlageUt$ = oSht.UsedRange.Address.ToString
                'Debug.WriteLine(oSht.Name & " : " & sPlageUt)
                Dim sFeuille$ = oSht.Name
                Dim sFeuilleDos$ = sConvNomDos(sFeuille)
                Dim sb As New StringBuilder
                Dim sbTmp2 As New StringBuilder
                Dim iLongUtile2% = -1
                Dim iLigne% = 1

                Dim iNbCol2% = oSht.UsedRange.Columns.Count
                Dim iNbLignes2% = oSht.UsedRange.Rows.Count

                ' Il faut plutot rechercher l'indice max.
                '  car si la 1ère ligne est vide on va manquer une ligne à la fin
                ' (pour les colonnes, pas de pb à priori)
                Dim sPlage$ = oSht.UsedRange.Address.ToString.Replace("$", "")

                'Dim iColMin% = iColPlage(sPlage)
                'Dim iLigneMin% = iLignePlage(sPlage)
                'iNbCol += iColMin - 1
                'iNbLignes += iLigneMin - 1
                Dim iNbCol% = iColFinPlage(sPlage)
                Dim iNbLignes% = iLigneFinPlage(sPlage)
                ' Mais parfois, la plage n'indique pas les colonnes, seulement les lignes !
                '  on prend le max. dans ce cas
                If iNbCol < iNbCol2 Then iNbCol = iNbCol2
                If iNbLignes < iNbLignes2 Then iNbLignes = iNbLignes2
                ' N'existent pas :
                'Dim iNbCol2% = oSht.Cells.LastColIndex + 1
                'Dim iNbLignes2% = oSht.Cells.LastRowIndex + 1

                Dim bAuMoinsUneVal As Boolean = False
                'For Each oRow As Object In oSht.UsedRange.Rows
                ' Toujours partir de 0, et non depuis le début utilisé (+ stable)
                For i As Integer = 0 To iNbLignes - 1

                    If iLigne = 1 Or iLigne = iNbLignes Or iLigne Mod 10 = 0 Then
                        msgDelegue.AfficherMsg(
                            "Feuille n°" & iFeuille & "/" & iNbFeuilles &
                            " : Ligne n°" & iLigne & "/" & iNbLignes &
                            " : Lecture en cours...")
                        If msgDelegue.m_bAnnuler Then Return False
                    End If

                    Dim sbTmp As New StringBuilder
                    Dim bAuMoinsUneValLigne As Boolean = False
                    Dim iLongUtile% = -1
                    Dim iCol% = 1
                    'For Each oCol As Object In oSht.UsedRange.Columns
                    For j As Integer = 0 To iNbCol - 1

                        Dim oVal As Object = oSht.Cells(iLigne, iCol).Value
                        Dim sVal$ = ""
                        If IsNothing(oVal) Then GoTo ColonneSuivante

                        ' Note : Si la cellule est au format monétaire,
                        '  alors la précision lue ici est tronquée à 4 décimales
                        '  (ce qui n'est pas le cas via ExcelLibrary :
                        '   le format monétaire ne peut être détecté qu'en analysant
                        '   le format d'affichage, car le type est System.Double)
                        ' Note : La méthode "Enregistrer sous csv" d'Excel
                        '  tient compte du format : pas ici
                        sVal = oVal.ToString
                        Dim bVal As Boolean = False
                        If sVal.Length > 0 Then
                            bAuMoinsUneValLigne = True : bAuMoinsUneVal = True
                            bVal = True
                            If String.Compare(sVal, sValErrNULL) = 0 Then
                                sVal = sErrNULL
                            ElseIf String.Compare(sVal, sValErrDIV0) = 0 Then
                                sVal = sErrDIV0
                            ElseIf String.Compare(sVal, sValErrVALUE) = 0 Then
                                sVal = sErrVALUE
                            ElseIf String.Compare(sVal, sValErrREF) = 0 Then
                                sVal = sErrREF
                            ElseIf String.Compare(sVal, sValErrNAME) = 0 Then
                                sVal = sErrNAME
                            ElseIf String.Compare(sVal, sValErrNUM) = 0 Then
                                sVal = sErrNUM
                            ElseIf String.Compare(sVal, sValErrNA) = 0 Then
                                sVal = sErrNA
                            Else
                                Dim sType = oVal.GetType().Name
                                If sType = "Decimal" Then
                                    Dim oDec As System.Decimal = CDec(oVal)
                                    sVal = sFormaterNumeriqueDec(oDec,
                                        iNbDecimales:=iNbDecimalesDef,
                                        sSeparateurMilliers:=sSeparateurMilliersDef)
                                    'Dim oDbl As Double = CDbl(oDec)
                                    'sVal = sFormaterNumeriqueDble(oDbl, sSeparateurMilliers:="")
                                ElseIf sType = "Double" Then
                                    Dim oDbl As Double = CDbl(oVal)
                                    sVal = sFormaterNumeriqueDble(oDbl,
                                        iNbDecimales:=iNbDecimalesDef,
                                        sSeparateurMilliers:=sSeparateurMilliersDef)
                                ElseIf sType = "DateTime" Then
                                    If sVal.EndsWith(sHeureVide) Then
                                        sVal = Left$(sVal, sVal.Length - sHeureVide.Length)
                                    End If
                                End If
                                ' Si réel alors appliquer le format
                                'oSht.Cells(iLigne, iCol).Select()
                                'Dim sFormat$ = oXLH.oXL.Selection.NumberFormat()
                                'Dim sAddresse$ = oSht.Cells(iLigne, iCol).Address
                                'Dim sFormat$ = oXLH.oXL.Range(sAddresse).NumberFormat()
                                'Dim oDec As System.Decimal = CDec(oVal)
                                'sVal = oDec.ToString(sFormat)
                                'Dim oDbl As Double = CDbl(oDec)
                                'sVal = sFormaterNumeriqueDble(oDbl)
                            End If
                        End If
                        sbTmp.Append(sVal)
                        If bVal Then iLongUtile = sbTmp.Length

ColonneSuivante:
                        'Debug.WriteLine("L" & iLigne & "C" & iCol & " : " & sVal)
                        If iCol < iNbCol Then sbTmp.Append(";")
                        iCol += 1
                    Next
                    If bAuMoinsUneValLigne Then
                        ' Retirer les ; à la fin
                        sbTmp.Length = iLongUtile
                        sbTmp2.Append(sbTmp)
                        iLongUtile2 = sbTmp2.Length
                    End If

                    sbTmp2.Append(vbCrLf)
                    iLigne += 1
                Next
                If Not bAuMoinsUneVal Then Continue For
                Dim sChemin$ = sCheminDossierXL & "\" & sFeuilleDos & ".csv"
                ' Limiter le sb à la taille utile (supprimer les lignes vides à la fin)
                sb.Append(sbTmp2)
                sb.Length = iLongUtile2 + 2 ' +2 pour vbCrLf
                ' 03/09/2017 Encodage UTF8
                If Not bEcrireFichier(sChemin, sb, bEncodageUTF8:=True) Then Return False
            Next

            bConvertirXLAutomation = True

            Dim sInfo$ = "Le classeur :" & vbCrLf & sCheminFichierXL & vbCrLf &
                "a été converti en fichiers csv avec succès ! (via automation)"
            MsgBox(sInfo, MsgBoxStyle.Information, m_sTitreMsg)

        Catch ex As Exception

            AfficherMsgErreur2(ex, "bConvertirXLAutomation",
                "Impossible de lire le document :" & vbLf &
                sCheminFichierXL, sMsgErrCausePoss)
            bConvertirXLAutomation = False

        Finally

            msgDelegue.AfficherMsg(sMsgFermetureClasseur)

            ' Option : Ne pas libérer l'instance, si elle n'appartient pas à cette fct
            'oXLH.Fermer(oSht, oWkb, bQuitter:=False, bLibererXLSiResteOuvert:=False)

            'oXLH.Quitter()
            oXLH.Fermer(oSht, oWkb, bQuitter:=True)
            msgDelegue.AfficherMsg("")
            Sablier(bDesactiver:=True)

        End Try

        msgDelegue.AfficherMsg(sMsgLectureTerminee)

Fin:
        'Sablier(bDesactiver:=True)

    End Function

#End Region

#Region "Utilitaires"

    ' Attribut pour éviter que l'IDE s'interrompt en cas d'exception
    <System.Diagnostics.DebuggerStepThrough()>
    Function bFeuilleExiste(sFeuille$, oWkb As Object) As Boolean
        On Error Resume Next
        bFeuilleExiste = CBool(Len(oWkb.Worksheets(sFeuille).Name) > 0)
    End Function

    'Const sListeCellules$ = "Feuil1!A1;Feuil1!B2;Feuil2!C3;Feuil2!C1"
    'Const sListeCellules$ = "Feuil1!D19"
    'Const sListeCellules$ = "Feuil1!A1"
    'Dim aoValeurs() As Object = Nothing
    'bOk = bLireCellulesXLAutomation(Me.m_sCheminFichierXL, sListeCellules, aoValeurs, Me.m_msgDelegue)
    'Debug.WriteLine(aoValeurs(0).ToString)
    'Debug.WriteLine(aoValeurs(0).ToString & ", " & aoValeurs(1).ToString & ", " & _
    '    aoValeurs(2).ToString & ", " & aoValeurs(3).ToString)

    Public Function bLireCellulesXLAutomation(sCheminFichierXL$,
            sListeCellules$, ByRef aoValeurs() As Object,
            msgDelegue As clsMsgDelegue,
            Optional bQuitter As Boolean = True) As Boolean

        ' Lire des cellules dans un classeur Excel de manière sécurisée (non ODBC)
        '  et retourner un tableau de valeurs

        bLireCellulesXLAutomation = False

        If Not bFichierExiste(sCheminFichierXL, bPrompt:=True) Then Return False

        If Not bFichierAccessibleMultiTest(sCheminFichierXL, msgDelegue) Then _
            Return False

        Dim sCheminDossierXL$ = IO.Path.GetDirectoryName(sCheminFichierXL)

        ' Optimisation : conserver la précédante instance d'Excel
        'Dim oXLH As clsHebExcel
        Dim oWkb As Object = Nothing
        Dim oSht As Object = Nothing

        msgDelegue.m_bAnnuler = False

        Try

            Sablier()
            msgDelegue.AfficherMsg(sMsgLancementExcel)
            If IsNothing(oXLH) Then oXLH = New clsHebExcel(bInterdireAppliAvant:=False)
            If IsNothing(oXLH.oXL) Then
                msgDelegue.AfficherMsg(sMsgPbExcel)
                GoTo Fin
            End If

            'oXLH.oXL.Visible = True ' Mode Debug
            oXLH.oXL.Visible = False

            msgDelegue.AfficherMsg("Ouverture du classeur...")
            oWkb = oXLH.oXL.Workbooks.Open(sCheminFichierXL)

            Dim asPlages$() = sListeCellules.Split(";"c)
            Dim iNbValeurs% = asPlages.GetUpperBound(0)
            ReDim aoValeurs(0 To iNbValeurs)
            Dim iNumValeur% = 0
            For Each sFeuillePlage As String In asPlages
                If String.IsNullOrEmpty(sFeuillePlage) Then Exit For
                Dim asChamps$() = sFeuillePlage.Split("!"c)
                Dim sFeuille$ = asChamps(0)
                Dim sPlage$ = asChamps(1)
                Dim iCol% = iColPlage(sPlage)
                Dim iLigne% = iLignePlage(sPlage)
                Dim oVal As Object = oWkb.Worksheets(sFeuille).Cells(iLigne, iCol).Value
                aoValeurs(iNumValeur) = oVal
                iNumValeur += 1
            Next

            bLireCellulesXLAutomation = True

        Catch ex As Exception

            AfficherMsgErreur2(ex, "bLireCellulesXLAutomation",
                "Impossible de lire le classeur :" & vbLf &
                sCheminFichierXL, sMsgErrCausePoss)
            Return False

        Finally

            msgDelegue.AfficherMsg(sMsgFermetureClasseur)

            ' Option : Ne pas libérer l'instance, si elle n'appartient pas à cette fct
            'oXLH.Fermer(oSht, oWkb, bQuitter:=False, bLibererXLSiResteOuvert:=False)

            'oXLH.Quitter()
            oXLH.Fermer(oSht, oWkb, bQuitter, bLibererXLSiResteOuvert:=bQuitter)
            If bQuitter Then oXLH = Nothing : LibererRessourceDotNet()
            msgDelegue.AfficherMsg("")

        End Try

        msgDelegue.AfficherMsg("Lecture terminée.")

Fin:
        Sablier(bDesactiver:=True)

    End Function

    Public Function bLireCellulesXLCouleurs(sCheminFichierXL$,
            sListeCellules$, ByRef aiIdxCouleurs%(),
            msgDelegue As clsMsgDelegue,
            Optional bQuitter As Boolean = True) As Boolean

        ' Lire la couleur des cellules dans un classeur Excel
        '  et retourner un tableau de valeurs

        bLireCellulesXLCouleurs = False

        If Not bFichierExiste(sCheminFichierXL, bPrompt:=True) Then Exit Function

        If Not bFichierAccessibleMultiTest(sCheminFichierXL, msgDelegue) Then _
            Exit Function

        Dim sCheminDossierXL$ = IO.Path.GetDirectoryName(sCheminFichierXL)

        ' Optimisation : conserver la précédante instance d'Excel
        'Dim oXLH As clsHebExcel
        Dim oWkb As Object = Nothing
        Dim oSht As Object = Nothing

        msgDelegue.m_bAnnuler = False

        Try

            Sablier()
            msgDelegue.AfficherMsg(sMsgLancementExcel)
            If IsNothing(oXLH) Then oXLH = New clsHebExcel(bInterdireAppliAvant:=False)
            If IsNothing(oXLH.oXL) Then
                msgDelegue.AfficherMsg(sMsgPbExcel)
                GoTo Fin
            End If

            'oXLH.oXL.Visible = True ' Mode Debug
            oXLH.oXL.Visible = False

            msgDelegue.AfficherMsg("Ouverture du classeur...")
            oWkb = oXLH.oXL.Workbooks.Open(sCheminFichierXL)

            Dim asPlages$() = sListeCellules.Split(";"c)
            Dim iNbValeurs% = asPlages.GetUpperBound(0)
            ReDim aiIdxCouleurs(0 To iNbValeurs)
            Dim iNumValeur% = 0
            For Each sFeuillePlage As String In asPlages
                If String.IsNullOrEmpty(sFeuillePlage) Then Exit For
                Dim asChamps$() = sFeuillePlage.Split("!"c)
                Dim sFeuille$ = asChamps(0)
                Dim sPlage$ = asChamps(1)
                Dim iCol% = iColPlage(sPlage)
                Dim iLigne% = iLignePlage(sPlage)
                aiIdxCouleurs(iNumValeur) = oWkb.Worksheets(sFeuille).
                    Cells(iLigne, iCol).Interior.ColorIndex
                iNumValeur += 1
            Next

            bLireCellulesXLCouleurs = True

        Catch ex As Exception

            AfficherMsgErreur2(ex, "bLireCellulesXLCouleurs",
                "Impossible de lire le classeur :" & vbLf &
                sCheminFichierXL, sMsgErrCausePoss)

        Finally

            msgDelegue.AfficherMsg(sMsgFermetureClasseur)

            ' Option : Ne pas libérer l'instance, si elle n'appartient pas à cette fct
            'oXLH.Fermer(oSht, oWkb, bQuitter:=False, bLibererXLSiResteOuvert:=False)

            'oXLH.Quitter()
            oXLH.Fermer(oSht, oWkb, bQuitter, bLibererXLSiResteOuvert:=bQuitter)
            If bQuitter Then oXLH = Nothing : LibererRessourceDotNet()
            msgDelegue.AfficherMsg("")

        End Try

        msgDelegue.AfficherMsg("Lecture terminée.")

Fin:
        Sablier(bDesactiver:=True)

    End Function

    Public Function iColPlage%(sPlage$, Optional iNumChamp% = 0)

        ' Renvoyer la première colonne d'une plage
        ' Voir aussi : iConvLettresEnNum

        iColPlage = 0

        If sPlage.Length = 0 Then Exit Function

        Dim asTab$() = sPlage.Split(":"c)
        Dim sDeb$ = ""
        If iNumChamp > asTab.GetUpperBound(0) Then
            ' Si la plage est par exemple A1 alors C=1 et L=1
            sDeb = asTab(0)
            GoTo Suite
        End If
        sDeb = asTab(iNumChamp)
Suite:
        Dim iValA% = Asc("A") ' 65
        ' Si la plage ne définie que les colonnes, il n'y a qu'un caractère
        If sDeb.Length = 1 Then
            Dim sCol$ = sDeb.Chars(0)
            iColPlage = 1 + Asc(sCol.Chars(0)) - iValA
            Exit Function
        End If
        iColPlage = 1
        'Dim sCar2Deb$ = sDeb.Chars(1)
        'If IsNumeric(sCar2Deb) Then
        If Char.IsNumber(sDeb.Chars(1)) Then
            ' Soit le début de la plage est du type A9 ou A99
            Dim sCol$ = sDeb.Chars(0)
            iColPlage = 1 + Asc(sCol.Chars(0)) - iValA
        Else
            ' Soit le début de la plage est du type AA9 ou AA99
            Dim sCol$ = sDeb.Substring(0, 2)
            iColPlage = 26 * (1 + Asc(sCol.Chars(0)) - iValA) + 1 + Asc(sCol.Chars(1)) - iValA
        End If

    End Function

    Public Function iLignePlage%(sPlage$, Optional iNumChamp% = 0)

        ' Renvoyer la première ligne d'une plage

        iLignePlage = 0

        If sPlage.Length = 0 Then Exit Function

        Dim asTab$() = sPlage.Split(":"c)
        Dim sDeb$ = ""
        If iNumChamp > asTab.GetUpperBound(0) Then
            ' Si la plage est par exemple A1 alors C=1 et L=1
            sDeb = asTab(0)
            GoTo Suite
        End If
        sDeb = asTab(iNumChamp)
        ' Si la plage ne définie que les colonnes, il n'y a qu'un caractère
        If sDeb.Length = 1 Then
            If iNumChamp = 0 Then
                iLignePlage = 1 '  La première ligne est donc 1
            ElseIf iNumChamp = 1 Then
                iLignePlage = 65535 ' Renvoyer la dernière ligne dans ce cas
            End If
            Exit Function
        End If
Suite:
        iLignePlage = 1
        'Dim sCar2Deb$ = sDeb.Chars(1)
        'If IsNumeric(sCar2Deb) Then
        If Char.IsNumber(sDeb.Chars(1)) Then
            ' Soit le début de la plage est du type A9 ou A99
            Dim sLigne$ = sDeb.Substring(1)
            iLignePlage = CInt(sLigne)
        Else
            ' Soit le début de la plage est du type AA9 ou AA99
            Dim sLigne$ = sDeb.Substring(2)
            iLignePlage = CInt(sLigne)
        End If

    End Function

    Public Function iColFinPlage%(sPlage$)

        ' Renvoyer la dernière colonne d'une plage
        iColFinPlage = iColPlage(sPlage, iNumChamp:=1)

    End Function

    Public Function iLigneFinPlage%(sPlage$)

        ' Renvoyer la dernière ligne d'une plage
        iLigneFinPlage = iLignePlage(sPlage, iNumChamp:=1)

    End Function

    Public Function sConvNumEnLettres$(iCol%)

        Dim iValA% = Asc("A") ' 65
        If iCol <= 26 Then
            sConvNumEnLettres = Chr(iValA + iCol - 1)
        Else ' Corrigé le 23/06/2010
            Dim iMult26% = (iCol - 1) \ 26 ' CInt(iCol / 26)
            Dim iReste% = (iCol - 1) Mod 26 ' iCol - iMult26 * 26
            sConvNumEnLettres = Chr(iValA + iMult26 - 1) & Chr(iValA + iReste)
        End If

    End Function

#End Region

End Module