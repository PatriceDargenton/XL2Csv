
' Gestion Excel via ExcelLibrary
' https://www.nuget.org/packages/ExcelLibrary
' https://code.google.com/archive/p/excellibrary/
' http://www.codeproject.com/KB/office/ExcelReader.aspx
' https://www.codeproject.com/Articles/16210/Excel-Reader

Imports ExcelLibrary ' Pour BinaryFileFormat
Imports ExcelLibrary.BinaryFileFormat ' Pour WorkbookDecoder
Imports ExcelLibrary.CompoundDocumentFormat ' Pour CompoundDocument
Imports ExcelLibrary.SpreadSheet ' Pour Workbook
Imports System.IO ' Pour MemoryStream
Imports System.Text ' Pour StringBuilder

Module modExcelLibrary

#Region "XL2Csv"

    Public Function bConvertirXLRapide(sCheminFichierXL$, msgDelegue As clsMsgDelegue) As Boolean

        ' Convertir un classeur Excel en fichiers csv de manière sécurisée (non ODBC)
        '  (cellule par cellule) et rapide, sans Excel, via ExcelLibrary

        bConvertirXLRapide = False

        ' 11/03/2012 bEcriture:=False : Non ne marche pas !
        If Not bFichierAccessibleMultiTest(sCheminFichierXL, msgDelegue) Then Exit Function

        Dim sCheminDossierXL$ = IO.Path.GetDirectoryName(sCheminFichierXL)

        Dim doc As CompoundDocument = Nothing
        Dim iNbFichiersCsvGeneres% = 0
        Dim sDernierCheminCsv$ = ""

        Try

            Sablier()
            msgDelegue.AfficherMsg(sMsgOuvertureClasseur)
            doc = CompoundDocument.Open(sCheminFichierXL)

            Dim streamData As Byte() = doc.GetStreamData("Workbook")
            If IsNothing(streamData) Then Exit Function

            Dim workbook0 As Workbook = WorkbookDecoder.Decode(New MemoryStream(streamData))

            Dim iFeuille% = 0
            Dim iNbFeuilles% = workbook0.Worksheets.Count
            For Each sheet As Worksheet In workbook0.Worksheets

                iFeuille += 1
                Dim bAuMoinsUneVal As Boolean = False

                Dim sb As New StringBuilder
                Dim sbTmp2 As New StringBuilder
                Dim iLongUtile2% = -1
                Dim sFeuille$ = sheet.Name
                Dim sFeuilleDos$ = sConvNomDos(sFeuille)

                ' Excel en csv ignore la 1ère ligne, mais ici on va la garder : + stable
                Dim iNumLigne% = 0 'worksheet.Cells.FirstRowIndex
                Dim iNbCol% = sheet.Cells.LastColIndex + 1
                Dim iNbLignes% = sheet.Cells.LastRowIndex

                Do While iNumLigne <= iNbLignes

                    If iNumLigne = 1 Or iNumLigne = iNbLignes Or iNumLigne Mod 1000 = 0 Then
                        msgDelegue.AfficherMsg(
                        "Feuille n°" & iFeuille & "/" & iNbFeuilles &
                        " : Ligne n°" & iNumLigne & "/" & iNbLignes &
                        " : Lecture en cours...")
                        If msgDelegue.m_bAnnuler Then Exit Function
                    End If

                    Dim row0 As ExcelLibrary.SpreadSheet.Row = sheet.Cells.GetRow(iNumLigne)

                    Dim iColMin% = row0.FirstColIndex
                    If iColMin = Integer.MaxValue Then GoTo FinLigne

                    Dim iColMax% = row0.LastColIndex
                    ' Inférieur à 0 signifie ligne vide : Integer.MinValue
                    If iColMax < 0 Then GoTo FinLigne

                    ' D'abord trouver la dernière cellule existante de la ligne
                    Dim bLigneVide As Boolean = True
                    Dim iColMaxLigne% = iColMax
                    Dim j%
                    For j = iNbCol - 1 To iColMin Step -1
                        Dim cell0 As Cell = row0.GetCell(j)
                        If IsNothing(cell0) Then Continue For
                        iColMaxLigne = j
                        bLigneVide = False
                        Exit For
                    Next
                    ' Si la ligne est vide, ne rien mettre
                    If bLigneVide Then GoTo FinLigne

                    Dim bAuMoinsUneValLigne As Boolean = False
                    Dim iLongUtile% = -1
                    Dim sbTmp As New StringBuilder
                    j = 0
                    Do While j <= iColMaxLigne
                        If j < iColMin Then GoTo Suite
                        Dim cell0 As Cell = row0.GetCell(j)
                        If IsNothing(cell0) Then GoTo Suite
                        If IsNothing(cell0.Value) Then GoTo Suite

                        Dim sVal$ = sLireValCelluleExcelLibrary(cell0)

                        sbTmp.Append(sVal)
                        If sVal.Length > 0 Then
                            bAuMoinsUneVal = True
                            bAuMoinsUneValLigne = True
                            iLongUtile = sbTmp.Length
                        End If
Suite:
                        If j < iNbCol - 1 Then sbTmp.Append(";")
                        j += 1
                    Loop ' Colonnes
                    If bAuMoinsUneValLigne Then
                        ' Retirer les ; à la fin
                        sbTmp.Length = iLongUtile
                        sbTmp2.Append(sbTmp)
                        iLongUtile2 = sbTmp2.Length
                    End If

FinLigne:
                    iNumLigne += 1
                    sbTmp2.Append(vbCrLf)
                Loop ' Lignes

                If Not bAuMoinsUneVal Then Continue For
                Dim sChemin$ = sCheminDossierXL & "\" & sFeuilleDos & ".csv"
                ' Limiter le sb à la taille utile (supprimer les lignes vides à la fin)
                sb.Append(sbTmp2)
                sb.Length = iLongUtile2 + 2 ' +2 pour vbCrLf
                ' 03/09/2017 Encodage UTF8
                If Not bEcrireFichier(sChemin, sb, bEncodageUTF8:=True) Then Exit Function
                iNbFichiersCsvGeneres += 1
                sDernierCheminCsv = sChemin

            Next sheet

            msgDelegue.AfficherMsg(sMsgOperationTerminee)
            If iNbFichiersCsvGeneres = 1 Then
                Dim sInfo$ = "(via le composant ExcelLibrary)"
                ProposerOuvrirFichier(sDernierCheminCsv, sInfo)
            Else
                Dim sInfo$ = "Le classeur :" & vbCrLf & sCheminFichierXL & vbCrLf &
                "a été converti en fichiers csv avec succès !" & vbCrLf &
                "(via le composant ExcelLibrary)"
                MsgBox(sInfo, MsgBoxStyle.Information, m_sTitreMsg)
            End If
            bConvertirXLRapide = True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bConvertirXLRapide",
                "Impossible de lire le document :" & vbLf &
                sCheminFichierXL)

        Finally
            If Not IsNothing(doc) Then doc.Close()
            Sablier(bDesactiver:=True)

        End Try

        msgDelegue.AfficherMsg(sMsgOperationTerminee)

    End Function

    Public Function bConvertirXL2Txt(sCheminFichierXL$, msgDelegue As clsMsgDelegue) As Boolean

        ' Convertir un classeur Excel en un fichier texte de manière sécurisée (non ODBC)
        '  (cellule par cellule) et rapide, sans Excel, via ExcelLibrary

        bConvertirXL2Txt = False

        ' En fait on n'a pas besoin de préfixer par le nom de la feuille
        '  car on peut utiliser le chapitrage dans VBTextFinder
        '  pour rappeler où on se trouve dans le classeur
        ' (si on l'active ce préfixage, reste à supprimer les lignes vides à la fin)
        Const bPrefixerParNomFeuille As Boolean = False

        ' 11/03/2012 bEcriture:=False : Non ne marche pas !
        If Not bFichierAccessibleMultiTest(sCheminFichierXL, msgDelegue) Then _
        Exit Function

        Dim sCheminDossierXL$ = IO.Path.GetDirectoryName(sCheminFichierXL)

        Dim doc As CompoundDocument = Nothing

        Try

            Sablier()
            msgDelegue.AfficherMsg(sMsgOuvertureClasseur)
            doc = CompoundDocument.Open(sCheminFichierXL)

            Dim streamData As Byte() = doc.GetStreamData("Workbook")
            If IsNothing(streamData) Then Exit Function

            Dim workbook0 As Workbook = WorkbookDecoder.Decode(New MemoryStream(streamData))

            Dim iFeuille% = 0
            Dim iNbFeuilles% = workbook0.Worksheets.Count
            Dim worksheet0 As Worksheet
            Dim sb As New StringBuilder

            sb.Append("Fichier source : " & sCheminFichierXL & vbCrLf)

            Dim fi As New IO.FileInfo(sCheminFichierXL)
            Dim lTailleFichier& = fi.Length
            Dim sTailleFichier$ = sFormaterTailleOctets(lTailleFichier)
            Dim sTailleFichierDetail$ = sFormaterTailleOctets(lTailleFichier, bDetail:=True)
            ' fi.LastWriteTime affiche toujours la bonne heure (et la même heure)
            sb.Append("Taille : " & sTailleFichierDetail &
                ", Date : " & fi.LastWriteTime & vbCrLf)

            Dim bAuMoinsUneValClasseur As Boolean = False

            For Each worksheet0 In workbook0.Worksheets

                iFeuille += 1
                Dim bAuMoinsUneVal As Boolean = False

                Dim sbTmp2 As New StringBuilder
                Dim iLongUtile2% = -1
                Dim sFeuille$ = worksheet0.Name
                'Dim sFeuilleDos$ = sConvNomDos(sFeuille)

                ' Excel en csv ignore la 1ère ligne, mais ici on va la garder : + stable
                Dim i% = 0 'worksheet.Cells.FirstRowIndex
                Dim iNbCol% = worksheet0.Cells.LastColIndex + 1
                Dim iNbLignes% = worksheet0.Cells.LastRowIndex

                sb.Append(vbCrLf & vbCrLf &
                    "Feuille Excel n°" & iFeuille & " : " & sFeuille & vbCrLf &
                    "-------------" & vbCrLf)

                Dim j%

                Do While (i <= iNbLignes)

                    If i = 1 Or i = iNbLignes Or i Mod 1000 = 0 Then
                        msgDelegue.AfficherMsg(
                            "Feuille n°" & iFeuille & "/" & iNbFeuilles &
                            " : Ligne n°" & i & "/" & iNbLignes &
                            " : Lecture en cours...")
                        If msgDelegue.m_bAnnuler Then Exit Function
                    End If

                    Dim row0 As ExcelLibrary.SpreadSheet.Row = worksheet0.Cells.GetRow(i)

                    Dim iColMin% = row0.FirstColIndex
                    If iColMin = Integer.MaxValue Then GoTo LigneVide

                    Dim iColMax% = row0.LastColIndex
                    ' Inférieur à 0 signifie ligne vide : Integer.MinValue
                    If iColMax < 0 Then GoTo LigneVide

                    ' D'abord trouver la dernière cellule existante de la ligne
                    Dim bLigneVide As Boolean = True
                    Dim iColMaxLigne% = iColMax
                    For j = iNbCol - 1 To iColMin Step -1
                        Dim cell0 As Cell = row0.GetCell(j)
                        If IsNothing(cell0) Then Continue For
                        iColMaxLigne = j
                        bLigneVide = False
                        Exit For
                    Next
                    ' Si la ligne est vide, ne rien mettre
                    If bLigneVide Then GoTo LigneVide

                    Dim bAuMoinsUneValLigne As Boolean = False
                    Dim iLongUtile% = -1
                    Dim sbTmp As New StringBuilder
                    If bPrefixerParNomFeuille Then sbTmp.Append(sFeuille & ";")
                    j = 0
                    Do While j <= iColMaxLigne
                        If j < iColMin Then GoTo Suite
                        Dim cell0 As Cell = row0.GetCell(j)
                        If IsNothing(cell0) Then GoTo Suite
                        If IsNothing(cell0.Value) Then GoTo Suite

                        Dim sVal$ = sLireValCelluleExcelLibrary(cell0)

                        sbTmp.Append(sVal)
                        If sVal.Length > 0 Then
                            bAuMoinsUneVal = True
                            bAuMoinsUneValLigne = True
                            iLongUtile = sbTmp.Length
                        End If
Suite:
                        If j < iNbCol - 1 Then sbTmp.Append(";")
                        j += 1
                    Loop
                    If bAuMoinsUneValLigne Then
                        ' Retirer les ; à la fin
                        sbTmp.Length = iLongUtile
                        sbTmp2.Append(sbTmp)
                        iLongUtile2 = sbTmp2.Length
                    Else
                        GoTo LigneVide
                    End If
                    sbTmp2.Append(vbCrLf)
                    GoTo LigneSuivante

LigneVide:
                    ' Afficher seulement le nom de la feuille Excel
                    If bPrefixerParNomFeuille Then
                        sbTmp2.Append(sFeuille)
                        iLongUtile2 = sbTmp2.Length
                    End If
                    sbTmp2.Append(vbCrLf)

LigneSuivante:
                    i += 1
                Loop

                If Not bAuMoinsUneVal Then Continue For
                'Dim sChemin$ = sCheminDossierXL & "\" & sFeuilleDos & ".csv"
                ' Limiter le sb à la taille utile (supprimer les lignes vides à la fin)
                sbTmp2.Length = iLongUtile2 + 2 ' +2 pour vbCrLf
                sb.Append(sbTmp2)
                bAuMoinsUneValClasseur = True

            Next

            msgDelegue.AfficherMsg(sMsgOperationTerminee)

            If Not bAuMoinsUneValClasseur Then
                Dim sInfo$ = "Le classeur est vide !" & vbCrLf & sCheminFichierXL
                MsgBox(sInfo, MsgBoxStyle.Information, m_sTitreMsg)
            Else
                Dim sChemin$ = sCheminDossierXL & "\" &
                IO.Path.GetFileNameWithoutExtension(sCheminFichierXL) & ".txt"
                ' 03/09/2017 Encodage UTF8
                If Not bEcrireFichier(sChemin, sb, bEncodageUTF8:=True) Then Exit Function
                Dim sInfo$ = "(via le composant ExcelLibrary)"
                ProposerOuvrirFichier(sChemin, sInfo)
            End If
            bConvertirXL2Txt = True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bConvertirXL2Txt",
                "Impossible de lire le document :" & vbLf & sCheminFichierXL)

        Finally
            If Not IsNothing(doc) Then doc.Close()
            Sablier(bDesactiver:=True)

        End Try

        msgDelegue.AfficherMsg(sMsgOperationTerminee)

    End Function

#End Region

#Region "Fonction utilitaires ExcelLibrary"

    Private Const sTypeErrorCode$ = "ExcelLibrary.BinaryFileFormat.ErrorCode"
    Private Const iAnneeNulleExcel% = 1899
    Private Const iAnneeMinExcel% = 1945

    Public Function sLireValCelluleExcelLibrary$(cell0 As Cell)

        Dim sVal$ = ""
        Dim sType$ = cell0.Value.GetType.ToString
        If sType = sTypeErrorCode Then
            Dim errCode As BinaryFileFormat.ErrorCode =
            DirectCast(cell0.Value, BinaryFileFormat.ErrorCode)
            sVal = errCode.Value.ToString()
        ElseIf cell0.Format.FormatType = CellFormatType.Date OrElse
           cell0.Format.FormatType = CellFormatType.DateTime OrElse
           cell0.Format.FormatType = CellFormatType.Time OrElse
           cell0.Format.FormatType = CellFormatType.Custom Then
            sVal = sLireDate(cell0)
        Else
            sVal = sLireVal(cell0)
        End If
        sLireValCelluleExcelLibrary = sVal

    End Function

    ' Attribut pour éviter que l'IDE s'interrompt en cas d'exception
    '<System.Diagnostics.DebuggerStepThrough()> _
    Private Function sLireDate$(cell0 As Cell, Optional bSupprimerHeureVide As Boolean = True)

        Try
            Dim dDate As Date = cell0.DateTimeValue

            If dDate.Year = iAnneeNulleExcel OrElse
           dDate.Year < iAnneeMinExcel Then
                ' Il n'y a aucun moyen de savoir si une valeur est une date ou pas sous Excel
                '  en particulier pour le format personnalisé
                ' Solution : si l'année est 1899, il ne s'agit probablement pas d'une date
                '  et aussi si < 1945 : rare
                'sLireDate = cell0.Value.ToString()
                sLireDate = sLireVal(cell0)
                Exit Function
            End If

            ' On impose un format
            'Dim sFormat$ = ""
            Dim sFormat$ = sFormatDateHeureFixe
            'If cell0.Format.FormatType = CellFormatType.Date Then sFormat = sFormatDateFixe
            If String.Compare(cell0.Format.FormatType.ToString,
                          CellFormatType.Date.ToString) = 0 Then _
            sFormat = sFormatDateFixe
            'cell0.Format.FormatType = CellFormatType.DateTime
            'cell0.Format.FormatType = CellFormatType.Time
            'cell0.Format.FormatType = CellFormatType.Custom

            sLireDate = dDate.ToString(sFormat)

        Catch 'ex As Exception
            ' Il peut y avoir un format date appliqué à du texte
            sLireDate = cell0.Value.ToString()
        End Try

        If Not bSupprimerHeureVide Then Exit Function
        If Not sLireDate.EndsWith(sHeureVide) Then Exit Function
        sLireDate = Left$(sLireDate, sLireDate.Length - sHeureVide.Length)

    End Function

    Private Function sLireVal$(cell0 As Cell)

        Dim sType$ = cell0.Value.GetType.ToString
        Dim bDbl As Boolean = (String.Compare(sType, "System.Double") = 0)
        Dim bDec As Boolean = (String.Compare(sType, "System.Decimal") = 0)
        ' Si réel alors appliquer le format
        'sLireVal = dVal.ToString(sFormat)
        'Dim sFormat$ = cell0.Format.FormatString
        If bDbl Then
            ' Dble -> Dec n'est pas autorisé
            ' Dec  -> Dble est autorisé : donc dble est le + général
            Dim dVal As Double = CDbl(cell0.Value)
            sLireVal = sFormaterNumeriqueDble(dVal, iNbDecimales:=iNbDecimalesDef,
            sSeparateurMilliers:=sSeparateurMilliersDef)
        ElseIf bDec Then
            Dim dVal As Decimal = CDec(cell0.Value)
            sLireVal = sFormaterNumeriqueDec(dVal, iNbDecimales:=iNbDecimalesDef,
            sSeparateurMilliers:=sSeparateurMilliersDef)
        Else
            sLireVal = cell0.Value.ToString()
        End If

    End Function

#End Region

End Module
