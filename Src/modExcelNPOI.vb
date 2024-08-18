
' NPOI : the .Net port of Apache POI : POIFS ('Poor Obfuscation Implementation' File System) 
' Librairie pour lire les fichiers MS-Office <= 2003
' https://www.nuget.org/packages/NPOI
' https://github.com/nissl-lab/npoi

Imports NPOI.HSSF.UserModel
Imports NPOI.POIFS.FileSystem ' Pour POIFSFileSystem
Imports NPOI.SS.UserModel

Imports System.Text ' Pour StringBuilder

Module modExcelNPOI

#Region "XL2Csv"

    Public Function bConvertirXLRapideNPOI(sCheminFichierXL$, msgDelegue As clsMsgDelegue) As Boolean

        ' Convertir un classeur Excel en fichiers csv de manière sécurisée (non ODBC)
        '  (cellule par cellule) et rapide, sans Excel, via NPOI

        bConvertirXLRapideNPOI = False

        Const bSupprimerPtVirgALaFinDesLignes As Boolean = True

        ' 11/03/2012 bEcriture:=False : Non ne marche pas !
        If Not bFichierAccessibleMultiTest(sCheminFichierXL, msgDelegue) Then Exit Function

        Dim sCheminDossierXL$ = IO.Path.GetDirectoryName(sCheminFichierXL)

        Dim iNbFichiersCsvGeneres% = 0
        Dim sDernierCheminCsv$ = ""

        Try

            Sablier()
            msgDelegue.AfficherMsg(sMsgOuvertureClasseur)

            Using inp As New System.IO.StreamReader(sCheminFichierXL)

                Dim workbook As New HSSFWorkbook(New POIFSFileSystem(inp.BaseStream))

                If IsNothing(workbook) Then
                    msgDelegue.AfficherMsg(String.Format(
                        "Impossible d'ouvrir le classeur Excel '{0}' !", sCheminFichierXL))
                    '"Excel Workbook '{0}' could not be opened.", sCheminFichierXL))
                    Exit Function
                End If

                Dim formulaEvaluator As New HSSFFormulaEvaluator(workbook)
                'Const sCulture$ = "fr-FR" '"en-US"
                'Dim sCulture$ = System.Globalization.CultureInfo.CurrentCulture.Name
                Dim dataFormatter As New HSSFDataFormatter(
                    System.Globalization.CultureInfo.CurrentCulture)
                'New System.Globalization.CultureInfo(System.Threading.Thread.CurrentCulture.))

                Dim iFeuille% = 0
                Dim iNbFeuilles% = workbook.NumberOfSheets

                For Each sheet As ISheet In workbook

                    iFeuille += 1
                    Dim bAuMoinsUneVal As Boolean = False

                    If IsNothing(sheet) Then Continue For

                    Dim sb As New StringBuilder
                    Dim sbTmp2 As New StringBuilder
                    Dim iLongUtile2% = -1
                    Dim sFeuille$ = sheet.SheetName
                    Dim sFeuilleDos$ = sConvNomDos(sFeuille)

                    'Dim iNbCol% = sheet.PhysicalNumberOfRows ' Pas fiable
                    'Dim iColMaxFeuille% = iTrouverColMaxFeuille(sheet)

                    ' 09/08/2012 0 au lieu de 1, sinon on peut rater la 1ère ligne
                    Dim iNumLigneDep% = 0 'sheet.FirstRowNum
                    Dim iNumLigne% = iNumLigneDep ' i
                    Dim iNbLignes% = sheet.LastRowNum + 1

                    Do While iNumLigne < iNbLignes '<= iNbLignes

                        If iNumLigne = iNumLigneDep Or iNumLigne = iNbLignes - 1 Or
                       iNumLigne Mod 1000 = 0 Then
                            msgDelegue.AfficherMsg(
                            "Feuille n°" & iFeuille & "/" & iNbFeuilles &
                            " : Ligne n°" & iNumLigne + 1 & "/" & iNbLignes &
                            " : Lecture en cours...")
                            If msgDelegue.m_bAnnuler Then Exit Function
                        End If

                        Dim row As IRow = sheet.GetRow(iNumLigne)
                        If IsNothing(row) Then GoTo FinLigne ' 09/08/2012

                        Const iColMin% = 0
                        'Dim iColMin% = row.FirstColIndex
                        'If iColMin = Integer.MaxValue Then GoTo FinLigne

                        Dim iColMax% = row.LastCellNum
                        ' Inférieur à 0 signifie ligne vide : Integer.MinValue
                        If iColMax < 0 Then GoTo FinLigne

                        ' D'abord trouver la dernière cellule existante de la ligne
                        Dim bLigneVide As Boolean = True
                        Dim iColMaxLigne% = iColMax
                        Dim j%
                        'For j = iNbCol - 1 To iColMin Step -1
                        For j = iColMax To iColMin Step -1
                            Dim cell0 As ICell = row.GetCell(j)
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
                            Dim cell0 As ICell = row.GetCell(j)
                            If IsNothing(cell0) Then GoTo Suite

                            ' GetValue perd les formules avec heure
                            ' Date au format anglais ! (même avec culture FR)
                            Dim sVal$ = GetValue(cell0, dataFormatter, formulaEvaluator)
                            'Dim sVal$ = sLireValeur(cell0) ' Ne lit pas les dates
                            'Dim sVal$ = If(IsNullOrWhiteSpace(sValue), "", sValue)
                            'Debug.WriteLine(sVal)

                            sbTmp.Append(sVal)
                            If sVal.Length > 0 Then
                                bAuMoinsUneVal = True
                                bAuMoinsUneValLigne = True
                                iLongUtile = sbTmp.Length
                            End If

Suite:
                            If j < iColMax Then sbTmp.Append(";")
                            j += 1
                        Loop ' Colonnes
                        If bAuMoinsUneValLigne Then
                            ' Retirer les ; à la fin
                            If bSupprimerPtVirgALaFinDesLignes Then
                                sbTmp.Length = iLongUtile
                            End If
                            sbTmp2.Append(sbTmp)
                            iLongUtile2 = sbTmp2.Length
                        End If

FinLigne:
                        sbTmp2.Append(vbCrLf)
                        iNumLigne += 1

                    Loop ' Lignes

                    If Not bAuMoinsUneVal Then Continue For
                    Dim sChemin$ = sCheminDossierXL & "\" & sFeuilleDos & ".csv"

                    ' Limiter le sb à la taille utile (supprimer les lignes vides à la fin)
                    sb.Append(sbTmp2)
                    If bSupprimerPtVirgALaFinDesLignes Then
                        sb.Length = iLongUtile2 + 2 ' +2 pour vbCrLf
                    End If

                    ' 03/09/2017 Encodage UTF8
                    If Not bEcrireFichier(sChemin, sb, bEncodageUTF8:=True) Then Exit Function
                    iNbFichiersCsvGeneres += 1
                    sDernierCheminCsv = sChemin

                Next sheet

            End Using

            msgDelegue.AfficherMsg(sMsgOperationTerminee)
            If iNbFichiersCsvGeneres = 1 Then
                Dim sInfo$ = "(via le composant NPOI)"
                ProposerOuvrirFichier(sDernierCheminCsv, sInfo)
            Else
                Dim sInfo$ = "Le classeur :" & vbCrLf & sCheminFichierXL & vbCrLf &
                    "a été converti en fichiers csv avec succès !" & vbCrLf &
                    "(via le composant NPOI)"
                MsgBox(sInfo, MsgBoxStyle.Information, m_sTitreMsg)
            End If
            bConvertirXLRapideNPOI = True

        Catch ex As NPOI.POIFS.FileSystem.OfficeXmlFileException
            Dim ex2 As Exception = ex
            ' "POI only supports OLE2 Office documents"
            AfficherMsgErreur2(ex2, "bConvertirXLRapideNPOI",
                "Impossible de lire le document :" & vbLf & sCheminFichierXL,
                "Excel >= 2007 n'est pas supporté actuellement dans XL2Csv via la techno. NPOI") ' 18/08/2024
            '"Le support d'Excel 2007 est prévu dans la version 1.3 de NPOI (actuellement 1.2.5)")
        Catch ex As Exception
            AfficherMsgErreur2(ex, "bConvertirXLRapideNPOI",
                "Impossible de lire le document :" & vbLf & sCheminFichierXL)

        Finally
            Sablier(bDesactiver:=True)

        End Try

        msgDelegue.AfficherMsg(sMsgOperationTerminee)

    End Function

#End Region

    Private Function sLireValeur$(cell As ICell)

        ' Ok, mais ne lit pas les dates ! (affiche un numérique à la place)

        If IsNothing(cell) Then Return String.Empty

        If cell.CellType = CellType.Numeric OrElse
          (cell.CellType = CellType.Formula AndAlso
           cell.CachedFormulaResultType = CellType.Numeric) Then
            Dim sVal1$ = cell.NumericCellValue.ToString
            Return sVal1
        Else
            Dim sVal1$ = cell.StringCellValue
            Return sVal1
        End If

    End Function

    Private Function GetValue(cell As ICell, dataFormatter As DataFormatter,
            formulaEvaluator As IFormulaEvaluator) As String

        Dim sRet$ = String.Empty
        If IsNothing(cell) Then Return sRet

        Try
            Dim sVal1$ = dataFormatter.FormatCellValue(cell, formulaEvaluator)
            Const cEsp As Char = " "c
            'Const sEsp$ = " "
            Dim sVal2$ = sVal1.Replace(Microsoft.VisualBasic.ChrW(10), cEsp)
            Return sVal2
        Catch ex As Exception
            Return sRet
        End Try

    End Function

End Module