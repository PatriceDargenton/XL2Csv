
' Pour compiler ce code, ajouter la dll SpreadsheetGear.dll dans le projet
'  et mettre cette constant à True

' On peut redistribuer la dll SpreadSheetGear à un client final, mais pas à un site 
'  de développeurs (donc pas VBFrance/Comment-ça-marche)
#Const bSpreadSheetGearRedist = False

#If Not bSpreadSheetGearRedist Then

Module modExcelSSG

    Public Const bSpreadSheetGear As Boolean = False

    Public Function bConvertirXLRapideSSG(sCheminFichierXL$,
        msgDelegue As clsMsgDelegue) As Boolean
        MsgBox("La conversion selon SpreadSheetGear n'est pas activée dans cette version !",
            MsgBoxStyle.Exclamation, m_sTitreMsg)
        Return False
    End Function

End Module

#Else

Imports System.Text ' Pour StringBuilder
Imports SpreadsheetGear
Imports SpreadsheetGear.Advanced.Cells

Module modExcelSSG

    Public const bSpreadSheetGear as boolean = True

    Public Enum DataTypeValue
        VideOuErr
        String_
        Date_
        Numeric_
        Bool_
    End Enum

    Public Function bConvertirXLRapideSSG(sCheminFichierXL$, msgDelegue As clsMsgDelegue) As Boolean

        ' Convertir un classeur Excel en fichiers csv de manière sécurisée (non ODBC)
        '  (cellule par cellule) et rapide, sans Excel, via SpreadSheetGear

        bConvertirXLRapideSSG = False

        Const bSupprimerPtVirgALaFinDesLignes As Boolean = True

        If Not bFichierAccessibleMultiTest(sCheminFichierXL, msgDelegue) Then _
            Exit Function

        Dim sCheminDossierXL$ = System.IO.Path.GetDirectoryName(sCheminFichierXL)

        Dim iNbFichiersCsvGeneres% = 0
        Dim sDernierCheminCsv$ = ""

        Try

            Sablier()
            msgDelegue.AfficherMsg(sMsgOuvertureClasseur)

            ' Open a workbook into a new workbook set which uses the current culture.
            Dim workbookSSG As SpreadsheetGear.IWorkbook = _
                SpreadsheetGear.Factory.GetWorkbook(sCheminFichierXL, _
                    Globalization.CultureInfo.CurrentCulture)

            If IsNothing(workbookSSG) Then
                msgDelegue.AfficherMsg(String.Format( _
                    "Impossible d'ouvrir le classeur Excel '{0}' !", sCheminFichierXL))
                Exit Function
            End If

            Dim iFeuille% = 0
            Dim iNbFeuilles% = workbookSSG.Sheets.Count

            For Each sheet As SpreadsheetGear.IWorksheet In workbookSSG.Worksheets

                iFeuille += 1
                Dim bAuMoinsUneVal As Boolean = False

                If IsNothing(sheet) Then Continue For

                Dim sFeuille$ = sheet.Name
                Dim sFeuilleDos$ = sConvNomDos(sFeuille)

                ' Limiter le sb à la taille utile (supprimer les lignes vides à la fin)
                Dim sb As New StringBuilder
                Dim sbTmp2 As New StringBuilder
                Dim iLongUtile2% = -1

                Dim iNumLigneDep% = 0 ' NPOI : sheet.FirstRowNum
                Dim iNumLigne% = iNumLigneDep
                'Dim iNbLignes% = sheet.LastRowNum + 1 ' NPOI
                'Dim iNbLignes0% = sheet.UsedRange.Rows.Count + 1
                Dim iNbLignes% = sheet.UsedRange.RowCount + 1
                'Dim iNbCol0% = sheet.UsedRange.Columns.Count + 1
                Dim iNbCol% = sheet.UsedRange.ColumnCount + 1
                'Dim range As IRange = sheet.UsedRange
                Dim values As IValues = CType(sheet, IValues)
                Dim range As IRange = sheet.Cells

                Do While iNumLigne < iNbLignes

                    If iNumLigne = iNumLigneDep Or iNumLigne = iNbLignes - 1 Or _
                       iNumLigne Mod 1000 = 0 Then
                        msgDelegue.AfficherMsg( _
                            "Feuille n°" & iFeuille & "/" & iNbFeuilles & _
                            " : Ligne n°" & iNumLigne + 1 & "/" & iNbLignes & _
                            " : Lecture en cours...")
                        If msgDelegue.m_bAnnuler Then Exit Function
                    End If

                    Const iColMin% = 0
                    Dim iColMax% = iNbCol

                    ' Inférieur à 0 signifie ligne vide : Integer.MinValue NPOI
                    'If iColMax < 0 Then GoTo FinLigne

                    ' D'abord trouver la dernière cellule existante de la ligne
                    Dim bLigneVide As Boolean = True
                    Dim iColMaxLigne% = iColMax
                    Dim j%
                    For j = iNbCol - 1 To iColMin Step -1

                        Dim cell0 As IValue = values(iNumLigne, j)
                        If IsNothing(cell0) Then Continue For
                        'Dim cell0 As ICell = row.GetCell(j) ' NPOI

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

                        Dim cell0 As IValue = values(iNumLigne, j)
                        If IsNothing(cell0) Then GoTo Suite

                        'If iFeuille = 1 And j = 0 And iNumLigne = 2 Then
                        '    Debug.WriteLine("!")
                        'End If

                        ' get the formatted value of a cell
                        Dim sVal$ = range(iNumLigne, j).Text
                        If String.IsNullOrEmpty(sVal) Then GoTo Suite

                        'Dim typeCell As DataTypeValue = GetValueType(range, iNumLigne, j)
                        'Dim sVal$ = ""
                        ''If typeCell = DataTypeValue.VideOuErr Then GoTo Suite
                        'Select Case typeCell
                        'Case DataTypeValue.VideOuErr : GoTo Suite
                        'Case DataTypeValue.String_ : sVal = cell0.Text
                        'Case DataTypeValue.Numeric_ : sVal = cell0.Number.ToString
                        'Case DataTypeValue.Date_
                        '    Dim dDate As Date = workbookSSG.NumberToDateTime(cell0.Number)
                        '    sVal = dDate.ToString(sFormatDateFixe)
                        'Case DataTypeValue.Bool_ : sVal = cell0.Text
                        'End Select

                        'Dim sVal$ = cell0.Text
                        'If String.IsNullOrEmpty(sVal) Then
                        '    sVal = cell0.Number.ToString
                        '    If String.IsNullOrEmpty(sVal) Then GoTo Suite
                        'End If

                        ' NPOI
                        'Dim cell0 As ICell = row.GetCell(j)
                        'If IsNothing(cell0) Then GoTo Suite
                        'Dim sValue$ = GetValue(cell0, dataFormatter, formulaEvaluator)
                        'Dim sVal$ = If(IsNullOrWhiteSpace(sValue), "", sValue)

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
                        If bSupprimerPtVirgALaFinDesLignes Then
                            sbTmp.Length = iLongUtile
                        End If
                        sbTmp2.Append(sbTmp)
                        iLongUtile2 = sbTmp2.Length
                    End If

    FinLigne:
                    'Debug.WriteLine(sbTmp2.ToString)
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

            Next

            msgDelegue.AfficherMsg(sMsgOperationTerminee)
            If iNbFichiersCsvGeneres = 1 Then
                Dim sInfo$ = "(via le composant SpreadSheetGear)"
                ProposerOuvrirFichier(sDernierCheminCsv, sInfo)
            Else
                Dim sInfo$ = "Le classeur :" & vbCrLf & sCheminFichierXL & vbCrLf & _
                    "a été converti en fichiers csv avec succès !" & vbCrLf & _
                    "(via le composant SpreadSheetGear)"
                MsgBox(sInfo, MsgBoxStyle.Information, m_sTitreMsg)
            End If
            bConvertirXLRapideSSG = True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bConvertirXLRapideSSG", _
                "Impossible de lire le document :" & vbLf & _
                sCheminFichierXL)

        Finally
            Sablier(bDesactiver:=True)

        End Try

        msgDelegue.AfficherMsg(sMsgOperationTerminee)

    End Function

    'Private Function sGetFormattedValue$(range As IRange, iLigne%, iCol%)
    '    Return range(iLigne, iCol).Text
    'End Function

    'Private Function GetValueType(range As IRange, iLigne%, iCol%) As DataTypeValue

    '    Select Case range(iLigne, iCol).ValueType
    '    Case SpreadsheetGear.ValueType.Text
    '        Return DataTypeValue.String_
    '        Exit Select
    '    Case SpreadsheetGear.ValueType.Number
    '        If range(iLigne, iCol).NumberFormatType = NumberFormatType.[Date] OrElse _
    '           range(iLigne, iCol).NumberFormatType = NumberFormatType.DateTime Then
    '            Return DataTypeValue.Date_
    '        Else
    '            Return DataTypeValue.Numeric_
    '        End If
    '        Exit Select
    '    Case SpreadsheetGear.ValueType.Logical
    '        Return DataTypeValue.Bool_
    '        Exit Select
    '    Case Else
    '        Return DataTypeValue.VideOuErr
    '    End Select

    'End Function

    '#Region "Ecritures cellule SSG"

    'Private Const rValNull! = -9999
    'Private Const rValNullDouble# = -9999
    'Private Const iValNull% = -9999
    'Private Const lValNull& = -9999

    'Private Sub EcrireCelluleSSG(sheet As SpreadsheetGear.IWorksheet, _
    '    iLigne%, iCol%, sVal$)

    '    sheet.Cells(iLigne - 1, iCol - 1).Value = sVal

    'End Sub

    'Private Sub EcrireCelluleSSG(sheet As SpreadsheetGear.IWorksheet, _
    '    iLigne%, iCol%, iVal%)

    '    If iVal = iValNull Then EffacerCelluleSSG(sheet, iLigne, iCol) : Exit Sub
    '    sheet.Cells(iLigne - 1, iCol - 1).Value = iVal

    'End Sub

    'Private Sub EcrireCelluleSSG(sheet As SpreadsheetGear.IWorksheet, _
    '    iLigne%, iCol%, rVal As Double, _
    '    Optional bEffacerSiNul As Boolean = True)

    '    If rVal = rValNullDouble Then
    '        If bEffacerSiNul Then EffacerCelluleSSG(sheet, iLigne, iCol)
    '        Exit Sub
    '    End If

    '    sheet.Cells(iLigne - 1, iCol - 1).Value = rVal

    'End Sub

    'Private Sub EcrireCelluleSSG(sheet As SpreadsheetGear.IWorksheet, _
    '    iLigne%, iCol%, dDate As Date)

    '    If dDate = dDateNulle Then EffacerCelluleSSG(sheet, iLigne, iCol) : Exit Sub

    '    sheet.Cells(iLigne - 1, iCol - 1).Value = dDate

    'End Sub

    'Private Sub EffacerCelluleSSG(sheet As SpreadsheetGear.IWorksheet, _
    '    iLigne%, iCol%)

    '    'sheet.Cells(iLigne - 1, iCol - 1).ClearContents()

    'End Sub

    '#End Region

    End Module

#End If