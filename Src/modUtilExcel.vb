
Imports System.Text ' Pour StringBuilder

Module modUtilExcel

    Public Const sMsgOuvertureClasseur$ = "Ouverture du classeur..."
    Public Const sMsgFermetureClasseur$ = "Fermeture du classeur..."
    Public Const sMsgLectureTerminee$ = "Lecture terminée."
    Public Const sMsgOperationTerminee$ = "Opération terminée."

    ' Créer des fichiers csv stables : fixer le format des dates
    Public Const sFormatDateFixe$ = "dd\/MM\/yyyy"
    ' Afficher toujours l'heure : si elle est vide, on la supprimera
    ' (\/ pour fixer le / et \: pour fixer le :, sinon dépend du format régional)
    Public Const sFormatDateHeureFixe$ = "dd\/MM\/yyyy HH\:mm\:ss"
    Public Const sHeureVide$ = " 00:00:00"

    Public Const sSeparateurMilliersDef$ = "" ' Ne pas séparer les milliers dans les csv
    Public Const iNbDecimalesMax% = -1 ' Précision maximale, par convention
    Public Const iNbDecimalesDef% = iNbDecimalesMax ' Précision maximale, par convention

    Public Function sFormaterNumeriqueDble$(rVal As Double,
            Optional bSupprimerPt0 As Boolean = True,
            Optional iNbDecimales% = 1,
            Optional sSeparateurMilliers$ = " ")

        ' Formater un numérique avec une précision d'une décimale

        ' Le format numérique standard est correct (séparation des milliers et plus), 
        '  il suffit juste d'enlever la décimale inutile si 0

        Dim nfi As Globalization.NumberFormatInfo =
            New Globalization.NumberFormatInfo
        ' Définition des spérateurs numériques
        nfi.NumberGroupSeparator = sSeparateurMilliers   ' Séparateur des milliers, millions...
        nfi.NumberDecimalSeparator = "." ' Séparateur décimal
        ' 3 groupes pour milliard, million et millier 
        ' (on pourrait en ajouter un 4ème pour les To : 1000 Go)
        nfi.NumberGroupSizes = New Integer() {3, 3, 3}
        If iNbDecimales >= 0 Then
            nfi.NumberDecimalDigits = iNbDecimales ' 1 décimale de précision
        ElseIf iNbDecimales = iNbDecimalesMax Then
            ' Même si 17 chiffres maximum sont gérés en interne, la précision de la valeur 
            '  Double ne comporte par défaut que 15 chiffres décimaux.
            ' http://msdn.microsoft.com/fr-fr/library/system.double.aspx
            nfi.NumberDecimalDigits = 15
        End If

        sFormaterNumeriqueDble = rVal.ToString("n", nfi) ' n : numérique général
        ' Enlever la décimale si 0
        If bSupprimerPt0 Then

            If iNbDecimales = iNbDecimalesMax Then

                ' Précision maximale
                If sSeparateurMilliers.Length = 0 Then

                    ' S'il n'y a pas de séparateur de millier, juste forcer . décimal
                    Dim sVal$ = rVal.ToString ' Le .0 est déjà automatiquement supprimé dans ce cas
                    sVal = sValeurPtDecimal(sVal) ' Il reste juste à forcer le . décimal
                    sFormaterNumeriqueDble = sVal

                Else

                    ' Sinon reprendre le format séparé et traiter les chiffres
                    Dim sVal$ = sFormaterNumeriqueDble
                    'Dim sValOrig$ = sVal
                    ' Déjà pris en charge par le format "n"
                    'sVal = sValeurPtDecimal(sVal) ' Forcer le . décimal
                    Dim iPosPt% = sVal.IndexOf(".")
                    If iPosPt > -1 Then
                        Dim iLong% = sVal.Length
                        Dim i%
                        ' Enlever les 0 à la fin seulement (non significatifs)
                        For i = iLong To iPosPt Step -1
                            Dim cChiffre As Char = sVal.Chars(i - 1)
                            If cChiffre <> "0" Then Exit For
                        Next
                        If i < iLong Then
                            If i = iPosPt + 1 Then i = iPosPt
                            sVal = sVal.Substring(0, i)
                            sFormaterNumeriqueDble = sVal
                        End If
                    End If
                    'Debug.WriteLine(sValOrig & " -> " & sVal)
                End If

            ElseIf iNbDecimales = 1 Then
                sFormaterNumeriqueDble = sFormaterNumeriqueDble.Replace(".0", "")
            ElseIf iNbDecimales > 1 Then
                Dim i%
                Dim sb As New StringBuilder(".")
                For i = 1 To iNbDecimales : sb.Append("0") : Next
                sFormaterNumeriqueDble = sFormaterNumeriqueDble.Replace(sb.ToString, "")
            End If
        End If

    End Function

    Public Function sFormaterNumeriqueDec$(rVal As Decimal,
            Optional bSupprimerPt0 As Boolean = True,
            Optional iNbDecimales% = 1,
            Optional sSeparateurMilliers$ = " ")

        ' Formater un numérique avec une précision d'une décimale
        ' Dble -> Dec n'est pas autorisé
        ' Dec  -> Dble est autorisé : donc dble est le + général

        Dim oDbl As Double = CDbl(rVal)
        sFormaterNumeriqueDec = sFormaterNumeriqueDble(oDbl,
            bSupprimerPt0, iNbDecimales, sSeparateurMilliers)

        ' Le format numérique standard est correct (séparation des milliers et plus), 
        '  il suffit juste d'enlever la décimale inutile si 0

        'Dim nfi As Globalization.NumberFormatInfo = _
        '    New Globalization.NumberFormatInfo
        '' Définition des spérateurs numériques
        'nfi.NumberGroupSeparator = sSeparateurMilliers   ' Séparateur des milliers, millions...
        'nfi.NumberDecimalSeparator = "." ' Séparateur décimal
        '' 3 groupes pour milliard, million et millier 
        '' (on pourrait en ajouter un 4ème pour les To : 1000 Go)
        'nfi.NumberGroupSizes = New Integer() {3, 3, 3}
        'nfi.NumberDecimalDigits = iNbDecimales ' 1 décimale de précision
        'sFormaterNumeriqueDec = rVal.ToString("n", nfi) ' n : numérique général
        '' Enlever la décimale si 0
        'If bSupprimerPt0 Then
        '    If iNbDecimales = 1 Then
        '        sFormaterNumeriqueDec = sFormaterNumeriqueDec.Replace(".0", "")
        '    ElseIf iNbDecimales > 1 Then
        '        Dim i%
        '        Dim sb As New StringBuilder(".")
        '        For i = 1 To iNbDecimales : sb.Append("0") : Next
        '        sFormaterNumeriqueDec = sFormaterNumeriqueDec.Replace(sb.ToString, "")
        '    End If
        'End If

    End Function

End Module