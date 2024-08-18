
' Fichier modUtil.vb
' ------------------

Module modUtil

    Public Sub AfficherMsgErreur2(ByRef Ex As Exception,
            Optional sTitreFct$ = "", Optional sInfo$ = "",
            Optional sDetailMsgErr$ = "",
            Optional bCopierMsgPressePapier As Boolean = True,
            Optional ByRef sMsgErrFinal$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        If Ex.Message <> "" Then
            sMsg &= vbCrLf & Ex.Message.Trim
            If Not IsNothing(Ex.InnerException) Then _
                sMsg &= vbCrLf & Ex.InnerException.Message
        End If
        If bCopierMsgPressePapier Then CopierPressePapier(sMsg)
        sMsgErrFinal = sMsg
        MsgBox(sMsg, MsgBoxStyle.Critical)

    End Sub

    Public Sub Sablier(Optional ByRef bDesactiver As Boolean = False)

        If bDesactiver Then
            Cursor.Current = Cursors.Default
        Else
            Cursor.Current = Cursors.WaitCursor
        End If

    End Sub

    Public Sub CopierPressePapier(sInfo$)

        ' Copier des informations dans le presse-papier de Windows
        ' (elles resteront jusqu'à ce que l'application soit fermée)

        Try
            Dim dataObj As New DataObject
            dataObj.SetData(DataFormats.Text, sInfo)
            Clipboard.SetDataObject(dataObj)
        Catch ex As Exception
            ' Le presse-papier peut être indisponible
            AfficherMsgErreur2(ex, "CopierPressePapier",
                bCopierMsgPressePapier:=False)
        End Try

    End Sub

    Public Sub TraiterMsgSysteme_DoEvents()

        Try
            Application.DoEvents() ' Peut planter avec OWC : Try Catch nécessaire
        Catch
        End Try

    End Sub

    Public Sub LibererRessourceDotNet()

        ' 19/01/2011 Il faut appeler 2x :
        '  cf. All-In-One Code Framework\Visual Studio 2008\VBAutomateWord

        ' Clean up the unmanaged Word COM resources by forcing a garbage 
        ' collection as soon as the calling function is off the stack (at 
        ' which point these objects are no longer rooted).
        GC.Collect()
        GC.WaitForPendingFinalizers()
        ' GC needs to be called twice in order to get the Finalizers called 
        ' - the first time in, it simply makes a list of what is to be 
        ' finalized, the second time in, it actually the finalizing. Only 
        ' then will the object do its automatic ReleaseComObject.
        GC.Collect()
        GC.WaitForPendingFinalizers()

        TraiterMsgSysteme_DoEvents()

    End Sub

    Public Function bFichierAccessibleMultiTest(sChemin$, msgDelegue As clsMsgDelegue,
            Optional iDelaiMSec% = 2000,
            Optional iNbTentatives% = 3,
            Optional bEcriture As Boolean = True) As Boolean

        ' Voir si un fichier est accessible (en simple lecture) avec plusieurs tentatives
        '  dans le cas de partage multi-utilisateur (on suppose que le vérrouillage 
        '  ne dure pas lontemps)

        bFichierAccessibleMultiTest = False

        ' 1ère tentative sans message
Retenter:
        If Not bFichierAccessible(sChemin, bInexistOk:=True, bEcriture:=bEcriture) Then
            Dim bOk As Boolean = False
            ' Tentatives suivantes avec message
            For i As Integer = 1 To iNbTentatives
                msgDelegue.AfficherMsg("Tentative n°" & i &
                    " de lecture du fichier : " & sChemin)
                Attendre(iDelaiMSec)
                bOk = bFichierAccessible(sChemin)
                If bOk Then Exit For
            Next
            If Not bOk Then
                Dim sMsg$ = "Le fichier n'est pas accessible en lecture :"
                If bEcriture Then sMsg = "Le fichier n'est accessible en écriture :"
                msgDelegue.AfficherMsg(sMsg & " " & sChemin)
                Dim sInfo$ = ""
                If IO.Path.GetExtension(sChemin).ToLower = ".xls" Then
                    sInfo = "(le classeur est probablement ouvert avec Excel)" & vbCrLf
                End If
                If MsgBoxResult.Retry = MsgBox(sMsg & vbCrLf & sChemin & vbCrLf & sInfo &
                    "Voulez-vous réessayer ?",
                    MsgBoxStyle.Exclamation Or MsgBoxStyle.RetryCancel, m_sTitreMsg) Then
                    GoTo Retenter
                End If
                Exit Function
            End If
        End If
        bFichierAccessibleMultiTest = True

    End Function

    Public Sub Attendre(iDelaiMSec%)
        'TraiterMsgSysteme_DoEvents()
        Threading.Thread.Sleep(iDelaiMSec)
    End Sub

    Public Function sValeurPtDecimal$(sVal$)

        sValeurPtDecimal = sVal
        If sVal.Length = 0 Then Exit Function
        sValeurPtDecimal = Replace(sValeurPtDecimal, ",", ".")
        Dim sSepDecimal$ = Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator()
        If sSepDecimal.Length = 0 Then Exit Function
        If sSepDecimal <> "." And sSepDecimal <> "," Then
            ' Quelque soit le séparateur décimal, le convertir en .
            sValeurPtDecimal = Replace(sValeurPtDecimal, sSepDecimal, ".")
        End If

    End Function

    Public Function iConv%(sVal$, Optional iValDef% = 0)

        If sVal.Length = 0 Then iConv = iValDef : Exit Function
        Try
            iConv = CInt(sVal)
        Catch
            iConv = iValDef
        End Try

    End Function

End Module