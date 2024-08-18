
' Fichier clsODBC.vb
' ------------------

Imports System.Text ' Pour StringBuilder

Public Class clsODBC

#Region "Déclarations"

    Public Const sValErreurDef$ = "#Erreur#"

    ' Evénement signalant l'arrivée d'un message 
    ' (avancement de l'opération en cours ou bien erreur par exemple)
    'Public Event EvAfficherMessage(sMsg$)
    ' CA1009
    Public Delegate Sub GestEvAfficherMessage(sender As Object, e As clsMsgEventArgs)
    Public Event EvAfficherMessage As GestEvAfficherMessage

    ' Si vous voulez contrôler strictement l'état des variables affectées 
    '  depuis l'extérieur de la classe, alors utilisez des propriétés 
    '  Set et Get, et passez ces variables membres en privé dans ce cas

    ' Requête faite à la volée par le code 
    ' (ou bien liste de requêtes SQL séparées par des ; )
    ' (au lieu de requêtes figurant dans un fichier .sql externe)
    Public m_sListeSQL$

    ' Requête spécifique dans le cas où la source est un fichier Excel
    Public m_sListeSQLExcel$

    ' Chaîne de connexion directe à un fichier source, par exemple un fichier Excel
    Public m_sChaineConnexionDirecte$

    ' Chemins vers un fichier DSN et une requête SQL stockés en externe
    Public m_sCheminDSN$, m_sCheminSQL$

    ' Chemins et SQL par défaut lors de la création automatique des fichiers DNS et SQL
    Public m_sCheminSrcExcel$, m_sCheminSrcAccess$, m_sCheminSrcOmnis$
    Public m_sSQLExcelDef$, m_sSQLAccessDef$, m_sSQLOmnisDef$ ' SQL ou liste de SQL
    Public m_sSQLNavisionDef$, m_sSQLDB2Def$

    ' Pour les accès ODBC nécessitant une authentification
    Public m_sCompteUtilisateur$, m_sMotDePasse$

    ' Pour les accès ODBC de type serveur, comme Navision, DB2, ...
    Public m_sCompteSociete$, m_sNomServeur$

    ' Afficher les messages dans les boites de dialogues
    Public m_bPrompt As Boolean
    ' Générer des événements pour afficher le détail des opérations en cours
    Public m_bAfficherMsg As Boolean

    ' Booléen pour indiquer si le pilote ODBC supporte le retour arrière 
    ' (vrai pour Excel et Access, faux pour Omnis)
    ' C'est utile pour connaitre à l'avance le nombre de lignes de la source ODBC
    '  mais cela peut faire perdre du temps : on peut laisser à faux dans ce cas
    Public m_bODBCArriere As Boolean

    ' Utile pour effectuer une requête action via une chaîne de connexion directe
    Public m_bModeEcriture As Boolean

    ' Copier tout le contenu retourné par les requêtes SQL dans le presse-papier
    Public m_bCopierDonneesPressePapier As Boolean

    ' Vérifier la présence du fichier source de données 
    ' (ne pas vérifier s'il n'y a pas de fichier spécifique)
    Public m_bVerifierFichierSourceDonnees As Boolean

    ' Vérifier le risque d'erreur de lecture avec Excel < 2003
    Public m_bVerifierConfigODBCExcel As Boolean

    ' Possibilité d'annuler proprement le requêtage depuis l'interface
    Private m_bAnnuler As Boolean

    ' Si on lance des requêtes succesives par petits groupes de données
    '  permet de conserver si une annulation a été demandé
    Public m_bNePasInitAnnulation As Boolean

    ' S'il y a plusieurs requêtes consécutives (liste de SQL séparés par un ;),
    ' cette option permet d'interrompre la requête en cours, 
    ' mais de poursuivre les autres requêtes
    Public m_bInterrompreSeulementRqEnCours As Boolean

    ' Remplacer le séparateur décimal dans les valeurs par le . 
    '  pour pouvoir convertir les nombres en réels via Val
    Public m_bRemplacerSepDec As Boolean
    ' Remplacer seulement les champs numériques : tester avec IsNumeric
    ' (attention : IsNumeric est très lent : mieux vaut remplacer tous les champs)
    ' Autre solution : se baser sur le schéma de la table pour détecter les numériques
    Public m_bRemplacerSepDecNumSeul As Boolean
    Private m_bRemplacerSepDecRequis As Boolean
    Private m_sSepDecimal$

    Public m_bEnleverEspacesFin As Boolean ' Appliquer un TrimEnd = RTrim
    Public m_bRemplacerVraiFaux As Boolean
    Public m_sValVrai$, m_sValFaux$ ' Valeurs à appliquer en guise de Vrai et Faux
    Public m_sValErreur$
    ' Indiquer la présence d'au moins 1 erreur de lecture de la valeur d'un champ
    '  (pour l'ensemble des requêtes successives)
    Public m_bErreursLecture As Boolean

    ' Méthode ADODB.GetString : Attention, le format des dates peut être différent
    Public m_bLireToutDUnBloc As Boolean
    ' Délimiteur ; par défaut et pas de traitement du contenu des champs :
    Public m_bLireToutDUnBlocRapide As Boolean
    Public m_sbLignes As StringBuilder

    ' Stocker les résultats 
    Public m_aoMetaTableau() As Object
    ' Explorateur ODBC
    Public m_alTables As ArrayList
    Public m_asChamps$(,)
    ' 18/11/2007
    Public m_sNomTableMaxChamps$, m_iNumTableMaxChamps%, m_aiNbChamps%()
    Public m_sbContenuRetour As StringBuilder
    Public m_bAjouterChronoDebug As Boolean

    Private Const sTypeODBCExcel$ = "Excel"
    Private Const sTypeODBCAccess$ = "Access"
    Private Const sTypeODBCOmnis$ = "Omnis"
    Private Const sTypeODBCNavision$ = "Navision"
    Private Const sTypeODBCDB2$ = "DB2"

    ' Nombre d'enregistrement alloués à l'avance pour le stockage des lignes
    Public m_iNbEnregAlloues%
    Private Const iNbEnregAllouesDef% = 100

    Private m_oConn As ADODB.Connection = Nothing
    Public m_iNbTentatives% = 0 ' Tentatives de lecture, par ex. fichier Excel partagé

    Public ReadOnly Property bAnnuler() As Boolean
        Get ' Savoir si une annulation est en cours
            bAnnuler = Me.m_bAnnuler
        End Get
    End Property
    Public Sub Annuler()
        ' Demander une annulation
        Me.m_bAnnuler = True
    End Sub

#End Region

#Region "Divers"

    Public Sub New()

        Me.m_sCheminDSN = ""
        Me.m_sCheminSQL = ""
        Me.m_sChaineConnexionDirecte = ""
        Me.m_sListeSQL = ""
        Me.m_sListeSQLExcel = ""
        Me.m_sCheminSrcExcel = ""
        Me.m_sCheminSrcAccess = ""
        Me.m_sCheminSrcOmnis = ""
        Me.m_sSQLExcelDef = ""
        Me.m_sSQLAccessDef = ""
        Me.m_sSQLOmnisDef = ""
        Me.m_sSQLNavisionDef = ""
        Me.m_sSQLDB2Def = ""
        Me.m_sNomTableMaxChamps = ""
        Me.m_iNumTableMaxChamps = 0

        Me.m_sCompteSociete = ""
        Me.m_sNomServeur = ""
        Me.m_sCompteUtilisateur = ""
        Me.m_sMotDePasse = ""
        Me.m_bODBCArriere = False
        Me.m_bCopierDonneesPressePapier = True
        Me.m_bPrompt = True
        Me.m_bRemplacerSepDec = True
        Me.m_bRemplacerSepDecNumSeul = False
        Me.m_bEnleverEspacesFin = True
        Me.m_bRemplacerVraiFaux = True
        Me.m_sValVrai = "1"
        Me.m_sValFaux = ""
        Me.m_sValErreur = sValErreurDef
        Me.m_bNePasInitAnnulation = False
        Me.m_bInterrompreSeulementRqEnCours = False
        Me.m_bAfficherMsg = True
        Me.m_bVerifierFichierSourceDonnees = True
        Me.m_bVerifierConfigODBCExcel = True
        Me.m_bLireToutDUnBloc = False
        Me.m_bLireToutDUnBlocRapide = False
        Me.m_bAjouterChronoDebug = True
        Me.m_iNbEnregAlloues = iNbEnregAllouesDef
        LibererRessources()

    End Sub

    Public Sub LibererRessources()

        Me.m_bErreursLecture = False
        Me.m_bAnnuler = False
        Me.m_aoMetaTableau = Nothing
        Me.m_alTables = Nothing
        Me.m_asChamps = Nothing
        Me.m_aiNbChamps = Nothing
        'Me.m_sLignes = ""
        Me.m_sbLignes = New StringBuilder
        ViderContenuResultat()

        If Not Me.m_oConn Is Nothing Then
            Me.m_oConn.Close()
            Me.m_oConn = Nothing
        End If

    End Sub

    Public Sub ViderContenuResultat()
        Me.m_sbContenuRetour = Nothing
    End Sub

    Private Sub AfficherMessage(sMsg$)

        If Not Me.m_bAfficherMsg Then Exit Sub

        ' CA1009
        'RaiseEvent EvAfficherMessage(sMsg)
        Dim e As New clsMsgEventArgs(sMsg)
        RaiseEvent EvAfficherMessage(Me, e)

        Application.DoEvents()

    End Sub

    Private Sub AfficherErreursADO(oConnexion As ADODB.Connection, ByRef sMsgErr$)

        ' Note sur ByVal oConnexion : 
        '  En VB .NET, il n'est plus nécessaire de passer les objets par 
        '  reférence. De plus, le ByVal est plus rapide (même pour les objets), 
        '  ce qui n'est pas le cas en VB6. Explication : en VB .NET
        '  si on utilise ByVal, l'objet est copié une fois, mais il est copié
        '  2 fois dans le cas du ByRef, selon "VB.NET Professionnel" de Wrox Team

        If oConnexion Is Nothing Then Exit Sub
        Dim sMsg$ = ""
        Dim oErrADO As ADODB.Error
        For Each oErrADO In oConnexion.Errors
            sMsg &= "Erreur ADO : " & oErrADO.Description & vbCrLf
            sMsg &= "Numéro : " & oErrADO.Number & " (" &
                Hex(oErrADO.Number) & ")" & vbCrLf
            If oErrADO.SQLState <> "" Then _
                sMsg &= "Erreur Jet : " & oErrADO.SQLState & vbCrLf
            If oErrADO.Number = -2147467259 Then
                ' Si le pilote ODBC n'est pas installé, on peut obtenir l'erreur :
                ' [Microsoft][Gestionnaire de pilotes ODBC] 
                ' Source de données introuvable et nom de pilote non spécifié"
                ' Numéro : -2147467259 (80004005), Erreur Jet : IM002
                sMsg &= "Cause possible : Le pilote ODBC spécifié n'est pas installé sur ce poste." & vbCrLf
            End If
            If oErrADO.Number = -2147217884 Then
                ' L'ensemble de lignes ne prend pas en charge les récupérations arrière
                sMsg &= "Explication : Le pilote ODBC ne supporte pas le retour en arrière." & vbCrLf
                sMsg &= "(Utilisez m_bODBCArriere = False en paramètre)" & vbCrLf
            End If
            MsgBox(sMsg, MsgBoxStyle.Critical, m_sTitreMsg)
        Next oErrADO
        sMsgErr &= vbCrLf & sMsg
    End Sub

    Public Shared Sub VerifierConfigODBCExcel()

        ' Vérifier la configuration ODBC pour Excel :
        ' Pour Excel < 2003, la configuration par défaut peut être insuffisante
        '  voir la fonction bCreerFichierDsnODBC()
        Const sCle$ = "SOFTWARE\Microsoft\Jet\4.0\Engines\Excel"
        Const sSousCleTGR$ = "TypeGuessRows"
        Dim sValCleTGR$ = ""
        If Not bCleRegistreLMExiste(sCle, sSousCleTGR, sValCleTGR) Then Exit Sub

        ' 14/10/2008 Même avec Office2003 le problème existe !
        ' Si on détermine qu'Office2003 ou > est installé, inutile de générer une alerte
        'Const sSousCleWin32$ = "win32"
        'Const sSousCleWin32Old$ = "win32old"
        'Dim sValCleWin32$ = ""
        'Dim sValCleWin32Old$ = ""
        'bCleRegistreLMExiste(sCle, sSousCleWin32, sValCleWin32)
        'bCleRegistreLMExiste(sCle, sSousCleWin32Old, sValCleWin32Old)
        'sValCleWin32 = sValCleWin32.ToLower
        'If sValCleWin32.Length > 0 And sValCleWin32Old.Length > 0 Then
        '    ' 24/11/2007 : Office10 = XP : insuffisant, il faut 11 ou >
        '    If (sValCleWin32.IndexOf("office11\msaexp30.dll") > -1 Or _
        '        sValCleWin32.IndexOf("office12\msaexp30.dll") > -1) And _
        '    sValCleWin32Old.IndexOf("msexcl40.dll") > -1 Then Exit Sub
        'End If

        If sValCleTGR.Length = 0 Then Exit Sub
        ' Eviter IsNumeric : très lent ! AndAlso IsNumeric(sValCleTGR) Then
        Dim iValCleTGR% = iConv(sValCleTGR, -1)
        If Not (iValCleTGR > -1 And iValCleTGR < 1024) Then Exit Sub

        'MsgBox("La configuration ODBC pour Excel risque d'être insuffisante :" & vbLf & _
        '    "Augmentez la valeur pour lire un plus grand nombre de lignes pour déterminer" & vbLf & _
        '    "le type de données capable de stocker les valeurs d'une colonne Excel" & vbLf & _
        '    "TypeGuessRow=" & iValCleTGR & " < 1024" & vbLf & _
        '    "Clé : HKEY_LOCAL_MACHINE\" & sCle & vbLf & _
        '    "Pour cela, il suffit de lancer ODBCExcelAugmenterTypeGuessRows.reg", _
        '    MsgBoxStyle.Exclamation, m_sTitreMsg)

        Dim sNouvVal$ = "16384"
        If MsgBoxResult.Cancel = MsgBox(
            "La configuration ODBC pour Excel risque d'être insuffisante :" & vbLf &
            "Cliquez sur OK pour augmentez la valeur (" & sNouvVal & ")" & vbLf &
            "pour lire un plus grand nombre de lignes pour déterminer" & vbLf &
            "le type de données capable de stocker les valeurs d'une colonne Excel" & vbLf &
            "TypeGuessRow=" & iValCleTGR & " < 1024" & vbLf &
            "Clé : HKEY_LOCAL_MACHINE\" & sCle,
            MsgBoxStyle.Exclamation Or MsgBoxStyle.OkCancel, m_sTitreMsg) Then Exit Sub

        ' Faire la modif par le code si on a le droit
        Dim sMsg$ = "Echec de la correction de TypeGuessRow !"
        Dim bOk As Boolean = False
        If bCleRegistreLMExiste(sCle, sSousCleTGR, sValCleTGR, sNouvVal) Then
            If bCleRegistreLMExiste(sCle, sSousCleTGR, sValCleTGR) Then
                If sValCleTGR = sNouvVal Then _
                    bOk = True : sMsg = "La correction de TypeGuessRow a réussie !"
            End If
        End If
        If bOk Then
            MsgBox(sMsg, MsgBoxStyle.Exclamation, m_sTitreMsg)
        Else
            MsgBox(sMsg, MsgBoxStyle.Critical, m_sTitreMsg)
        End If

    End Sub

#End Region

#Region "Lecture de la source ODBC"

    Public Function bLireSQL(ByRef sListeSQL$, ByRef sContenuDSN$,
        bNoterContenu As Boolean, ByRef sbContenu As StringBuilder,
        Optional bVerifierSQL As Boolean = True,
        Optional ByRef bExcel As Boolean = False) As Boolean

        sListeSQL$ = ""
        sContenuDSN$ = ""

        'Dim bExcel As Boolean = False

        If Me.m_sChaineConnexionDirecte.Length > 0 Then

            If bNoterContenu Then _
                sbContenu.Append("Chaîne de connexion directe : " &
                    Me.m_sChaineConnexionDirecte & vbCrLf)
            sListeSQL = Me.m_sListeSQL
            If Me.m_sChaineConnexionDirecte.IndexOf("Excel") > -1 Then
                bExcel = True
                If Me.m_bVerifierConfigODBCExcel Then VerifierConfigODBCExcel()
            End If

        Else

            ' S'il n'y a pas de chaîne de connexion directe, on utilise un fichier DSN
            '  ainsi qu'un fichier SQL : on peut ainsi personnaliser les requêtes en 
            '  fonction de la source ODBC (si la source DSN est détectée comme étant de 
            '  type Excel, c'est plus simple d'utiliser une requête spécifique 
            '  (Me.m_sListeSQLExcel) que d'ajouter un $ à la fin des noms des tables, 
            '  ce qui n'est envisageable que pour une requête simple

            ' Si le fichier DSN est absent, on peut le créer automatiquement
            If Not bFichierExiste(Me.m_sCheminDSN) Then
                If Not bCreerFichiersDsnEtSQLODBCDefaut() Then Return False
            End If

            sContenuDSN = sLireFichier(Me.m_sCheminDSN)

            ' Si par exemple base AS400, alors ne pas faire de vérification
            '  car DBQ n'indique pas un chemin vers un fichier spécifique du disque dur
            If Me.m_bVerifierFichierSourceDonnees Then
                ' Lorsque le fichier DSN est déjà créé, vérifier la présence de la source ODBC
                '  si le pilote fonctionne ainsi (on teste toutes les possibilités)
                ' Dans le cas d'un accès réseau, cela permet de tester l'accessibilité
                '  à la base plutôt que d'afficher un message d'erreur obscur
                If Not bVerifierCheminODBC("DataFilePath=", sContenuDSN) Then Return False
                If Not bVerifierCheminODBC("DBQ=", sContenuDSN) Then Return False
                If Not bVerifierCheminODBC("Database=", sContenuDSN) Then Return False
                If Not bVerifierCheminODBC("Dbf=", sContenuDSN) Then Return False
                If Not bVerifierCheminODBC("SourceDB=", sContenuDSN) Then Return False
                ' Vérification des dossiers aussi
                If Not bVerifierCheminODBC("DefaultDir=", sContenuDSN,
                    bDossier:=True) Then Return False
                If Not bVerifierCheminODBC("PPath=", sContenuDSN,
                    bDossier:=True) Then Return False
            End If

            ' Si le pilote est pour Omnis et qu'on a oublié de désactiver m_bODBCArriere 
            '  on le fait, car un MoveLast() peut être très très long !
            If Me.m_bODBCArriere AndAlso
                sContenuDSN.IndexOf("DRIVER=OMNIS ODBC Driver") > -1 Then
                Me.m_bODBCArriere = False
            End If

            If sContenuDSN.IndexOf("DRIVER=Microsoft Excel Driver") > -1 Then
                bExcel = True
                If Me.m_bVerifierConfigODBCExcel Then VerifierConfigODBCExcel()
            End If

            If bNoterContenu Then
                sbContenu.Append("Fichier DSN : " & Me.m_sCheminDSN & " : " & vbCrLf)
                sbContenu.Append(sContenuDSN & vbCrLf)
            End If

            If Me.m_sListeSQL.Length > 0 Then
                ' Requête(s) à la volée par le code
                sListeSQL = Me.m_sListeSQL
            Else
                If bVerifierSQL Then
                    If Me.m_sCheminSQL.Length = 0 Then _
                        MsgBox("Le chemin vers le fichier SQL est vide !",
                            MsgBoxStyle.Critical, m_sTitreMsg) : Return False
                    ' S'il n'y a pas de requête à la volée par le code, 
                    '  alors lire le contenu du fichier SQL externe
                    If Not bFichierExiste(Me.m_sCheminSQL, bPrompt:=True) Then _
                        Return False
                    sListeSQL = sLireFichier(Me.m_sCheminSQL)
                End If
            End If

        End If

        If bExcel AndAlso Me.m_sListeSQLExcel.Length > 0 Then _
            sListeSQL = Me.m_sListeSQLExcel

        bLireSQL = True

    End Function

    Private Function bCheminFichierProbable(sChemin$) As Boolean

        ' Voir si le chemin supposé est un vrai chemin, ou bien simplement
        '  un nom de base de données de type serveur,
        '  auquel cas, il ne faut pas chercher à vérifier la présence du fichier
        '  de source de donnée
        If sChemin.IndexOf("\") > -1 Then Return True
        Return False

    End Function

    Public Function bExplorerSourceODBC(
        Optional bExplorerChamps As Boolean = True,
        Optional sNomTableAExplorer$ = "",
        Optional bRenvoyerContenu As Boolean = False) As Boolean

        ' Explorer la structure de la source ODBC indiquée par le fichier .dsn

        ' Pour manipuler des grandes quantités de chaînes, 
        '  StringBuilder est beaucoup plus rapide que String
        Dim sbContenu As StringBuilder = Nothing
        Dim bNoterResultat As Boolean = False
        If bRenvoyerContenu Or Me.m_bCopierDonneesPressePapier Then
            bNoterResultat = True
            sbContenu = New StringBuilder
        End If

        Dim sListeSQL$ = ""
        Dim sContenuDSN$ = ""
        Dim bExcel As Boolean = False
        If Not bLireSQL(sListeSQL, sContenuDSN, bNoterResultat, sbContenu,
            bVerifierSQL:=False, bExcel:=bExcel) Then
            Me.AfficherMessage("Erreur !")
            Return False
        End If

        ' On initialise à Nothing pour éviter les avertissements intempestifs de VB8
        Dim oConn As ADODB.Connection = Nothing
        Dim oRq As ADODB.Recordset = Nothing
        Dim bConnOuverte As Boolean, bRqOuverte As Boolean

        If Not Me.m_bNePasInitAnnulation Then
            Me.m_bAnnuler = False
            Me.m_bErreursLecture = False
        End If

        Try

            oConn = New ADODB.Connection
            oRq = New ADODB.Recordset
            AfficherMessage("Ouverture de la connexion ODBC en cours...")
            Sablier()
            oConn.Mode = ADODB.ConnectModeEnum.adModeRead
            Dim sConnexion$
            If Me.m_sChaineConnexionDirecte.Length = 0 Then
                sConnexion = "FILEDSN=" & Me.m_sCheminDSN & ";"
            Else
                sConnexion = Me.m_sChaineConnexionDirecte
            End If
            oConn.Open(sConnexion)
            bConnOuverte = True

            Me.m_alTables = New ArrayList

            If bNoterResultat Then _
                sbContenu.Append(vbCrLf & vbCrLf & "Tables :" & vbCrLf)

            AfficherMessage("Exploration des tables en cours...")
            oRq.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly

            ' Exploration des clés primaires : non pris en charge par les pilotes ODBC
            'ADODB.SchemaEnum.adSchemaPrimaryKeys

            'Dim asRestrictions$(4) ' Non pris en charge par les pilotes ODBC
            'asRestrictions(0) = Nothing            ' TABLE_CATALOG
            'asRestrictions(1) = Nothing            ' TABLE_SCHEMA
            'asRestrictions(2) = sNomTableAExplorer ' TABLE_NAME
            'asRestrictions(3) = Nothing            ' TABLE_TYPE
            ' http://www.sahirshah.com/articles/ADOOpenSchema.html
            oRq = oConn.OpenSchema(ADODB.SchemaEnum.adSchemaTables) ', asRestrictions)
            bRqOuverte = True

            Dim iNbChamps% = oRq.Fields.Count
            If iNbChamps = 0 Then GoTo RequeteSuivante ' 18/11/2007
            'If iNbChamps = 0 Then bRqOuverte = False : GoTo RequeteSuivante

            If oRq.EOF Then
                If Me.m_bPrompt Then _
                    MsgBox("Aucune table trouvée !", MsgBoxStyle.Exclamation)
                GoTo RequeteSuivante
            End If

            Dim iNumTable% = 0
            While Not oRq.EOF

                Dim sNomTable$ = oRq.Fields("TABLE_NAME").Value.ToString

                'If (iNumTable Mod 10 = 0) And iNumTable > 0 Then
                '    Dim sAvancement$ = _
                '        "Exploration des tables en cours... (enreg. n°" & _
                '        iNumTable + 1 & ")"
                '    AfficherMessage(sAvancement)
                '    ' Interrompre l'exploration
                '    If Me.m_bAnnuler Then Exit While
                'End If

                ' Si une table à explorer est précisée, ne lister que cette table
                ' (car l'exploration peut être très longue sur les grosses bases)
                If sNomTableAExplorer.Length > 0 AndAlso
                   sNomTable <> sNomTableAExplorer Then GoTo TableSuivante

                ' Un classeur Excel contient parfois aussi
                '  des tables fantômes (sauvegarde de l'aperçu impression ?)
                Dim sTypeObjet$ = oRq.Fields("TABLE_TYPE").Value.ToString
                If bExcel AndAlso sNomTable.EndsWith("$Impression_des_t") Then
                    If bNoterResultat Then _
                        sbContenu.Append(sTypeObjet & " : [" &
                            sNomTable & "] : Table fantôme Excel ignorée" & vbCrLf)
                    GoTo TableSuivante
                End If

                ' Autre exemple de table fantôme sous Excel : [MonClasseur$_]
                If bExcel AndAlso Not (sNomTable.EndsWith("$") Or sNomTable.EndsWith("$'")) Then
                    ' Normalement, le nom de la table Excel doit se terminer par $ ou $'
                    ' Parfois (???) on ne peut pas explorer ce genre de table
                    ' Il peut s'agir aussi de plages nommées sous Excel
                    If bNoterResultat Then _
                        sbContenu.Append(sTypeObjet & " : [" &
                            sNomTable & "] : Table fantôme Excel ignorée" & vbCrLf)
                    GoTo TableSuivante
                End If

                Me.m_alTables.Add(sNomTable)
                iNumTable += 1 ' 18/11/2007

                ' Pour Excel, la plupart des tables sont de type "SYSTEM TABLE"

                ' Ignorer les tables systèmes de MS-Access
                'If Left(sNomTable, 4) = "MSys" Then GoTo TableSuivante

                If bNoterResultat Then sbContenu.Append(
                    sTypeObjet & " : [" & sNomTable & "]" & vbCrLf) ' 25/11/2007

                'If bNoterResultat Then
                '    sbContenu.Append(vbCrLf).Append("Informations sur la table :").Append(vbCrLf)
                '    sbContenu.Append(sTypeObjet & " : [" & sNomTable & "]" & vbCrLf)
                '    Dim i%, j%
                '    For i = 0 To oRq.Fields.Count - 1
                '        sbContenu.Append(oRq.Fields(i).Name & _
                '            " : [" & oRq.Fields(i).Value.ToString & "]" & vbCrLf)
                '        'For j = 0 To oRq.Fields(i).Properties.Count - 1
                '        '    sbContenu.Append( _
                '        '        "P " & oRq.Fields(i).Properties(j).Name & _
                '        '        " : [" & oRq.Fields(i).Properties(j).Value.ToString & "]" & vbCrLf)
                '        'Next j
                '    Next i
                'End If

TableSuivante:
                oRq.MoveNext()
                'iNumTable += 1 ' 18/11/2007

            End While
            AfficherMessage("Exploration des tables terminée : " & iNumTable)
            'If bDebug Then Threading.Thread.Sleep(500)

RequeteSuivante:
            If bRqOuverte Then oRq.Close() : bRqOuverte = False

            If Not bExplorerChamps Then GoTo FinOk

            ' Exploration des champs des tables
            ' Documentation : ADO Data Types (incomplet pour Access)
            ' http://www.w3schools.com/ado/ado_datatypes.asp
            ' Comment interpréter les données via ADO OpenSchema adSchemaColumns :
            ' MS SQL DataTypes QuickRef
            ' http://webcoder.info/reference/MSSQLDataTypes.html

            If bNoterResultat Then sbContenu.Append(vbCrLf)
            Dim sTable$
            'Dim iNbTables% = iNumTable
            Dim iNbTables% = Me.m_alTables.Count ' 18/11/2007
            ReDim Me.m_aiNbChamps(iNbTables - 1)
            ReDim Me.m_asChamps(iNbTables, 0)
            iNumTable = 0
            Dim iNbChampsTableMax% = 0
            For Each sTable In Me.m_alTables

                If (iNumTable Mod 10 = 0 Or iNumTable = iNbTables - 1) And iNumTable > 0 Then
                    Dim sAvancement$ =
                        "Exploration des champs en cours... (table n°" &
                        iNumTable + 1 & "/" & iNbTables & ")"
                    AfficherMessage(sAvancement)
                    ' Interrompre l'exploration
                    If Me.m_bAnnuler Then
                        sbContenu.Append(
                            "(interruption de l'utilisateur)").Append(vbCrLf)
                        Exit For
                    End If
                End If

                ' Attention, avec une connexion directe sur un fichier Excel
                '  l'ordre des champs est perdu ! mais pas avec un dsn !!!
                ' Heureusement, en lisant la valeur du champ ORDINAL_POSITION 
                '  et en stockant le résultat dans un tableau de string,
                '  on retrouve l'ordre exact des champs
                oRq.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                oRq = oConn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns,
                    New Object() {Nothing, Nothing, sTable})
                bRqOuverte = True
                If bNoterResultat Then _
                    sbContenu.Append(vbCrLf & "Table [" & sTable & "] :" & vbCrLf)

                ' Ne marche pas ici :
                'oRq.MoveLast()
                'Dim iNbChampsTable% = oRq.RecordCount
                'oRq.MoveFirst()

                Dim iNumChampMax% = 0
                Dim iNumChamp% = 0
                If (oRq.BOF And oRq.EOF) Then GoTo TableSuivante2 ' Table vide 18/11/2007
                While Not oRq.EOF
                    Dim iNumChampTable% = 0
                    Dim oValChamp As Object = oRq.Fields("ORDINAL_POSITION").Value
                    If IsDBNull(oValChamp) Then
                        iNumChampTable = iNumChamp
                    Else
                        iNumChampTable = CInt(oValChamp) - 1
                    End If
                    If iNumChampTable > iNumChampMax Then _
                        iNumChampMax = iNumChampTable
                    iNumChamp += 1
                    oRq.MoveNext()
                End While
                oRq.MoveFirst()
                Me.m_aiNbChamps(iNumTable) = iNumChampMax
                Dim iNbChampsTable% = iNumChampMax
                If iNbChampsTable > iNbChampsTableMax Then
                    iNbChampsTableMax = iNbChampsTable
                    Me.m_sNomTableMaxChamps = sTable ' 18/11/2007
                    Me.m_iNumTableMaxChamps = iNumTable
                End If

                ' Prendre tjrs le max du nbre de champs sur toutes les tables
                ReDim Preserve Me.m_asChamps(iNbTables, iNbChampsTableMax)

                iNumChamp = 0

                While Not oRq.EOF

                    Dim sDescription$ = ""
                    If Not IsDBNull(oRq.Fields("Description").Value) Then _
                        sDescription = oRq.Fields("Description").Value.ToString
                    Dim sChamp$ = oRq.Fields("COLUMN_NAME").Value.ToString
                    Dim oValChamp As Object = oRq.Fields("ORDINAL_POSITION").Value
                    Dim iNumChampTable% = 1
                    If IsDBNull(oValChamp) Then
                        iNumChampTable = iNumChamp
                    Else
                        iNumChampTable = CInt(oValChamp) - 1
                    End If

                    If bNoterResultat Then
                        Dim sAffTaille$ = ""
                        Dim lTailleCar& = 0
                        If Not IsDBNull(oRq.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Then
                            lTailleCar = CLng(oRq.Fields("CHARACTER_MAXIMUM_LENGTH").Value)
                            If lTailleCar = 1073741823 Then
                                sAffTaille = ":1Go"
                            Else
                                sAffTaille = ":" & lTailleCar.ToString
                            End If
                        End If

                        Dim sAffTypeDonnees$ = ""
                        Dim lDataType& = 0
                        If Not IsDBNull(oRq.Fields("DATA_TYPE").Value) Then
                            lDataType& = CLng(oRq.Fields("DATA_TYPE").Value)
                            Dim lVal As ADODB.DataTypeEnum = CType(lDataType, ADODB.DataTypeEnum)
                            sAffTypeDonnees = " (" & lVal.ToString & sAffTaille & ")"
                        End If

                        Dim sAffDescr$ = ""
                        If sDescription.Length > 0 Then sAffDescr = " : " & sDescription
                        sbContenu.Append("  [" & sChamp & "]" &
                            sAffTypeDonnees & sAffDescr & vbCrLf)
                    End If

                    'Dim lFlags& = 0
                    'If Not IsDBNull(oRq.Fields("COLUMN_FLAGS").Value) Then
                    '    lFlags = CLng(oRq.Fields("COLUMN_FLAGS").Value)
                    'End If

                    Me.m_asChamps(iNumTable, iNumChampTable) = sChamp

                    iNumChamp += 1
                    oRq.MoveNext()

                End While

TableSuivante2:
                oRq.Close() : bRqOuverte = False
                iNumTable += 1

            Next sTable
            If iNbTables > 0 Then
                AfficherMessage("Exploration des champs terminée : " &
                    iNumTable & "/" & iNbTables)
                If bDebug Then Threading.Thread.Sleep(500)
            End If

FinOk:
            If bNoterResultat Then
                If sNomTableAExplorer.Length > 0 And Me.m_alTables.Count = 0 Then
                    sbContenu.Append(
                        "Table [" & sNomTableAExplorer & "] non trouvée !" & vbCrLf)
                End If
                sbContenu.Append(vbCrLf & vbCrLf)
                sbContenu.Append(
                    "Documentation : ADO Data Types (incomplet pour Access) :" & vbCrLf)
                sbContenu.Append("www.w3schools.com/ado/ado_datatypes.asp" & vbCrLf)
            End If

        Catch ex As Exception

            Sablier(bDesactiver:=True)
            Dim sMsg$ = ""
            If Me.m_sChaineConnexionDirecte.Length = 0 Then
                sMsg &= vbCrLf & "Dsn : " & Me.m_sCheminDSN
            Else
                sMsg &= vbCrLf & "Chaîne de connexion : " & Me.m_sChaineConnexionDirecte
            End If
            Dim sDetailMsgErr$ = ""
            ' Ne pas copier l'erreur dans le presse-papier maintenant 
            '  car on va le faire dans le Finally

            Dim sMsgErrFinal$, sMsgErrADO$, sDetail$
            If bConnOuverte Then
                sDetail = "Certains champs sont peut-être introuvables, ou bien :"
            Else
                sDetail = "Erreur lors de l'ouverture de la connexion "
                If sContenuDSN.Length > 0 Then
                    sDetail &= "'" & sLireNomPiloteODBC(sContenuDSN) & "' :"
                Else
                    sDetail &= ":"
                End If
            End If
            sMsgErrFinal = "" : sMsgErrADO = ""
            AfficherMsgErreur2(ex, "bExplorerSourceODBC", sMsg, sDetail,
                bCopierMsgPressePapier:=False, sMsgErrFinal:=sMsgErrFinal)
            If Me.m_bCopierDonneesPressePapier Then _
                sbContenu.Append(vbCrLf & sMsgErrFinal & vbCrLf)
            AfficherErreursADO(oConn, sMsgErrADO)
            If Me.m_bCopierDonneesPressePapier Then _
                sbContenu.Append(sMsgErrADO & vbCrLf)
            Me.AfficherMessage("Erreur !")
            Return False

        Finally

            Sablier(bDesactiver:=True)
            If bRqOuverte Then oRq.Close() : bRqOuverte = False
            ' Connexion ADODB et non OleDb
            If bConnOuverte Then oConn.Close() : bConnOuverte = False

            ' Copier les informations dans le presse-papier (utile pour le debogage)
            If Me.m_bCopierDonneesPressePapier Then _
                CopierPressePapier(sbContenu.ToString)
            ' Dans le cas de plusieurs accès ODBC, 
            '  on peut avoir besoin de mémoriser tous les contenus successifs
            If bRenvoyerContenu Then
                sbContenu.Append(vbCrLf).Append(vbCrLf).Append(vbCrLf)
                If IsNothing(Me.m_sbContenuRetour) Then _
                    Me.m_sbContenuRetour = New StringBuilder
                Me.m_sbContenuRetour.Append(sbContenu)
            End If

        End Try

        If Me.m_bPrompt Then
            Me.AfficherMessage("Opération terminée.")
            Dim sMsg$ = "L'exploration de la source ODBC a été effectuée avec succès !"
            If Me.m_bCopierDonneesPressePapier Then sMsg &= " (cf. presse-papier)"
            MsgBox(sMsg, vbExclamation, m_sTitreMsg)
        End If

        bExplorerSourceODBC = True

    End Function

    Public Function bLireSourceODBC(
        Optional bRenvoyerContenu As Boolean = False,
        Optional bNePasFermerConnexion As Boolean = False) As Boolean

        ' Extraire les données de la requête SQL via la source ODBC 
        '  indiquée par le fichier .dsn

        ' Pour manipuler des grandes quantités de chaînes, 
        '  StringBuilder est beaucoup plus rapide que String
        Dim sbContenu As StringBuilder = Nothing
        Dim sbLigne As StringBuilder = Nothing
        Dim bNoterResultat As Boolean = False
        If bRenvoyerContenu Or Me.m_bCopierDonneesPressePapier Then
            bNoterResultat = True
            sbContenu = New StringBuilder
            sbLigne = New StringBuilder
        End If

        Dim sListeSQL$ = ""
        Dim sContenuDSN$ = ""
        If Not bLireSQL(sListeSQL, sContenuDSN,
            bNoterResultat, sbContenu) Then
            Me.AfficherMessage("Erreur !")
            Return False
        End If

        ' On initialise à Nothing pour éviter les avertissements intempestifs de VB8
        Dim oRq As ADODB.Recordset = Nothing
        Dim bConnOuverte As Boolean, bRqOuverte As Boolean
        Dim asSQL$() = sListeSQL.Split(CChar(";"))
        Dim iNbSQL% = 0
        Dim sSQL$ = ""

        Me.m_bRemplacerSepDecRequis = False
        Me.m_sSepDecimal = ""
        If Me.m_bRemplacerSepDec Then
            ' Remplacer , par . dans toutes les valeurs des champs 
            Me.m_sSepDecimal = Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator()
            If Me.m_sSepDecimal.Length > 0 AndAlso Me.m_sSepDecimal <> "." Then _
                Me.m_bRemplacerSepDecRequis = True
        End If

        If Not Me.m_bNePasInitAnnulation Then
            Me.m_bAnnuler = False
            Me.m_bErreursLecture = False
        End If

        Try

            Sablier()

            If IsNothing(Me.m_oConn) Then
                Me.m_oConn = New ADODB.Connection
                AfficherMessage("Ouverture de la connexion ODBC en cours...")
                If m_bModeEcriture Then
                    Me.m_oConn.Mode = ADODB.ConnectModeEnum.adModeReadWrite
                Else
                    Me.m_oConn.Mode = ADODB.ConnectModeEnum.adModeRead
                    'http://www.w3schools.com/ado/prop_mode.asp
                    'Allows others to open a connection with any permissions.
                    'Me.m_oConn.Mode = ADODB.ConnectModeEnum.adModeShareDenyNone
                End If
                Dim sConnexion$
                If Me.m_sChaineConnexionDirecte.Length = 0 Then
                    sConnexion = "FILEDSN=" & Me.m_sCheminDSN & ";"
                Else
                    sConnexion = Me.m_sChaineConnexionDirecte
                End If
                Me.m_oConn.Open(sConnexion)
            End If
            bConnOuverte = True

            oRq = New ADODB.Recordset

            Dim iNbRqMax% = asSQL.GetLength(0)
            Dim iNbChampsMax% = 0
            For Each sSQL In asSQL

                sSQL = sSQL.Trim
                If sSQL.Length = 0 Then Exit For

                ReDim Preserve Me.m_aoMetaTableau(iNbSQL)
                iNbSQL += 1

                Dim dDate As Date
                If bNoterResultat Then
                    sbContenu.Append(vbCrLf & vbCrLf & "SQL n°" & iNbSQL & " : " &
                        sSQL & vbCrLf & vbCrLf)
                    dDate = Now
                    AjouterTemps(sbContenu, "Heure début ouverture", dDate, dDate)
                End If

                If iNbRqMax >= 100 Then
                    If ((iNbSQL Mod 100 = 0) Or iNbSQL = iNbRqMax) And iNbSQL > 0 Then
                        Dim sAvancement$ =
                            "Exécution des requêtes en cours... : SQL n°" &
                            iNbSQL & "/" & iNbRqMax
                        AfficherMessage(sAvancement)
                        If Me.m_bAnnuler Then Exit For
                    End If
                Else
                    AfficherMessage("Exécution de la requête n°" & iNbSQL & " en cours...")
                    If Me.m_bAnnuler Then Exit For
                End If

                If Me.m_bODBCArriere Then
                    oRq.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                Else
                    oRq.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
                End If
                ' Par défaut : oRq.LockType = ADODB.LockTypeEnum.adLockReadOnly

                ' 10/04/2009 Tentatives de lecture, par ex. pour Excel
                Dim bOk As Boolean = False
                If m_iNbTentatives > 0 Then
                    For iNumTentative As Integer = 1 To m_iNbTentatives - 1
                        Try
                            oRq.Open(sSQL, Me.m_oConn)
                            bRqOuverte = True
                            bOk = True
                            Exit For
                        Catch
                            'Attendre(3000)
                            Threading.Thread.Sleep(3000) ' iDelaiMSec
                        End Try
                    Next
                End If
                If Not bOk Then
                    oRq.Open(sSQL, Me.m_oConn)
                    bRqOuverte = True
                End If

                Dim asTableau$(,) = Nothing ' Penser à réinitialiser le tableau
                Dim iNumEnreg%, i%, sValChamp$, iNbEnregAllouesAct%
                Dim oValChamp As Object

                Dim iNbChamps% = oRq.Fields.Count
                ' Cela peut arriver pour les requêtes en écriture, par exemple :
                '  UPDATE [Article$] SET [Article] = [Article] & '_Test'
                '  Dans ce cas, pensez à mettre ReadOnly=0 dans le fichier .dsn
                ' Ne pas faire oRq.Close() pour une requete insertion : cela plante !
                'If iNbChamps = 0 Then GoTo RequeteSuivante 
                If iNbChamps = 0 Then bRqOuverte = False : GoTo RequeteSuivante

                ' On peut noter les noms des champs systématiquement : pas couteux
                'If bNoterResultat Then
                Dim iNumSQL% = iNbSQL - 1
                ' Prendre tjrs le max du nbre de champs sur toutes les rq
                If iNbChamps > iNbChampsMax Then iNbChampsMax = iNbChamps
                If iNumSQL = 0 Then
                    ReDim Me.m_asChamps(iNbRqMax, iNbChampsMax)
                Else
                    ReDim Preserve Me.m_asChamps(iNbRqMax, iNbChampsMax)
                End If
                For i = 0 To iNbChamps - 1
                    Me.m_asChamps(iNumSQL, i) = oRq.Fields(i).Name
                Next i
                'End If

                If oRq.EOF Then
                    If bNoterResultat Then
                        AjouterTemps(sbContenu, "Heure début analyse  ", Now, dDate)
                        dDate = Now
                        AjouterEntete(sbContenu, iNbSQL - 1, iNbChamps)
                    End If
                    If Me.m_bPrompt Then _
                        MsgBox("La requête ne renvoie aucun enregistrement !",
                            MsgBoxStyle.Exclamation)
                    GoTo MemoriserTab_RqSuivante
                End If

                Dim iNbLignes% = -1

                If Me.m_bODBCArriere Then
                    ' Si l'ODBC ne supporte pas le retour en arrière MoveFirst, on obtient 
                    '  l'erreur -2147217884 (80040E24) avec la traduction en petit-nègre :
                    ' L'ensemble de lignes ne prend pas en charge les récupérations arrière
                    ' (Le jeu de données - RecordSet : l'objet requête - 
                    '  ne prend pas en charge le retour en arrière)
                    ' Les pilotes ODBC Access et Excel le supporte, on peut donc dimensionner
                    '  le tableau à l'avance (quoique le MoveLast ralenti au départ) :
                    AfficherMessage("Détermination du nombre de lignes...")
                    oRq.MoveLast()
                    iNbLignes = oRq.RecordCount
                    AfficherMessage("Retour au début du jeu de données...")
                    oRq.MoveFirst()
                    ReDim asTableau(iNbChamps - 1, iNbLignes - 1)
                Else
                    iNbLignes = 0
                    ' Bug corrigé : attendre d'avoir au moins un enregistrement
                    '  sinon on ne pourra pas distinguer entre 0 et 1 enregistrement
                    'ReDim asTableau(iNbChamps - 1, 0)
                End If

                ' On peut optimiser la lecture, mais de toute façon se sera long en ODBC
                '  GetString est surtout utile conjointement avec OWC
                ' (test réalisé : beaucoup plus rapide pour lire un fichier Excel en local,
                '  mais pas de gain constaté pour lire dans un PGI sur le réseau, 
                '  et on n'a plus l'avancement en temps réel)
                If Me.m_bLireToutDUnBloc Or Me.m_bLireToutDUnBlocRapide Then
                    If bNoterResultat Then
                        'AjouterTemps(sbContenu, "Heure début lecture  ", dDate, dDate)
                        AjouterTemps(sbContenu, "Heure début lecture  ", Now, dDate) ' 08/11/2007
                        dDate = Now
                        AjouterEntete(sbContenu, iNbSQL - 1, iNbChamps) ' 13/04/2008
                    End If
                    AfficherMessage("SQL n°" & iNbSQL &
                        " : Lecture des données d'un seul bloc...")
                    If bDebug Then Threading.Thread.Sleep(500)

                    ' Avec un délimiteur ; on peut afficher la ligne directement,
                    '  mais on ne traite pas les champs et il ne faut pas que
                    '  le signe ; se trouve dans le contenu d'un champ texte
                    If Me.m_bLireToutDUnBlocRapide Then
                        Const sDelimiteurColonnesRapide$ = ";"
                        Const sDelimiteurLignesRapide$ = vbCrLf ' 13/04/2008
                        ' 13/04/2008 : m_bLireToutDUnBlocRapide incompatible avec
                        '  multi-rq, sauf si les rq sont de même structure
                        'Me.m_sbLignes = New StringBuilder( _
                        '    oRq.GetString(, , sDelimiteurColonnesRapide))
                        Dim sb As New StringBuilder(
                            oRq.GetString(, ,
                                sDelimiteurColonnesRapide, sDelimiteurLignesRapide))
                        If bNoterResultat Then sbContenu.Append(sb)
                        If IsNothing(Me.m_sbLignes) Then
                            Me.m_sbLignes = sb
                        Else
                            Me.m_sbLignes.Append(sb)
                        End If
                        ' On laisse le tableau vide, on ne renvoi que Me.m_sLignes 
                        GoTo MemoriserTab_RqSuivante
                    End If

                    Const sDelimiteurColonnes$ = vbTab ' ";"
                    Dim asLignes$() = oRq.GetString(, ,
                        sDelimiteurColonnes).Split(CChar(vbCr))
                    If bNoterResultat Then
                        AjouterTemps(sbContenu, "Heure début analyse  ", Now, dDate)
                        dDate = Now
                        AjouterEntete(sbContenu, iNbSQL - 1, iNbChamps)
                    End If
                    AfficherMessage("SQL n°" & iNbSQL &
                        " : Analyse des données en cours...")
                    If bDebug Then Threading.Thread.Sleep(500)
                    Dim sLigne$
                    iNumEnreg = 0
                    For Each sLigne In asLignes
                        If sLigne.Length = 0 Then GoTo LigneSuivante
                        Dim asChamps$() = sLigne.Split(CChar(sDelimiteurColonnes))
                        If iNumEnreg = 0 Then
                            iNbLignes = asLignes.GetLength(0)
                            iNbChamps = asChamps.GetLength(0)
                            ReDim asTableau(iNbChamps - 1, iNbLignes - 1)
                        End If
                        Dim sValChamp0$
                        Dim iNumChamp% = 0
                        If bNoterResultat Then sbLigne.Length = 0
                        For Each sValChamp0 In asChamps
                            If sValChamp0.Length > 0 Then
                                TraiterValChamp(sValChamp0)
                            End If

                            ' 19/09/2010 Vérification du dépassement de colonnes
                            If iNumChamp >= iNbChampsMax Then
                                ' Le contenu du champ contient le séparateur : bug
                                'Debug.WriteLine("!")
                            Else
                                asTableau(iNumChamp, iNumEnreg) = sValChamp0
                            End If

                            If bNoterResultat Then
                                sbLigne.Append(sValChamp0)
                                If iNumChamp < iNbChamps - 1 Then sbLigne.Append(";")
                            End If
                            iNumChamp += 1
                        Next sValChamp0
                        If bNoterResultat Then
                            sbContenu.Append(sbLigne)
                            sbContenu.Append(vbCrLf)
                        End If
                        iNumEnreg += 1
LigneSuivante:
                    Next sLigne

                    GoTo MemoriserTab_RqSuivante
                End If

                ' Autre idée : DataAdaptater.Fill(DataTable) en une instruction 
                ' (méta-tableau de DataTable), mais on n'aura plus l'avancement 
                ' (on peut faire une boucle seulement pour débug)

                If bNoterResultat Then
                    'AjouterTemps(sbContenu, "Heure début lecture  ", dDate, dDate)
                    AjouterTemps(sbContenu, "Heure début lecture  ", Now, dDate) ' 08/11/2007
                    dDate = Now
                    AjouterEntete(sbContenu, iNbSQL - 1, iNbChamps)
                End If

                iNumEnreg = 0 : iNbEnregAllouesAct = 0
                While Not oRq.EOF

                    If (iNumEnreg Mod 100 = 0) And iNumEnreg > 0 Then
                        Dim sAvancement$ =
                            "Lecture de la source ODBC en cours... : SQL n°" &
                            iNbSQL & " : enreg. n°" & iNumEnreg + 1
                        If Me.m_bODBCArriere Then sAvancement &= "/" & iNbLignes
                        AfficherMessage(sAvancement)
                        ' Interrompre la requête en cours
                        If Me.m_bAnnuler Then Exit While
                    End If

                    If bNoterResultat Then sbLigne.Length = 0
                    If Not Me.m_bODBCArriere Then
                        ' Bug corrigé : attendre le premier enregistrement 
                        '  pour commencer à dimensionner le tableau : ReDim
                        If iNumEnreg = 0 Then
                            'ReDim asTableau(iNbChamps - 1, iNumEnreg)
                            ' Première allocation
                            iNbEnregAllouesAct = Me.m_iNbEnregAlloues
                            ReDim asTableau(iNbChamps - 1, iNbEnregAllouesAct - 1)
                        ElseIf iNumEnreg > iNbEnregAllouesAct - 1 Then
                            ' Redim ne peut changer que la dimension la plus à droite : iNbLignes
                            'ReDim Preserve asTableau(iNbChamps - 1, iNumEnreg)
                            ' Allocations suivantes
                            iNbEnregAllouesAct += Me.m_iNbEnregAlloues
                            ReDim Preserve asTableau(iNbChamps - 1, iNbEnregAllouesAct - 1)
                        End If
                    End If

                    For i = 0 To iNbChamps - 1

                        oValChamp = Nothing
                        sValChamp = ""
                        Try
                            oValChamp = oRq.Fields(i).Value
                            If Not IsDBNull(oValChamp) Then
                                ' Attention : La conversion ToString utilise le format
                                '  en vigueur dans les paramètres régionaux de Windows
                                '  par exemple pour le séparateur décimal
                                sValChamp = oValChamp.ToString
                            End If
                        Catch ex As Exception
                            Me.m_bErreursLecture = True
                            sValChamp = Me.m_sValErreur
                            'Dim s$ = ex.ToString
                            ' Une date du type 30/11/1899 provoque l'erreur suivante
                            '  pourtant IsDate("30/11/1899") est vrai
                            '  et une table Access liée sur cette source renvoie bien 
                            '  une vrai date 30/11/1899
                            ' Run-Time error '-2147217887 (80040E21)'
                            ' Multi-step OLE DB operation generated errors.
                            ' Une opération OLE-DB en plusieurs étapes a généré des erreurs.
                            ' Vérifiez chaque valeur d'état OLE-DB disponible. 
                            ' Aucun travail n'a été effectué.
                            'AfficherErreursADO(oConn)
                            'Exit Function
                        End Try

                        If sValChamp.Length > 0 Then
                            TraiterValChamp(sValChamp)
                        End If

                        If bNoterResultat Then
                            sbLigne.Append(sValChamp)
                            If i < iNbChamps - 1 Then sbLigne.Append(";")
                        End If

                        asTableau(i, iNumEnreg) = sValChamp

                    Next i

                    If bNoterResultat Then
                        sbContenu.Append(sbLigne)
                        sbContenu.Append(vbCrLf)
                    End If

                    oRq.MoveNext()
                    iNumEnreg += 1

                End While

                ' Avec Me.m_bInterrompreSeulementRqEnCours = True, on peut annuler une requête 
                '  mais poursuivre avec les autres, s'il y en a plusieurs
                If Me.m_bInterrompreSeulementRqEnCours Then
                    Me.m_bAnnuler = False
                Else
                    If Me.m_bAnnuler Then
                        sbContenu.Append(
                            "(interruption de l'utilisateur)").Append(vbCrLf)
                        Return False
                    End If
                End If

MemoriserTab_RqSuivante:
                ' Réduire la taille allouée du tableau à la taille effective
                If Me.m_iNbEnregAlloues > 1 AndAlso Not IsNothing(asTableau) Then
                    If asTableau.GetUpperBound(1) >= iNumEnreg Then
                        ReDim Preserve asTableau(iNbChamps - 1, iNumEnreg - 1)
                    End If
                End If

                ' Stocker le tableau dans le méta-tableau (tableau de tableaux de string)
                Me.m_aoMetaTableau(iNbSQL - 1) = asTableau
                If bNoterResultat Then
                    AjouterTemps(sbContenu, "Heure fin   analyse  ", Now, dDate)
                    dDate = Now
                End If

RequeteSuivante:
                If bRqOuverte Then oRq.Close() : bRqOuverte = False
            Next sSQL

        Catch ex As Exception

            Sablier(bDesactiver:=True)
            ' Si l'erreur a lieu lors de l'ouverture de la connexion
            '  afficher la liste des SQL
            If sSQL.Length = 0 Then
                sSQL = sListeSQL
                If sSQL.Length > 80 Then sSQL = sSQL.Substring(0, 80) & "..."
            End If
            Dim sMsg$ = "SQL : " & sSQL
            If Me.m_sChaineConnexionDirecte.Length = 0 Then
                sMsg &= vbCrLf & "Dsn : " & Me.m_sCheminDSN
            Else
                sMsg &= vbCrLf & "Chaîne de connexion : " & Me.m_sChaineConnexionDirecte
            End If
            Dim sDetailMsgErr$ = ""
            ' Ne pas copier l'erreur dans le presse-papier maintenant 
            '  car on va le faire dans le Finally

            Dim sMsgErrFinal$, sMsgErrADO$, sDetail$
            If bConnOuverte Then
                sDetail = "Certains champs sont peut-être introuvables, ou bien :"
            Else
                sDetail = "Erreur lors de l'ouverture de la connexion "
                If sContenuDSN.Length > 0 Then
                    sDetail &= "'" & sLireNomPiloteODBC(sContenuDSN) & "' :"
                Else
                    sDetail &= ":"
                End If
            End If
            sMsgErrFinal = "" : sMsgErrADO = ""
            AfficherMsgErreur2(ex, "bLireSourceODBC", sMsg, sDetail,
                bCopierMsgPressePapier:=False, sMsgErrFinal:=sMsgErrFinal)
            If bNoterResultat Then sbContenu.Append(vbCrLf & sMsgErrFinal & vbCrLf)
            AfficherErreursADO(Me.m_oConn, sMsgErrADO)
            If bNoterResultat Then sbContenu.Append(sMsgErrADO & vbCrLf)
            Me.AfficherMessage("Erreur !")
            Return False

        Finally

            Sablier(bDesactiver:=True)
            If bRqOuverte And Not IsNothing(oRq) Then _
                oRq.Close() : bRqOuverte = False
            If Not bNePasFermerConnexion Then
                ' Connexion ADODB et non OleDb
                If bConnOuverte Then Me.m_oConn.Close() : bConnOuverte = False
                Me.m_oConn = Nothing
            End If

            ' Copier les informations dans le presse-papier (utile pour le debogage)
            If Me.m_bCopierDonneesPressePapier Then _
                CopierPressePapier(sbContenu.ToString)
            ' Dans le cas de plusieurs accès ODBC, 
            '  on peut avoir besoin de mémoriser tous les contenus successifs
            If bRenvoyerContenu Then
                ' Autre syntaxe possible (pour éviter & vbCrLf & vbCrLf)
                sbContenu.Append(vbCrLf).Append(vbCrLf).Append(vbCrLf)
                If IsNothing(Me.m_sbContenuRetour) Then _
                    Me.m_sbContenuRetour = New StringBuilder
                Me.m_sbContenuRetour.Append(sbContenu)
            End If

        End Try

        Me.AfficherMessage("Opération terminée.")
        If Me.m_bPrompt Then
            Dim sMsg$ = "La lecture de la source ODBC a été effectuée avec succès !"
            If Me.m_bCopierDonneesPressePapier Then sMsg &= " (cf. presse-papier)"
            MsgBox(sMsg, vbExclamation, m_sTitreMsg)
        End If

        bLireSourceODBC = True

    End Function

    Private Sub TraiterValChamp(ByRef sValChamp$)

        ' Traiter la valeur des champs au cas où
        If Me.m_bRemplacerSepDecRequis Then
            ' Quel que soit le séparateur décimal, le convertir en .
            '  pour pouvoir convertir les nombres en réels via Val()
            ' IsNumeric dépend du séparateur régional, mais il est très lent
            ' Voir dans la doc : Notes sur la conversion en nombre réel
            Dim bRemp As Boolean = True
            If Me.m_bRemplacerSepDecNumSeul Then
                If Not IsNumeric(sValChamp) Then bRemp = False
            End If
            If bRemp Then sValChamp = sValChamp.Replace(Me.m_sSepDecimal, ".")
        End If
        If Me.m_bEnleverEspacesFin Then _
            sValChamp = sValChamp.TrimEnd ' = RTrim
        If Me.m_bRemplacerVraiFaux Then
            Dim sValChampMin$ = sValChamp.ToLower
            If sValChampMin = "faux" OrElse sValChampMin = "false" Then _
                sValChamp = Me.m_sValFaux
            If sValChampMin = "vrai" OrElse sValChampMin = "true" Then _
                sValChamp = Me.m_sValVrai
        End If

    End Sub

    Private Sub AjouterTemps(ByRef sbContenu As StringBuilder,
            sTexte$, dDate2 As Date, dDate1 As Date)
        If Not Me.m_bAjouterChronoDebug Then Exit Sub
        sbContenu.Append(sTexte).Append(" : ")
        sbContenu.Append(Now.ToLongTimeString)
        If dDate2 > dDate1 Then
            sbContenu.Append(" : ")
            Dim tsDelai As System.TimeSpan = dDate2.Subtract(dDate1)
            If tsDelai.TotalMinutes >= 1 Then _
                sbContenu.Append(tsDelai.TotalMinutes.ToString("0")).Append(" mn : ")
            sbContenu.Append(tsDelai.TotalSeconds).Append(" sec.")
        End If
        sbContenu.Append(vbCrLf)
    End Sub

    Private Sub AjouterEntete(ByRef sbContenu As StringBuilder,
            iNumSQL%, iNbChamps%)
        Dim i%
        For i = 0 To iNbChamps - 1
            sbContenu.Append(Me.m_asChamps(iNumSQL, i))
            If i < iNbChamps - 1 Then sbContenu.Append(";")
        Next i
        sbContenu.Append(vbCrLf)
    End Sub

#End Region

#Region "Creation d'un fichier DSN"

    Private Function bCreerFichiersDsnEtSQLODBCDefaut() As Boolean

        ' Créer un fichier DSN ODBC par défaut en fonction des sources
        '  possibles trouvées, ainsi que les requêtes SQL correspondantes

        ' Chemins des sources ODBC possibles
        ' Autres fichiers DSN ODBC : www.prosygma.com/odbc-dsn.htm
        Dim sListeSrcPossibles$ = ""
        If Me.m_sCheminSrcExcel.Length > 0 Then _
            sListeSrcPossibles &= Me.m_sCheminSrcExcel & vbLf
        If Me.m_sCheminSrcAccess.Length > 0 Then _
            sListeSrcPossibles &= Me.m_sCheminSrcAccess & vbLf
        If Me.m_sCheminSrcOmnis.Length > 0 Then _
            sListeSrcPossibles &= Me.m_sCheminSrcOmnis

        If Me.m_sSQLNavisionDef.Length > 0 And
           Me.m_sCompteSociete.Length > 0 And Me.m_sNomServeur.Length > 0 Then
            If Not bCreerFichierDsnODBC(sTypeODBCNavision, Me.m_sCheminDSN,
                Me.m_sCheminSQL, "", Me.m_sSQLNavisionDef,
                Me.m_sCompteUtilisateur, Me.m_sMotDePasse,
                Me.m_sCompteSociete, Me.m_sNomServeur) Then _
                Return False
        ElseIf Me.m_sSQLDB2Def.Length > 0 And
            Me.m_sCompteSociete.Length > 0 And Me.m_sNomServeur.Length > 0 Then
            If Not bCreerFichierDsnODBC(sTypeODBCDB2, Me.m_sCheminDSN,
                Me.m_sCheminSQL, "", Me.m_sSQLDB2Def,
                Me.m_sCompteUtilisateur, Me.m_sMotDePasse,
                Me.m_sCompteSociete, Me.m_sNomServeur) Then _
                Return False
        ElseIf Me.m_sCheminSrcExcel.Length > 0 AndAlso
            bFichierExiste(Me.m_sCheminSrcExcel) Then
            If Not bCreerFichierDsnODBC(sTypeODBCExcel, Me.m_sCheminDSN,
                Me.m_sCheminSQL, Me.m_sCheminSrcExcel, Me.m_sSQLExcelDef) Then _
                Return False
        ElseIf Me.m_sCheminSrcAccess.Length > 0 AndAlso
            bFichierExiste(Me.m_sCheminSrcAccess) Then
            If Not bCreerFichierDsnODBC(sTypeODBCAccess, Me.m_sCheminDSN,
                Me.m_sCheminSQL, Me.m_sCheminSrcAccess, Me.m_sSQLAccessDef) Then _
                Return False
        ElseIf Me.m_sCheminSrcOmnis.Length > 0 AndAlso
            bFichierExiste(Me.m_sCheminSrcOmnis) Then
            If Not bCreerFichierDsnODBC(sTypeODBCOmnis, Me.m_sCheminDSN,
                Me.m_sCheminSQL, Me.m_sCheminSrcOmnis, Me.m_sSQLOmnisDef,
                Me.m_sCompteUtilisateur, Me.m_sMotDePasse) Then _
                Return False
        Else
            Dim sMsg$ = "Aucune source ODBC possible n'a été trouvée pour créer un fichier DSN !"
            If sListeSrcPossibles.Length > 0 Then _
                sMsg &= vbLf & "Liste des sources possibles : " & vbLf & sListeSrcPossibles
            MsgBox(sMsg, MsgBoxStyle.Critical, m_sTitreMsg)
            Return False
        End If
        Return True

    End Function

    Private Function bCreerFichierDsnODBC(sTypeODBC$, sCheminDsn$,
            sCheminSQL$, sFichierSrc$, sSQL$,
            Optional sCompteUtilisateur$ = "",
            Optional sMotDePasse$ = "",
            Optional sCompteSociete$ = "",
            Optional sNomServeur$ = "") As Boolean

        ' Créer un fichier DSN ODBC par défaut en fonction des sources possibles trouvées
        '  ainsi que les requêtes SQL correspondantes

        Dim sSource$ = sFichierSrc
        Dim sDossierSrc$ = ""
        If sFichierSrc.Length > 0 Then _
            sDossierSrc = IO.Path.GetDirectoryName(sFichierSrc)

        Dim sb As New StringBuilder

        ' Autres fichiers DSN ODBC : www.prosygma.com/odbc-dsn.htm
        Select Case sTypeODBC
            Case sTypeODBCExcel
                sb.Append("[ODBC]" & vbCrLf)
                sb.Append("DRIVER=Microsoft Excel Driver (*.xls)" & vbCrLf)
                sb.Append("UID=admin" & vbCrLf)
                sb.Append("UserCommitSync=Yes" & vbCrLf)
                sb.Append("Threads=3" & vbCrLf)
                sb.Append("SafeTransactions=0" & vbCrLf)
                If Me.m_bModeEcriture Then
                    sb.Append("ReadOnly=0" & vbCrLf)
                Else
                    sb.Append("ReadOnly=1" & vbCrLf)
                End If
                sb.Append("PageTimeout=5" & vbCrLf)

                ' En pratique MaxScanRows n'est pas utilisé dans le fichier DSN !
                ' Seule la clé TypeGuessRows de la base de registre :
                '  HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel
                '  permet de prendre en compte un plus grand nombre de lignes
                '  pour déterminer automatiquement le type du champ,
                '  ce qui est nécessaire si les n premières occurrences 
                '  du champs sont vides dans la feuille Excel :
                ' www.dicks-blog.com/archives/2004/06/03/external-data-mixed-data-types/
                ' Utilisez la fonction VerifierConfigODBCExcel() pour vérifier sa valeur
                '  sauf si vous travaillez avec Excel 2003, qui fonctionne bien 
                '  dans tous les cas, car il utilise une dll plus efficace :
                '  Microsoft Access Expression Builder : 
                '  C:\Program Files\Microsoft Office\Office11\msaexp30.dll (11.0.6561.0) 
                '  la dll par défaut étant : Microsoft Jet Excel Isam :
                '  C:\Windows\System32\msexcl40.dll (4.0.8618.0) 

                sb.Append("MaxScanRows=8" & vbCrLf)

                sb.Append("MaxBufferSize=2048" & vbCrLf)
                sb.Append("FIL=excel 8.0" & vbCrLf)
                sb.Append("DriverId=790" & vbCrLf)
                sb.Append("DefaultDir=" & sDossierSrc & vbCrLf)
                sb.Append("DBQ=" & sFichierSrc & vbCrLf)
            ' On peut aussi indiquer un chemin relatif avec .
            ' Ex.: DefaultDir=.\SourcesODBC\SourceODBC_MSExcel
            '      DBQ=.\SourcesODBC\SourceODBC_MSExcel\XLDB.xls

            Case sTypeODBCAccess
                sb.Append("[ODBC]" & vbCrLf)
                sb.Append("DRIVER=Microsoft Access Driver (*.mdb)" & vbCrLf)
                sb.Append("UID=admin" & vbCrLf)
                sb.Append("UserCommitSync=Yes" & vbCrLf)
                sb.Append("Threads=3" & vbCrLf)
                sb.Append("SafeTransactions=0" & vbCrLf)
                sb.Append("PageTimeout=5" & vbCrLf)
                sb.Append("MaxScanRows=8" & vbCrLf)
                sb.Append("MaxBufferSize=2048" & vbCrLf)
                sb.Append("FIL=MS Access" & vbCrLf)
                sb.Append("DriverId=25" & vbCrLf)
                sb.Append("DefaultDir=" & sDossierSrc & vbCrLf)
                sb.Append("DBQ=" & sFichierSrc & vbCrLf)

            Case sTypeODBCOmnis
                ' Pilote : www.omnis.net/downloads/odbc/win32/Omnis%20ODBC%20Driver.exe
                sb.Append("[ODBC]" & vbCrLf)
                sb.Append("DRIVER=OMNIS ODBC Driver" & vbCrLf)
                sb.Append("UID=admin" & vbCrLf)
                sb.Append("Password=" & sMotDePasse & vbCrLf)
                sb.Append("Username=" & sCompteUtilisateur & vbCrLf)
                sb.Append("DataFilePath=" & sFichierSrc & vbCrLf)

            Case sTypeODBCNavision
                sSource = sCompteSociete
                ' Doc sur le pilote C-Odbc : 
                '  http://www.comsolag.de/old/pdf/Handbuch/W1/w1w1codbc.pdf
                sb.Append("[ODBC]" & vbCrLf)
                sb.Append("DRIVER=C/ODBC 32 bit" & vbCrLf)
                sb.Append("UID=" & sCompteUtilisateur & vbCrLf)
                sb.Append("SERVER=N" & vbCrLf) ' Non documenté !
                sb.Append("CN=" & sCompteSociete & vbCrLf) ' The account/company to open
                sb.Append("RD=No" & vbCrLf) ' Non documenté !

                ' ML indique la langue utilisée : 1033 pour l'anglais (USA),
                '  1036 pour le français. Les tables et les champs de la requête SQL 
                '  doivent être dans la langue choisie. Il est apparemment impossible 
                '  de faire passer les accents en français, donc laisser 1033.
                sb.Append("ML=1033" & vbCrLf)
                ' CD Specifies whether the connection supports closing date.
                sb.Append("CD=No" & vbCrLf)
                ' BE Specifies whether BLOB fields should be visible from ODBC.
                sb.Append("BE=Yes" & vbCrLf)
                ' CC Specifies whether the commit cache should be used.
                sb.Append("CC=Yes" & vbCrLf)
                ' RO Specifies whether access to the Microsoft Business Solutions  
                '  database should be read-only.
                sb.Append("RO=No" & vbCrLf)
                sb.Append("QTYesNo=Yes" & vbCrLf) ' Enables or disables query time-out
                ' IT Specify the way identifiers are returned to an external application
                sb.Append("IT=All Except Space" & vbCrLf)
                ' OPT Specifies how the contents of a Navision option field are 
                '  transferred to an application.
                sb.Append("OPT=Text" & vbCrLf)
                ' PPath : The name of the folder where the program files are located.
                Dim sLecteur$ = IO.Path.GetPathRoot(Environment.SystemDirectory) ' Ex.: C:\
                sb.Append("PPath=" & sLecteur &
                "Program Files\Microsoft Business Solutions-Navision\Client" & vbCrLf)
                ' NType : The name of the network protocol module (tcp or netb).
                sb.Append("NType=tcp" & vbCrLf)
                sb.Append("SName=" & sNomServeur & vbCrLf) ' The name of the server host computer.
                ' CSF Specifies whether the driver operates as a client in a 
                '  client/server environment or as a stand-alone.
                sb.Append("CSF=Yes" & vbCrLf)
                ' Attention : il n'est pas possible de crypter le mot de passe avec ce pilote :
                '  La doc recommande de créer un compte utilisateur spécifique avec les seuls
                '  droits requis pour l'exécution de la requête.
                sb.Append("PWD=" & sMotDePasse & vbCrLf)

            Case sTypeODBCDB2 ' DB2 = iSeries d'IBM (anciennement AS/400)
                sSource = sCompteSociete
                sb.Append("[ODBC]" & vbCrLf)
                sb.Append("DRIVER=Client Access ODBC Driver (32-bit)" & vbCrLf)

                sb.Append("UID=" & sCompteUtilisateur & vbCrLf) ' ou CA400 par défaut
                ' Pour DB2, il n'y a pas de mot de passe, il faut laisser une connexion 
                '  ouverte et le pilote ODBC va réutiliser cette connexion.
                '  voir la doc avec SIGNON=1
                '  (si la connexion n'est pas ouverte, le système devrait ouvrir une 
                '   boite de dialogue pour saisir le mot de passe, mais je n'ai pas 
                '   réussi à le faire marcher ainsi)

                sb.Append("DEBUG=64" & vbCrLf)
                sb.Append("SIGNON=1" & vbCrLf)
                sb.Append("LIBVIEW=1" & vbCrLf)
                sb.Append("TRANSLATE=1" & vbCrLf)
                sb.Append("NAM=1" & vbCrLf)
                sb.Append("DESC=Source de données ODBC iSeries Access for Windows" & vbCrLf)
                sb.Append("SQDIAGCODE=" & vbCrLf)
                sb.Append("DATABASE=" & vbCrLf)
                sb.Append("QAQQINILIB=" & vbCrLf)
                sb.Append("PKG=QGPL/DEFAULT(IBM),2,0,1,0,512" & vbCrLf)
                Dim sLecteur$ = IO.Path.GetPathRoot(Environment.SystemDirectory) ' Ex.: C:\
                Dim sUtilisateur$ = Environment.UserName
                ' A vérifier : sUtilisateur = 'Utilisateur' littéralement ?
                sb.Append("TRACEFILENAME=" & sLecteur &
                "Documents and Settings\" & sUtilisateur &
                "\Mes documents\IBM\Client Access\Maintenance\Fichiers trace" & vbCrLf)
                sb.Append("SORTTABLE=" & vbCrLf)
                sb.Append("LANGUAGEID=ENU" & vbCrLf)
                sb.Append("XLATEDLL=" & vbCrLf)
                sb.Append("DFTPKGLIB=QGPL" & vbCrLf)

                ' A vérifier : ici on peut indiquer une autre librairie 
                '  que la librairie QGPL par défaut
                '  ce qui évite d'avoir à préfixer les noms de table 
                '  par la librairie dans les requêtes, le cas échéant
                sb.Append("DBQ=QGPL" & vbCrLf)

                sb.Append("SYSTEM=" & sNomServeur & vbCrLf) ' autre poss.: Adresse IP

        End Select

        If Not bEcrireFichier(sCheminDsn, sb) Then Return False

        ' On peut ne pas avoir besoin d'un fichier de requête SQL,
        '  si on les crée à la volée
        If sCheminSQL.Length > 0 And sSQL.Length > 0 Then
            If bFichierExiste(sCheminSQL) Then _
                If Not bRenommerFichier(sCheminSQL, sCheminSQL & ".bak") Then Return False
            If Not bEcrireFichier(sCheminSQL, sSQL) Then Return False
        End If

        If Me.m_bPrompt Then _
            MsgBox("Le fichier DSN pour la source ODBC " & sTypeODBC & " : " & vbLf &
                sSource & vbLf & "a été créé avec les chemins en local :" & vbLf &
                sCheminDsn, vbExclamation, m_sTitreMsg)

        Return True

    End Function

    Public Function bVerifierCheminODBC(sChampBD$, sContenuDSN$,
            Optional bDossier As Boolean = False) As Boolean

        ' Vérifier la présence de la source ODBC si le fichier DSN existe déjà

        Dim sContenuDSNMin$ = sContenuDSN.ToLower
        sChampBD = sChampBD.ToLower
        Dim iPosDeb% = sContenuDSNMin.IndexOf(sChampBD)
        Dim sCheminBd$ = ""
        If iPosDeb > -1 Then
            Dim iPosFin% = sContenuDSNMin.IndexOf(vbLf, iPosDeb + sChampBD.Length)
            If iPosFin > -1 Then
                sCheminBd = sContenuDSN.Substring(
                    iPosDeb + sChampBD.Length, iPosFin - 1 - iPosDeb - sChampBD.Length)
            Else
                sCheminBd = sContenuDSN.Substring(iPosDeb + sChampBD.Length)
            End If
            If sCheminBd.Length = 0 Then
                MsgBox("Le chemin indiqué dans le fichier DSN pour " & sChampBD &
                    " est vide !", MsgBoxStyle.Critical, m_sTitreMsg)
                Return False
            End If
            If Not bCheminFichierProbable(sCheminBd) Then
                ' Si le chemin ne correspond pas à un vrai chemin
                '  alors ne pas chercher à vérifier la présence du fichier
                '  poursuivre sans erreur
                Return True
            End If
            Dim sDebutLigneChamp$ = sContenuDSNMin.Substring(
                iPosDeb - 3, sChampBD.Length)
            If sDebutLigneChamp.IndexOf(";") > -1 Then
                ' Si le chemin indiqué est en commentaire
                '  alors ignorer la ligne, poursuivre sans erreur
                Return True
            End If
        End If
        bVerifierCheminODBC = True
        If sCheminBd.Length > 0 Then
            If bDossier Then
                Return bDossierExiste(sCheminBd, bPrompt:=True)
            Else
                Return bFichierExiste(sCheminBd, bPrompt:=True)
            End If
        End If

    End Function

    Public Function sLireNomPiloteODBC$(sContenuDSN$)

        ' Vérifier la présence de la source ODBC si le fichier DSN existe déjà

        Dim sContenuDSNMin$ = sContenuDSN.ToLower
        Dim sChampPilote$ = "driver="
        Dim iPosDeb% = sContenuDSNMin.IndexOf(sChampPilote)
        Dim sNomPilote$ = ""
        If iPosDeb > -1 Then
            Dim iPosFin% = sContenuDSNMin.IndexOf(vbLf, iPosDeb + sChampPilote.Length)
            If iPosFin > -1 Then
                sNomPilote = sContenuDSN.Substring(
                    iPosDeb + sChampPilote.Length, iPosFin - 1 - iPosDeb - sChampPilote.Length)
            Else
                sNomPilote = sContenuDSN.Substring(iPosDeb + sChampPilote.Length)
            End If
        End If
        sLireNomPiloteODBC = sNomPilote

    End Function

#End Region

End Class