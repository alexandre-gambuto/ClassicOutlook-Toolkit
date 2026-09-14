VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfig 
   Caption         =   "Configuration de l'analyse"
   ClientHeight    =   9940.001
   ClientLeft      =   20
   ClientTop       =   100
   ClientWidth     =   13360
   OleObjectBlob   =   "frmConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' frmConfig - CODE V23
' Configuration de l'analyse des emails
' Avec nouvelle option : Analyse longueur des messages
' ============================================================================
Option Explicit

' --- Variables de module ---
Private m_MailboxConfigs As Object    ' Dictionary: storeName -> config (Dictionary)
Private m_Cancelled As Boolean
Private m_LastSelectedStore As String

' ============================================================================
' INITIALISATION
' ============================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrInit
    
    Debug.Print "[frmConfig] Initialize d�but..."
    
    Set m_MailboxConfigs = CreateObject("Scripting.Dictionary")
    m_Cancelled = True
    m_LastSelectedStore = ""
    
    Debug.Print "[frmConfig] Remplissage cboCouleur..."
    ' Remplir la liste des couleurs
    With cboCouleur
        .Clear
        .AddItem "Blue"
        .AddItem "Green"
        .AddItem "Orange"
        .AddItem "Red"
        .AddItem "Purple"
        .AddItem "Cyan"
        .AddItem "Pink"
        .AddItem "Yellow"
        .AddItem "Navy"
        .AddItem "Forest green"
        .AddItem "Brown"
        .AddItem "Gray"
        .AddItem "Turquoise"
        .AddItem "Coral"
        .AddItem "Plum"
        .ListIndex = 0
    End With
    
    Debug.Print "[frmConfig] LoadMailboxes..."
    ' Charger les bo�tes aux lettres
    Call LoadMailboxes
    
    Debug.Print "[frmConfig] Initialisation p�riode..."
    ' Initialiser p�riode
    optTout.Value = True
    cboDepuis.Enabled = False
    txtDateDebut.Enabled = False
    txtDateFin.Enabled = False
    Call LoadYearsCombo
    
    Debug.Print "[frmConfig] Initialisation options..."
    ' Initialiser options
    optTest.Value = True
    chkResolutionExchange.Value = True
    chkAnalyseDistincte.Value = False
    chkAnalyseLongueur.Value = False
    
    m_Cancelled = True
    
    Debug.Print "[frmConfig] Initialize termin� OK"
    Exit Sub
    
ErrInit:
    Debug.Print "[frmConfig] !!! ERREUR Initialize: " & Err.Number & " - " & Err.Description
    MsgBox "frmConfig initialization error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' CHARGEMENT DES BO�TES
' ============================================================================
Private Sub LoadMailboxes()
    Dim olNS As Object
    Dim store As Object
    Dim colorIndex As Long
    Dim storeName As String
    Dim emailUser As String
    Dim cfg As Object
    
    On Error Resume Next
    Set olNS = Application.GetNamespace("MAPI")
    
    lstBoites.Clear
    colorIndex = 0
    
    For Each store In olNS.Stores
        storeName = store.displayName
        
        If Len(storeName) > 0 Then
            lstBoites.AddItem storeName
            
            Set cfg = CreateObject("Scripting.Dictionary")
            
            ' Appliquer pr�-configuration si disponible, sinon extraction auto
            Call ApplyPresetConfig(storeName, cfg, colorIndex)
            
            m_MailboxConfigs.Add storeName, cfg
            colorIndex = colorIndex + 1
        End If
    Next store
    
    On Error GoTo 0
End Sub

' ============================================================================
' PRESET CONFIGURATIONS (per-mailbox shortcuts)
' ============================================================================
Private Sub ApplyPresetConfig(storeName As String, ByRef cfg As Object, colorIndex As Long)
    Dim lowerStore As String
    lowerStore = LCase(Trim(storeName))
    
    ' Hard-coded shortcuts (match on the mailbox store name).
    Select Case True
        ' --- Example preset (optional shortcut for a known mailbox) ----------
        ' Add one Case per mailbox you want pre-configured: match on the store
        ' name as shown in Outlook, and set the user email / domain / keyword.
        ' Anything not matched here falls through to Case Else (auto-extraction).
        Case InStr(1, lowerStore, "example", vbTextCompare) > 0
            cfg("EmailUser") = "first.last@example.com"
            cfg("Domain") = "example.com"
            cfg("Keyword") = "last"
            cfg("AddressMode") = False
            cfg("ColorIndex") = 0
            cfg("ColorRGB") = GetColorByIndex(0)
            
        Case Else
            ' Configuration par d�faut (extraction automatique)
            Dim emailUser As String
            emailUser = ExtractEmailFromStore(storeName)
            cfg("EmailUser") = emailUser
            cfg("Domain") = ExtractDomainFromEmail(emailUser)
            cfg("Keyword") = ExtractKeywordFromEmail(emailUser)
            cfg("AddressMode") = False
            cfg("ColorIndex") = colorIndex Mod 15
            cfg("ColorRGB") = GetColorByIndex(colorIndex Mod 15)
    End Select
End Sub

Private Function ExtractEmailFromStore(storeName As String) As String
    ' Si le nom de la bo�te contient @, c'est l'adresse email
    If InStr(storeName, "@") > 0 Then
        ExtractEmailFromStore = storeName
    Else
        ExtractEmailFromStore = ""
    End If
End Function

Private Function ExtractDomainFromEmail(email As String) As String
    Dim atPos As Long
    atPos = InStr(email, "@")
    If atPos > 0 Then
        ExtractDomainFromEmail = mid(email, atPos + 1)
    Else
        ExtractDomainFromEmail = ""
    End If
End Function

Private Function ExtractKeywordFromEmail(email As String) As String
    Dim atPos As Long
    Dim localPart As String
    Dim dotPos As Long
    
    atPos = InStr(email, "@")
    If atPos > 1 Then
        localPart = LCase(Left(email, atPos - 1))
        ' Prendre le dernier segment apr�s le point (g�n�ralement le pr�nom ou nom)
        dotPos = InStrRev(localPart, ".")
        If dotPos > 0 Then
            ExtractKeywordFromEmail = mid(localPart, dotPos + 1)
        Else
            ExtractKeywordFromEmail = localPart
        End If
    Else
        ExtractKeywordFromEmail = ""
    End If
End Function

Private Sub LoadYearsCombo()
    Dim y As Long
    cboDepuis.Clear
    For y = year(Now) To 2000 Step -1
        cboDepuis.AddItem y
    Next y
    If cboDepuis.ListCount > 0 Then cboDepuis.ListIndex = 0
End Sub

' ============================================================================
' GESTION DES COULEURS
' ============================================================================
Private Function GetColorByIndex(idx As Long) As Long
    Select Case idx
        Case 0: GetColorByIndex = RGB(68, 114, 196)   ' Bleu
        Case 1: GetColorByIndex = RGB(112, 173, 71)   ' Vert
        Case 2: GetColorByIndex = RGB(255, 152, 0)    ' Orange
        Case 3: GetColorByIndex = RGB(255, 0, 0)      ' Rouge
        Case 4: GetColorByIndex = RGB(128, 0, 128)    ' Violet
        Case 5: GetColorByIndex = RGB(0, 176, 183)    ' Cyan
        Case 6: GetColorByIndex = RGB(215, 45, 124)   ' Rose
        Case 7: GetColorByIndex = RGB(255, 192, 0)    ' Jaune
        Case 8: GetColorByIndex = RGB(0, 43, 89)      ' Bleu marine
        Case 9: GetColorByIndex = RGB(34, 139, 34)    ' Vert for�t
        Case 10: GetColorByIndex = RGB(139, 69, 19)   ' Marron
        Case 11: GetColorByIndex = RGB(128, 128, 128) ' Gris
        Case 12: GetColorByIndex = RGB(64, 224, 208)  ' Turquoise
        Case 13: GetColorByIndex = RGB(255, 127, 80)  ' Corail
        Case 14: GetColorByIndex = RGB(142, 69, 133)  ' Prune
        Case Else: GetColorByIndex = RGB(128, 128, 128)
    End Select
End Function

' ============================================================================
' �V�NEMENTS LISTE DES BO�TES
' ============================================================================
Private Sub lstBoites_Change()
    Dim selectedCount As Long
    Dim lastSelectedIndex As Long
    Dim storeName As String
    Dim cfg As Object
    Dim i As Long
    
    ' Compter les s�lections et trouver le dernier index s�lectionn�
    selectedCount = 0
    lastSelectedIndex = -1
    
    For i = 0 To lstBoites.ListCount - 1
        If lstBoites.Selected(i) Then
            selectedCount = selectedCount + 1
            lastSelectedIndex = i
        End If
    Next i
    
    ' Aucune s�lection : d�sactiver le panneau config
    If selectedCount = 0 Then
        fraConfigBoite.Enabled = False
        fraConfigBoite.caption = "Selected mailbox configuration"
        Call ClearConfigFields
        m_LastSelectedStore = ""
        Exit Sub
    End If
    
    ' Si l'�l�ment actuellement cliqu� est s�lectionn�, l'utiliser
    If lstBoites.ListIndex >= 0 And lstBoites.Selected(lstBoites.ListIndex) Then
        lastSelectedIndex = lstBoites.ListIndex
    End If
    
    storeName = lstBoites.list(lastSelectedIndex)
    
    ' Sauvegarder la config pr�c�dente si on change de bo�te
    If storeName <> m_LastSelectedStore Then
        If Len(m_LastSelectedStore) > 0 And m_MailboxConfigs.Exists(m_LastSelectedStore) Then
            Call SaveCurrentConfig(m_LastSelectedStore)
        End If
    Else
        ' M�me bo�te, pas besoin de recharger
        Exit Sub
    End If
    
    m_LastSelectedStore = storeName
    
    If Not m_MailboxConfigs.Exists(storeName) Then Exit Sub
    
    ' Activer le panneau et mettre � jour le titre
    fraConfigBoite.Enabled = True
    If Len(storeName) > 30 Then
        fraConfigBoite.caption = "Config: " & Left(storeName, 27) & "..."
    Else
        fraConfigBoite.caption = "Config: " & storeName
    End If
    
    ' Charger la config de la bo�te s�lectionn�e
    Set cfg = m_MailboxConfigs(storeName)
    
    txtEmailUser.text = cfg("EmailUser")
    txtDomaine.text = cfg("Domain")
    txtMotCle.text = cfg("Keyword")
    
    If cfg("AddressMode") Then
        optModeAdresse.Value = True
    Else
        optModeMotCle.Value = True
    End If
    
    cboCouleur.ListIndex = cfg("ColorIndex")
    lblCouleurApercu.BackColor = cfg("ColorRGB")
    lblResultatTest.caption = ""
    
    ' Charger DateMin si les contr�les existent
    On Error Resume Next
    Dim dateMinEnabled As Boolean
    dateMinEnabled = False
    If cfg.Exists("DateMinEnabled") Then dateMinEnabled = cfg("DateMinEnabled")
    
    ' Checkbox : coch� ou non (pas gris�)
    Me.controls("chkDateMin").Value = dateMinEnabled
    
    ' TextBox : activ�/d�sactiv� avec couleur appropri�e
    Me.controls("txtDateMin").Enabled = dateMinEnabled
    If dateMinEnabled Then
        Me.controls("txtDateMin").BackColor = &H80000005  ' Blanc (couleur fen�tre)
        If cfg.Exists("DateMin") Then
            If IsDate(cfg("DateMin")) Then
                Me.controls("txtDateMin").text = Format(cfg("DateMin"), "DD/MM/YYYY")
            Else
                Me.controls("txtDateMin").text = ""
            End If
        Else
            Me.controls("txtDateMin").text = ""
        End If
    Else
        Me.controls("txtDateMin").BackColor = &H8000000F  ' Gris (couleur bouton)
        Me.controls("txtDateMin").text = ""
    End If
    On Error GoTo 0
End Sub

Private Sub ClearConfigFields()
    txtEmailUser.text = ""
    txtDomaine.text = ""
    txtMotCle.text = ""
    optModeMotCle.Value = True
    cboCouleur.ListIndex = 0
    lblCouleurApercu.BackColor = GetColorByIndex(0)
    lblResultatTest.caption = ""
End Sub

Private Sub SaveCurrentConfig(storeName As String)
    If Not m_MailboxConfigs.Exists(storeName) Then Exit Sub
    
    Dim cfg As Object
    Set cfg = m_MailboxConfigs(storeName)
    
    cfg("EmailUser") = txtEmailUser.text
    cfg("Domain") = txtDomaine.text
    cfg("Keyword") = txtMotCle.text
    cfg("AddressMode") = optModeAdresse.Value
    cfg("ColorIndex") = cboCouleur.ListIndex
    cfg("ColorRGB") = GetColorByIndex(cboCouleur.ListIndex)
    
    ' Sauvegarder DateMin si les contr�les existent
    On Error Resume Next
    cfg("DateMinEnabled") = Me.controls("chkDateMin").Value
    If Me.controls("chkDateMin").Value Then
        Dim dateStr As String
        dateStr = Me.controls("txtDateMin").text
        If IsDate(dateStr) Then
            cfg("DateMin") = CDate(dateStr)
        End If
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' �V�NEMENTS CONFIGURATION BO�TE
' ============================================================================
Private Sub txtEmailUser_Change()
    ' Auto-remplir domaine et mot-cl�
    Dim email As String
    email = txtEmailUser.text
    
    If InStr(email, "@") > 0 Then
        txtDomaine.text = mid(email, InStr(email, "@") + 1)
        Dim localPart As String
        localPart = Left(email, InStr(email, "@") - 1)
        If InStr(localPart, ".") > 0 Then
            txtMotCle.text = Left(localPart, InStr(localPart, ".") - 1)
        Else
            txtMotCle.text = localPart
        End If
    End If
End Sub

Private Sub cboCouleur_Change()
    lblCouleurApercu.BackColor = GetColorByIndex(cboCouleur.ListIndex)
End Sub

Private Sub btnAppliquerConfig_Click()
    If Len(m_LastSelectedStore) > 0 Then
        Call SaveCurrentConfig(m_LastSelectedStore)
        lblResultatTest.caption = "Configuration applied."
        lblResultatTest.ForeColor = RGB(0, 128, 0)
    End If
End Sub

Private Sub btnTesterResolution_Click()
    If Len(m_LastSelectedStore) = 0 Then Exit Sub
    
    lblResultatTest.caption = "Testing..."
    lblResultatTest.ForeColor = RGB(0, 0, 0)
    DoEvents
    
    Dim olNS As Outlook.NameSpace
    Dim store As Outlook.store
    Dim rootFolder As Outlook.folder
    Dim testFolder As Outlook.folder
    Dim item As Object
    Dim testCount As Long, resolvedCount As Long
    
    Set olNS = Application.GetNamespace("MAPI")
    
    On Error Resume Next
    For Each store In olNS.Stores
        If store.displayName = m_LastSelectedStore Then
            Set rootFolder = store.GetRootFolder
            Set testFolder = GetInboxOrFirst(rootFolder)
            If Not testFolder Is Nothing Then
                testCount = 0
                resolvedCount = 0
                For Each item In testFolder.items
                    If TypeName(item) = "MailItem" Then
                        Dim mail As Outlook.mailItem
                        Set mail = item
                        If Not mail.Sender Is Nothing Then
                            testCount = testCount + 1
                            Dim resolved As String
                            resolved = TryResolveAddress(mail.Sender)
                            If InStr(resolved, "@") > 0 Then resolvedCount = resolvedCount + 1
                        End If
                        If testCount >= 20 Then Exit For
                    End If
                Next item
            End If
            Exit For
        End If
    Next store
    On Error GoTo 0
    
    If testCount > 0 Then
        Dim pct As Long
        pct = (resolvedCount * 100) \ testCount
        lblResultatTest.caption = "Resolution: " & resolvedCount & "/" & testCount & " (" & pct & "%)" & vbCrLf & _
                                  IIf(pct >= 80, "Address mode recommended", "Keyword mode recommended")
        If pct >= 80 Then
            lblResultatTest.ForeColor = RGB(0, 128, 0)
        Else
            lblResultatTest.ForeColor = RGB(255, 128, 0)
        End If
    Else
        lblResultatTest.caption = "No email found for the test"
        lblResultatTest.ForeColor = RGB(255, 0, 0)
    End If
End Sub

Private Function GetInboxOrFirst(rootFolder As Outlook.folder) As Outlook.folder
    Dim subFolder As Outlook.folder
    On Error Resume Next
    For Each subFolder In rootFolder.Folders
        If LCase(subFolder.name) Like "*inbox*" Or LCase(subFolder.name) Like "*r�ception*" Or _
           LCase(subFolder.name) Like "*bo�te de r�ception*" Then
            Set GetInboxOrFirst = subFolder
            Exit Function
        End If
    Next subFolder
    If rootFolder.Folders.count > 0 Then
        Set GetInboxOrFirst = rootFolder.Folders(1)
    End If
    On Error GoTo 0
End Function

Private Function TryResolveAddress(Sender As Object) As String
    On Error Resume Next
    TryResolveAddress = ""
    
    If Sender.AddressEntry.Type = "SMTP" Then
        TryResolveAddress = Sender.AddressEntry.Address
        Exit Function
    End If
    
    If Sender.AddressEntry.Type = "EX" Then
        Dim exUser As Object
        Set exUser = Sender.AddressEntry.GetExchangeUser
        If Not exUser Is Nothing Then
            If Len(exUser.PrimarySmtpAddress) > 0 Then
                TryResolveAddress = exUser.PrimarySmtpAddress
                Exit Function
            End If
        End If
    End If
    
    TryResolveAddress = Sender.AddressEntry.Address
    On Error GoTo 0
End Function

' ============================================================================
' BOUTONS S�LECTION
' ============================================================================
Private Sub btnToutSel_Click()
    Dim i As Long
    For i = 0 To lstBoites.ListCount - 1
        lstBoites.Selected(i) = True
    Next i
End Sub

Private Sub btnToutDesel_Click()
    Dim i As Long
    For i = 0 To lstBoites.ListCount - 1
        lstBoites.Selected(i) = False
    Next i
End Sub

' ============================================================================
' �V�NEMENTS P�RIODE
' ============================================================================
Private Sub optTout_Click()
    cboDepuis.Enabled = False
    txtDateDebut.Enabled = False
    txtDateFin.Enabled = False
End Sub

Private Sub optDepuis_Click()
    cboDepuis.Enabled = True
    txtDateDebut.Enabled = False
    txtDateFin.Enabled = False
End Sub

Private Sub optPerso_Click()
    cboDepuis.Enabled = False
    txtDateDebut.Enabled = True
    txtDateFin.Enabled = True
End Sub

' ============================================================================
' BOUTONS PRINCIPAUX
' ============================================================================
Private Sub btnLancer_Click()
    Debug.Print "[frmConfig] btnLancer_Click d�but..."
    
    ' V�rifier qu'au moins une bo�te est s�lectionn�e
    Dim i As Long, hasSelection As Boolean
    hasSelection = False
    For i = 0 To lstBoites.ListCount - 1
        If lstBoites.Selected(i) Then hasSelection = True: Exit For
    Next i
    
    Debug.Print "[frmConfig] hasSelection = " & hasSelection
    
    If Not hasSelection Then
        MsgBox "Please select at least one mailbox.", vbExclamation, "Attention"
        Exit Sub
    End If
    
    ' Sauvegarder la derni�re config
    If Len(m_LastSelectedStore) > 0 Then
        Debug.Print "[frmConfig] Sauvegarde config: " & m_LastSelectedStore
        Call SaveCurrentConfig(m_LastSelectedStore)
    End If
    
    Debug.Print "[frmConfig] m_Cancelled = False"
    m_Cancelled = False
    
    Debug.Print "[frmConfig] Appel Hide..."
    On Error Resume Next
    Me.Hide
    Debug.Print "[frmConfig] Hide termin�, err=" & Err.Number
    On Error GoTo 0
End Sub

Private Sub btnAnnuler_Click()
    m_Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then ' Clic sur X
        Cancel = True  ' Emp�cher la fermeture automatique qui d�truit l'objet
        m_Cancelled = True
        Me.Hide  ' Cacher le formulaire au lieu de le d�truire
    End If
End Sub

' ============================================================================
' PROPRI�T�S PUBLIQUES
' ============================================================================
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

Public Property Get SelectedStores() As Collection
    Dim col As New Collection
    Dim i As Long
    For i = 0 To lstBoites.ListCount - 1
        If lstBoites.Selected(i) Then
            col.Add lstBoites.list(i)
        End If
    Next i
    Set SelectedStores = col
End Property

Public Property Get IsTestMode() As Boolean
    IsTestMode = optTest.Value
End Property

Public Property Get UseExchangeResolution() As Boolean
    UseExchangeResolution = chkResolutionExchange.Value
End Property

Public Property Get AnalyzeLength() As Boolean
    AnalyzeLength = chkAnalyseLongueur.Value
End Property

' V22 : Mode debug (optionnel - si chkDebugMode n'existe pas, retourne False)
Public Property Get DebugMode() As Boolean
    On Error Resume Next
    DebugMode = False
    DebugMode = Me.controls("chkDebugMode").Value
    On Error GoTo 0
End Property

Public Property Get PeriodStart() As Date
    If optTout.Value Then
        PeriodStart = DateSerial(2000, 1, 1)
    ElseIf optDepuis.Value Then
        PeriodStart = DateSerial(CLng(cboDepuis.Value), 1, 1)
    Else
        On Error Resume Next
        PeriodStart = CDate(txtDateDebut.text)
        If Err.Number <> 0 Then PeriodStart = DateSerial(2000, 1, 1)
        On Error GoTo 0
    End If
End Property

Public Property Get PeriodEnd() As Date
    If optTout.Value Or optDepuis.Value Then
        PeriodEnd = Now
    Else
        On Error Resume Next
        PeriodEnd = CDate(txtDateFin.text)
        If Err.Number <> 0 Then PeriodEnd = Now
        On Error GoTo 0
    End If
End Property

Public Function GetStoreColor(storeName As String) As Long
    If m_MailboxConfigs.Exists(storeName) Then
        GetStoreColor = m_MailboxConfigs(storeName)("ColorRGB")
    Else
        GetStoreColor = RGB(128, 128, 128)
    End If
End Function

Public Function GetStoreDomain(storeName As String) As String
    If m_MailboxConfigs.Exists(storeName) Then
        GetStoreDomain = m_MailboxConfigs(storeName)("Domain")
    Else
        GetStoreDomain = ""
    End If
End Function

Public Function GetStoreEmailUser(storeName As String) As String
    If m_MailboxConfigs.Exists(storeName) Then
        GetStoreEmailUser = m_MailboxConfigs(storeName)("EmailUser")
    Else
        GetStoreEmailUser = ""
    End If
End Function

Public Function GetStoreKeyword(storeName As String) As String
    If m_MailboxConfigs.Exists(storeName) Then
        GetStoreKeyword = m_MailboxConfigs(storeName)("Keyword")
    Else
        GetStoreKeyword = ""
    End If
End Function

Public Function GetStoreAddressMode(storeName As String) As Boolean
    If m_MailboxConfigs.Exists(storeName) Then
        GetStoreAddressMode = m_MailboxConfigs(storeName)("AddressMode")
    Else
        GetStoreAddressMode = False
    End If
End Function

Public Function GetStoreDateMin(storeName As String) As Date
    ' Retourne la date minimum pour cette bo�te (ou 1/1/1900 si pas activ�)
    On Error Resume Next
    If m_MailboxConfigs.Exists(storeName) Then
        If m_MailboxConfigs(storeName)("DateMinEnabled") Then
            GetStoreDateMin = m_MailboxConfigs(storeName)("DateMin")
            If GetStoreDateMin < DateSerial(1990, 1, 1) Then
                GetStoreDateMin = DateSerial(1900, 1, 1)
            End If
        Else
            GetStoreDateMin = DateSerial(1900, 1, 1)
        End If
    Else
        GetStoreDateMin = DateSerial(1900, 1, 1)
    End If
    On Error GoTo 0
End Function

' �v�nement pour le checkbox DateMin
Private Sub chkDateMin_Click()
    On Error Resume Next
    Dim isEnabled As Boolean
    isEnabled = Me.controls("chkDateMin").Value
    
    Me.controls("txtDateMin").Enabled = isEnabled
    If isEnabled Then
        Me.controls("txtDateMin").BackColor = &H80000005  ' Blanc
    Else
        Me.controls("txtDateMin").BackColor = &H8000000F  ' Gris
        Me.controls("txtDateMin").text = ""
    End If
    On Error GoTo 0
End Sub


