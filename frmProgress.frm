VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Analyse en cours..."
   ClientHeight    =   7105
   ClientLeft      =   90
   ClientTop       =   410
   ClientWidth     =   10620
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' frmProgress - CODE V23
' Formulaire de progression de l'analyse
' ADAPT� aux noms de contr�les existants
' ============================================================================
Option Explicit

Private m_Cancelled As Boolean
Private m_Paused As Boolean
Private m_ProgressBarWidth As Double  ' Largeur max de la barre

Private Sub UserForm_Initialize()
    Debug.Print "[frmProgress] Initialize..."
    
    On Error Resume Next
    
    m_Cancelled = False
    m_Paused = False
    
    ' Stocker la largeur de la barre de fond pour le calcul du pourcentage
    m_ProgressBarWidth = 0
    If Not fraProgressBack Is Nothing Then
        m_ProgressBarWidth = fraProgressBack.Width
    End If
    If m_ProgressBarWidth = 0 Then m_ProgressBarWidth = 480
    
    ' Initialiser les labels (sans emojis pour compatibilit�)
    lblTitre.caption = "Analyzing emails..."
    lblStatus.caption = "Initializing..."
    lblBoite.caption = "Mailbox : -"
    lblEmails.caption = "Emails processed: 0 / 0"
    lblTempsEcoule.caption = "Elapsed time: 0 min 0 sec"
    lblTempsRestant.caption = "Estimated time left: calculating..."
    
    ' Si lblDetail existe
    On Error Resume Next
    Me.controls("lblDetail").caption = ""
    On Error GoTo 0
    
    ' Si lblPercent existe, l'initialiser
    On Error Resume Next
    Me.controls("lblPercent").caption = "0%"
    On Error GoTo 0
    
    ' Si lblProgressBar existe, l'initialiser
    On Error Resume Next
    Me.controls("lblProgressBar").Width = 0
    Me.controls("lblProgressBar").BackColor = RGB(68, 114, 196)
    On Error GoTo 0
    
    ' Boutons
    On Error Resume Next
    btnPause.caption = "Pause"
    btnPause.Enabled = True
    btnAnnuler.caption = "Cancel"
    btnAnnuler.Enabled = True
    On Error GoTo 0
    
    ' Bouton Fermer (si existe)
    On Error Resume Next
    Me.controls("btnFermer").Visible = False
    On Error GoTo 0
    
    Debug.Print "[frmProgress] Initialize termin�"
End Sub

' ============================================================================
' MISE � JOUR DE LA PROGRESSION
' ============================================================================
Public Sub UpdateProgress(current As Long, total As Long, storeName As String, _
                          storeIndex As Long, storeCount As Long)
    On Error Resume Next
    
    ' Gestion de la pause
    Do While m_Paused And Not m_Cancelled
        DoEvents
    Loop
    
    If m_Cancelled Then Exit Sub
    
    ' Calcul du pourcentage
    Dim pct As Double
    If total > 0 Then
        pct = (current / total) * 100
    Else
        pct = 0
    End If
    
    ' Mise � jour des labels (sans emojis)
    lblBoite.caption = "Mailbox : " & storeName & " (" & storeIndex & "/" & storeCount & ")"
    lblEmails.caption = "Emails processed: " & Format(current, "#,##0") & " / " & Format(total, "#,##0")
    lblStatus.caption = "Analyzing... " & Format(pct, "0") & "%"
    
    ' Pourcentage (si contr�le existe)
    On Error Resume Next
    Me.controls("lblPercent").caption = Format(pct, "0") & "%"
    On Error GoTo 0
    
    ' Barre de progression (si contr�le lblProgressBar existe)
    On Error Resume Next
    Me.controls("lblProgressBar").Width = (pct / 100) * m_ProgressBarWidth
    On Error GoTo 0
    
    DoEvents
End Sub

Public Sub SetTimeInfo(elapsedSeconds As Long, remainingSeconds As Long)
    On Error Resume Next
    
    ' Temps �coul�
    Dim elapsedMin As Long, elapsedSec As Long
    elapsedMin = elapsedSeconds \ 60
    elapsedSec = elapsedSeconds Mod 60
    lblTempsEcoule.caption = "Elapsed time: " & elapsedMin & " min " & elapsedSec & " sec"
    
    ' Temps restant
    If remainingSeconds > 0 Then
        Dim remMin As Long, remSec As Long
        remMin = remainingSeconds \ 60
        remSec = remainingSeconds Mod 60
        lblTempsRestant.caption = "Estimated time left: " & remMin & " min " & remSec & " sec"
    ElseIf remainingSeconds = 0 Then
        lblTempsRestant.caption = "Estimated time left: < 1 min"
    Else
        lblTempsRestant.caption = "Estimated time left: calculating..."
    End If
    
    DoEvents
End Sub

Public Sub SetComplete(message As String)
    On Error Resume Next
    
    lblTitre.caption = "Analysis complete!"
    lblStatus.caption = message
    lblStatus.ForeColor = RGB(0, 128, 0)  ' Vert
    
    ' Barre � 100%
    Me.controls("lblPercent").caption = "100%"
    Me.controls("lblProgressBar").Width = m_ProgressBarWidth
    Me.controls("lblProgressBar").BackColor = RGB(112, 173, 71)  ' Vert
    
    ' Boutons
    btnPause.Enabled = False
    btnAnnuler.Enabled = False
    
    ' Afficher bouton Fermer si existe
    Me.controls("btnFermer").Visible = True
    
    DoEvents
End Sub

Public Sub SetError(message As String)
    On Error Resume Next
    
    lblTitre.caption = "Error"
    lblStatus.caption = "Error: " & message
    lblStatus.ForeColor = RGB(255, 0, 0)  ' Rouge
    
    ' Barre en rouge
    Me.controls("lblProgressBar").BackColor = RGB(255, 0, 0)
    
    ' Boutons
    btnPause.Enabled = False
    btnAnnuler.Enabled = False
    Me.controls("btnFermer").Visible = True
    
    DoEvents
End Sub

' ============================================================================
' PROPRI�T�S
' ============================================================================
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

Public Property Get Paused() As Boolean
    Paused = m_Paused
End Property

' ============================================================================
' BOUTONS
' ============================================================================
Private Sub btnPause_Click()
    On Error Resume Next
    m_Paused = Not m_Paused
    If m_Paused Then
        btnPause.caption = Chr(9654) & " Reprendre"  ' ?
        lblStatus.caption = "Paused..."
    Else
        btnPause.caption = Chr(9208) & " Pause"  ' ?
    End If
End Sub

Private Sub btnAnnuler_Click()
    If MsgBox("Voulez-vous annuler l'analyse ?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        m_Cancelled = True
        m_Paused = False
    End If
End Sub

Private Sub btnFermer_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then  ' Clic sur X
        Cancel = True
        m_Cancelled = True
        m_Paused = False
        Me.Hide
    End If
End Sub


