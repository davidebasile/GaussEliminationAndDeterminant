VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Metodo di Gauss per la risoluzione dei sistemi lineari"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancella 
      Caption         =   "cancella"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      ToolTipText     =   "cancella il sistema"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdrisolvi 
      Caption         =   "risolvi"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      ToolTipText     =   "risolve "
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtincognite 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "numero incognite"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtequazioni 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "numero equazioni"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdcrea 
      Caption         =   "crea"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "crea il sistema"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblsoluzioni 
      Height          =   3495
      Left            =   6720
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblsistema 
      Height          =   4695
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label lblincognite 
      Caption         =   "n. incognite"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblequazioni 
      Caption         =   "n. equazioni"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnucrea 
         Caption         =   "&Crea"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucancella 
         Caption         =   "C&ancella"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnurisolvi 
         Caption         =   "&Risolvi"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnustampa 
         Caption         =   "&Stampa"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusepara1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuesci 
         Caption         =   "&Esci"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuopzioni 
      Caption         =   "&Opzioni"
      Begin VB.Menu mnuvisualizza 
         Caption         =   "&Visualizza passaggi"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnunonvisualizza 
         Caption         =   "&Non visualizzare passaggi"
         Checked         =   -1  'True
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuguida 
      Caption         =   "&?"
      Begin VB.Menu mnucome 
         Caption         =   "Come usare il programma"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnugauss 
         Caption         =   "&Informazioni sul metodo di Giordan Gauss"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuprogramma 
         Caption         =   "In&formazioni sul programma"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim impossibile As Boolean
Dim scrivi As Boolean
Dim stampa As String
Dim appoggio(0 To 100, 0 To 100) As String
Dim sistema(0 To 100, 0 To 100) As String
Dim scrivesistema As String
Dim soluzioni As String

Private Sub cmdcancella_Click()
   lblsistema.Caption = ""
   lblsoluzioni.Caption = ""
   txtincognite.Text = ""
   txtequazioni.Text = ""
   For cont = 0 To 100
     For cont1 = 0 To 100
       sistema(cont, cont1) = ""
       appoggio(cont, cont1) = ""
     Next
   Next
   scrivesistema = ""
   stampa = ""
   cmdrisolvi.Enabled = False
   mnurisolvi.Enabled = False
End Sub

Private Sub cmdcrea_Click()
   Dim stringa As String
   On Error GoTo fine
   stampa = ""
   If (txtincognite.Text = "") Or (txtequazioni.Text = "") Then
     MsgBox "Inserire prima il numero di equazioni e di incognite"
     GoTo fine
   End If
   scrivesistema = ""
   If txtequazioni >= txtincognite Then
     lblsistema.Caption = ""
     For cont = 1 To txtequazioni
       For cont2 = 1 To txtincognite
           stringa = "EQUAZIONE " & cont & ":" & vbCrLf & "INSERIRE COEFFICIENTE DELLA " & cont2 & "° INCOGNITA" _
                     & vbCrLf & "Attenzione: inserire anche il segno del coefficiente"
           sistema(cont, cont2) = InputBox(stringa)
           If sistema(cont, cont2) = "" Then
             GoTo fine
           End If
           scrivesistema = scrivesistema & sistema(cont, cont2) & "X" & cont2 & " "
           lblsistema.Caption = scrivesistema
       Next
       stringa = "INSERIRE TERMINE NOTO DELL'EQUAZIONE" & cont & vbCrLf _
                 & "Attenzione:il termine note verrà inserito a secondo membro"
       sistema(cont, txtincognite + 1) = InputBox(stringa)
       scrivesistema = scrivesistema & "=" & sistema(cont, txtincognite + 1) & vbCrLf
       lblsistema.Caption = scrivesistema
     Next
     
     For cont = 1 To txtequazioni.Text
       For cont2 = 1 To (txtincognite.Text + 1)
         appoggio(cont, cont2) = sistema(cont, cont2)
       Next
     Next
     cmdrisolvi.Enabled = True
     mnurisolvi.Enabled = True
   Else
      MsgBox "il sistema non è risolubile"
   End If
fine:
End Sub

Private Sub cmdrisolvi_Click()
   On Error GoTo fine
   impossibile = False
   scrivi = False
   scritto = False
   If mnuvisualizza.Checked = True Then
     scrivi = True
     scritto = True
   End If
rifai:
   stampa = ""
   stampa = stampa & lblsistema.Caption & vbCrLf
   For cont = 1 To txtincognite - 1
     If sistema(cont, cont) <> 0 Then
       For cont2 = 1 To (txtequazioni - cont)
         moltiplicatore = -1 * (sistema(cont2 + cont, cont) / sistema(cont, cont))
         For cont3 = 1 To (txtincognite)
           sistema(cont2 + cont, cont3) = sistema(cont2 + cont, cont3) + (moltiplicatore * sistema(cont, cont3))
         Next
         sistema(cont2 + cont, txtincognite + 1) = sistema(cont2 + cont, txtincognite + 1) + (moltiplicatore * sistema(cont, txtincognite + 1))
         moltiplicatore = Format(moltiplicatore, "fixed")
         moltip = "R" & cont2 + cont & " = R" & cont2 + cont & " +(" & moltiplicatore & "R" & cont & ")"
         If scrivi = True Then
           MsgBox moltip
           stampa = stampa & vbCrLf & moltip & vbCrLf & vbCrLf
         End If
        scrivesistem
       Next
     Else
       MsgBox "Il sistema non è risolubile"
       lblsoluzioni.Caption = "Non risolubile"
       impossibile = True
       Exit For
     End If
   Next
   If impossibile <> True Then
     sostituisci
   End If
   If (scritto = False) And (mnunonvisualizza.Checked = False) Then
     If scrivi = True Then
       scritto = True
       GoTo rifai
     End If
   Else
     If impossibile = False Then lblsoluzioni.Caption = soluzioni
   End If
fine:
End Sub

Private Sub sostituisci()
   On Error GoTo fine
   impossibile = False
   soluzioni = "soluzioni :" & vbCrLf
   For cont = 0 To (txtincognite - 1)
     If sistema(txtincognite - cont, txtincognite - cont) <> 0 Then
       sistema(txtincognite - cont, txtincognite + 1) = sistema(txtincognite - cont, txtincognite + 1) / sistema(txtincognite - cont, txtincognite - cont)
       sistema(txtincognite - cont, txtincognite - cont) = "1"
       For cont1 = 1 To (txtincognite - (cont + 1))
          sistema(txtincognite - cont - cont1, txtincognite + 1) = sistema(txtincognite - cont - cont1, txtincognite + 1) - (sistema(txtincognite - cont - cont1, txtincognite - cont) * sistema(txtincognite - cont, txtincognite + 1))
          sistema(txtincognite - cont - cont1, txtincognite - cont) = 0
       Next
       sistema(txtincognite - cont, txtincognite + 1) = Format(sistema(txtincognite - cont, txtincognite + 1), "fixed")
       soluzioni = soluzioni & "X" & (txtincognite - cont) & "=" & sistema(txtincognite - cont, txtincognite + 1) & vbCrLf
     Else
       MsgBox "IL SISTEMA NON E' RISOLUBILE"
       impossibile = True
       lblsoluzioni.Caption = "Non risolubile"
       Exit For
     End If
     scrivesistem
     If scrivi = True Then MsgBox "sostituzione..."
   Next
   If txtequazioni.Text > txtincognite.Text Then
     For cont = 1 To (txtequazioni.Text - txtincognite.Text)
       If (sistema(txtincognite + cont, txtincognite + 1) <> 0) And (sistema(txtincognite + cont, txtincognite) <> 0) Then
         sistema(txtincognite + cont, txtincognite + 1) = sistema(txtincognite + cont, txtincognite + 1) / (sistema(txtincognite + cont, txtincognite) * sistema(txtincognite, txtincognite + 1))
       End If
       sistema(txtincognite + cont, txtincognite) = 0
       scrivesistem
       If sistema(txtincognite + cont, txtincognite + 1) <> 0 Then
         impossibile = True
         MsgBox " IL SISTEMA NON E' RISOLUBILE"
         lblsoluzioni.Caption = "Non risolubile"
         Exit For
       End If
     Next
   End If
   If impossibile = False Then
       If scrivi = True Then
         lblsoluzioni.Caption = soluzioni
       Else
         scrivi = True
         For cont = 1 To txtequazioni.Text
           For cont2 = 1 To (txtincognite.Text + 1)
             sistema(cont, cont2) = appoggio(cont, cont2)
           Next
         Next
       End If
  End If
fine:
End Sub

Private Sub scrivesistem()
  If scrivi = True Then
    scrivesistema = ""
    For cont = 1 To txtequazioni
      For cont2 = 1 To txtincognite
        If sistema(cont, cont2) <> "0" Then
          sistema(cont, cont2) = Format(sistema(cont, cont2), "fixed")
          If sistema(cont, cont2) > 0 Then
             sistema(cont, cont2) = "+" & sistema(cont, cont2)
           End If
          scrivesistema = scrivesistema & sistema(cont, cont2) & "X" & cont2 & " "
        Else
          scrivesistema = scrivesistema & "            "
        End If
      Next
      sistema(cont, txtincognite + 1) = Format(sistema(cont, txtincognite + 1), "fixed")
      If sistema(cont, txtincognite + 1) > 0 Then
        sistema(cont, txtincognite + 1) = "+" & sistema(cont, txtincognite + 1)
      End If
      scrivesistema = scrivesistema & " = " & sistema(cont, txtincognite + 1) & vbCrLf
    Next
    lblsistema.Caption = scrivesistema
    stampa = stampa & lblsistema.Caption & vbCrLf
  End If
End Sub
         

Private Sub mnucancella_Click()
   cmdcancella_Click
End Sub

Private Sub mnucome_Click()
   Dim info As String
   info = "Come usare il programma:" & vbCrLf & "1) Inserire il numero di equazioni e di incognite nelle apposite caselle." & vbCrLf & _
           "2) Cliccare nel pulsante crea oppure dal menu file cliccare nella voce Crea." & vbCrLf & _
           "3) Immettere i coefficienti comprensivi di segno(uno alla volta),il sistema deve essere semplificato e con termine noto a secondo membro." & vbCrLf & _
           "4) Dopo aver scritto il sistema,cliccare sul pulsante Risolvi o sulla voce risolvi dal menu file. " & vbCrLf & _
           "5) Se dal menu Operazioni è stata selezionata la voce Visualizza Passaggi il programma visualizzerà i passaggi della risoluzione" & vbCrLf & _
           "6) Se si desidera continuare, cliccare sul pulsante Cancella o sulla voce cancella dal menu file per cancellare il precedente sistema, e riprendere dal passo 1 ." & vbCrLf & _
           "7) Se si desidera stampare l'esercizio, dal menu file cliccare sulla voce stampa. " & vbCrLf & _
           "8) Per uscire dal programma, cliccare sulla voce esci dal menu file." & vbCrLf & _
           "Avvertenza: il programma riconosce non risolubile un sistema impossibile o indeterminato."
   MsgBox info, vbInformation, "Guida"
End Sub

Private Sub mnucrea_Click()
   cmdcrea_Click
End Sub

Private Sub mnuesci_Click()
   End
End Sub

Private Sub mnugauss_Click()
  Dim info As String
  Dim info2 As String
  Dim info3 As String
  info = "Informazioni sul metodo di Giordan Gauss: " & vbCrLf & _
        "Il metodo di Gauss serve a risolvere i sistemi lineari (sistemi costituiti da un numero qualsiasi di equazioni e da un numero qualsiasi di incognite di primo grado) attraverso " & _
        "l'annullamento progressivo dei coefficienti che si ottiene sommando all'equazione i-ma stessa un'altra equazione moltiplicata per una costante fino a giungere alla soluzione. " & _
        " Di seguito le caratteristiche delle soluzioni dei sistemi lineari in corrispondenza delle relazioni intercorrenti fra il rango r(A) della matrice incompleta,il rango r(A/B) della matrice " & _
        " completa, il numero m di equazioni e il numero n delle incognite"
  
  info2 = "                          SISTEMA OMOGENEO                             " & vbCrLf & _
          "1   r(A)= m =n       Soluzione banale (0,0,..0)                         " & vbCrLf & _
          "2   r(A) = m < n     infinito n-m autosoluzioni + quella banale         " & vbCrLf & _
          "3   r (A) < m < n    infinito n-r autosoluzioni + quella banale        " & vbCrLf & _
          "4   r (A) < m = n    infinito n-r autosoluzioni + quella banale         " & vbCrLf & _
          "5   r(A) = n < m     Soluzione banale (0,0,..0)                         " & vbCrLf & _
          "6   r (A) < n < m    infinito n-r autosoluzioni + quella banale         "
   info3 = "           SISTEMA NON OMOGENEO    " & vbCrLf & _
          "1   r(A)= m =n       Una sola soluzione " & vbCrLf & _
          "2   r(A) = m < n     infinito n-m soluzioni" & vbCrLf & _
          "3   r (A) < m < n    infinito n-r soluzioni se  r(A/B)=r(A) " & vbCrLf & _
          "4   r (A) < m = n    infinito n-r soluzioni se  r(A/B)=r(A) " & vbCrLf & _
          "5   r(A) = n < m     Una sola soluzione se  r(A/B)=r(A) " & vbCrLf & _
          "6   r (A) < n < m    infinito n-r soluzioni se  r(A/B)=r(A) "
  MsgBox info, vbInformation, "Informazioni"
  MsgBox info2, vbInformation, "Informazioni"
  MsgBox info3, vbInformation, "Informazioni"
End Sub

Private Sub mnunonvisualizza_Click()
  mnuvisualizza.Checked = False
  mnunonvisualizza.Checked = True
End Sub

Private Sub mnuprogramma_Click()
  Dim info As String
  info = "Programma realizzato dall'alunno Davide Basile" & vbCrLf & _
         "Docenti: Brafa Raffaele,Cannizzaro Giovanni " & vbCrLf & _
         "Classe: 5° pr2" & vbCrLf & _
         "Data:04/07/2003" & vbCrLf & _
         "Testo:Procedura software per la risoluzione mediante il metodo di Gauss Giordan" & vbCrLf & _
         "di un sistema lineare nxm .La procedura deve verificare prioritariamente se" & vbCrLf & _
         "il sistema è risolubile."


  MsgBox info, vbInformation, "Info"
End Sub

Private Sub mnurisolvi_Click()
    cmdrisolvi_Click
End Sub

Private Sub mnustampa_Click()
   Printer.Print "   "; stampa; "  "; lblsoluzioni.Caption
   Printer.EndDoc
End Sub
Private Sub mnuvisualizza_Click()
   mnuvisualizza.Checked = True
   mnunonvisualizza.Checked = False
End Sub
