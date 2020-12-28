VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Calcolo del determinante"
   ClientHeight    =   3915
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancella 
      Caption         =   "cancella"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      ToolTipText     =   "cancella i dati scritti"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtrisultato 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "determinante"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdcalcola 
      Caption         =   "calcola"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "calcola il determinante"
      Top             =   2760
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid matrice 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2143
      _Version        =   327680
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdcrea 
      Caption         =   "crea"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "crea la matrice"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtordine 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "inserire l'ordine della matrice"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblordine 
      Caption         =   "ordine"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnucrea 
         Caption         =   "&Crea"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucalcola 
         Caption         =   "C&alcola"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnucancella 
         Caption         =   "Cance&lla"
         Shortcut        =   ^L
      End
      Begin VB.Menu separa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuesci 
         Caption         =   "&Esci"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&?"
      Begin VB.Menu mnucome 
         Caption         =   "Come usare il programma"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnucalcolo 
         Caption         =   "&Informazioni sul calcolo del determinante"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuprogramma 
         Caption         =   "I&nformazioni sul programma"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim matrix(1 To 100, 1 To 100) As String
Dim ris As Integer
Dim risultato As Integer
Private Sub cmdcalcola_Click()
  On Error GoTo fine
   risultato = 0
   If txtordine.Text = 2 Then
     risultato = (matrix(1, 1) * matrix(2, 2)) - (matrix(1, 2) * matrix(2, 1))
   Else
     If txtordine.Text = 3 Then laplace
   End If
   txtrisultato.Text = "Det=" & risultato
fine:
 End Sub

Private Sub cmdcancella_Click()
  On Error GoTo fine
  For cont = 0 To (txtordine.Text - 1)
    For cont2 = 0 To (txtordine.Text - 1)
      matrix(cont + 1, cont2 + 1) = 0
      matrice.Col = cont
      matrice.Row = cont2
      matrice.Text = ""
    Next
  Next
  txtordine.Text = ""
  txtrisultato.Text = ""
fine:
End Sub

Private Sub cmdcrea_Click()
   On Error GoTo fine
   If (Val(txtordine.Text) > 1) Or (txtordine.Text = " ") Then
      matrice.Visible = True
      matrice.Cols = txtordine.Text
      matrice.Rows = txtordine.Text
   Else
      MsgBox "Immettere numeri maggiori di 1", vbCritical, "Errore"
   End If
   If txtordine.Text > 3 Then
     matrice.Visible = False
     MsgBox "Il programma non calcola il determinante di matrici superiori al terzo ordine."
   End If
fine:
End Sub




Private Sub matrice_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
    If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 45) Then
      MsgBox "inserire solo numeri", vbCritical, "Errore"
    Else
      matrice.Text = matrice.Text & Chr(KeyAscii)
    End If
  Else
    matrice.Text = ""
  End If
  matrix(matrice.Row + 1, matrice.Col + 1) = Val(matrice.Text)
  End Sub

Private Sub laplace()
  On Error GoTo fine
  ris = 0
  If txtordine.Text = 3 Then
    For cont = 1 To txtordine.Text
      indice = 0
      indice2 = 0
      complementoalgebrico (cont)
      indice = indice + 1
      If indice = cont Then
        indice = indice + 1
      End If
      indice2 = indice + 1
      If indice2 = cont Then
        indice2 = indice2 + 1
      End If
      ris = ris * ((matrix((txtordine.Text - 1), indice) * matrix(txtordine.Text, indice2)) - (matrix((txtordine.Text - 1), indice2) * matrix(txtordine.Text, indice)))
      risultato = risultato + ris
    Next
  End If
fine:
End Sub
Private Sub complementoalgebrico(cont)
   On Error GoTo fine
      a = (1 + cont) Mod 2
      If a > 0 Then
        ris = matrix(1, cont) * -1
      Else
        ris = matrix(1, cont) * 1
      End If
fine:
End Sub

Private Sub mnucalcola_Click()
  cmdcalcola_Click
End Sub

Private Sub mnucalcolo_Click()
  Dim info As String
  info = "Il determinante è una funzione che associa un numero reale a tutte e sole le matrici quadrate." & vbCrLf & _
         "Per calcolare il determinante esistono vari metodi:Prodotti incrociati,Regola di Sarrus,Regola di La Place(per i determinanti di ordine superiore al terzo)." & vbCrLf & _
         "Le proprietà del determinante sono:" & vbCrLf & _
         "1) Se gli elementi di una riga o di una colonna sono nulli allora il determinante è nullo." & vbCrLf & _
         "2) Se gli elementi di 2 righe o di 2 colonne sono uguali allora il determinante della matrice è nullo." & vbCrLf & _
         "3) Se gli elementi di una riga o di una colonna si ottengono moltiplicando per gli elementi di un'altra riga o un'altra colonna per uno scalare allora il determinante è nullo." & vbCrLf & _
         "4) Se gli elementi di una riga o di una colonna sono combinazioni lineari degli elementi di altre righe o altre colonne il determinante è nullo."
         MsgBox info, vbInformation, "Informazioni"
End Sub

Private Sub mnucancella_Click()
  cmdcancella_Click
End Sub

Private Sub mnucome_Click()
   Dim info As String
   info = "Calcolo del determinante: " & vbCrLf & _
        "1) Inserire l'ordine della matrice nell'apposita casella di testo." & vbCrLf & _
        "2) Cliccare sul pulsante Crea o dal menu file sulla voce crea." & vbCrLf & _
        "3) Inserire i numeri nelle caselle della tabella che appare." & vbCrLf & _
        "4) Cliccare nel pulsante Calcola o dal menu file nella voce Calcola." & vbCrLf & _
        "5) Per cancellare cliccare nel pulsante Cancella o dal menu file nella voce Cancella." & vbCrLf & _
        "Attenzione:il programma calcola il determinante di matrici di ordine 2 o 3."
   MsgBox info, vbInformation, "Informazioni"
        
End Sub

Private Sub mnucrea_Click()
   cmdcrea_Click
End Sub

Private Sub mnuesci_Click()
  End
End Sub

Private Sub mnuprogramma_Click()
  Dim info As String
  info = "Programma realizzato dall'alunno Davide Basile" & vbCrLf & _
         "Docenti: Brafa Raffaele,Cannizzaro Giovanni " & vbCrLf & _
         "Classe: 5° pr2" & vbCrLf & _
         "Data:04/07/2003" & vbCrLf & _
         "Testo:Procedura  software per il calcolo del determinante" & vbCrLf & _
         "di una matrice. La procedura deve verificare prioritariamente" & vbCrLf & _
         "se la matrice è singolare."

  MsgBox info, vbInformation, "Info"
End Sub
