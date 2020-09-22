VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Referenziare "Microsoft Word 8.0 Object Library" o superiore
'Reference "Microsoft Word 8.0 Object Library" or more
Dim xWord As Word.Application ' L'applicazione Word         (The Word Application
Dim xRange As Range           ' Oggetto Range               (Object Range)
Dim xSelection As Find        ' Oggetto Find                (Object Find)
Dim xTabella As Table         ' Oggetto Tabella             (Object Table)
Dim xCella As Cell            ' Oggetto Cella               (Object Cell)
    Set xWord = New Application
    xWord.Visible = False
'Aggiungo un documento o un modello fatto precedentemente che si chiama prova.dot
'Add a document or a model do precedent call "prova.dot"
    xWord.Documents.Add App.Path & "\prova.dot"
'Protetto da una password supponiamo sia "pass"
'Protect by a password like "pass"
    xWord.ActiveDocument.Unprotect "pass"
'Aggiungo i valori dal record (io ho usato dati fissi per semplicità) al posto di
'%%nome%%, %%cognome%%, %%data%% che devono comparire nel documento di word
'come "%%<nome campo>%%"
'
'Add Record Value
    Set xRange = xWord.ActiveDocument.Range
        xRange.Find.Execute "%%nome%%", , , , , , , , , "Paperino", True: '"Paperino" Can Be substitute by a TextBox
    Set xRange = xWord.ActiveDocument.Range
        xRange.Find.Execute "%%cognome%%", , , , , , , , , "Pippo", True '"Pippo" Can Be substitute by a TextBox
    Set xRange = xWord.ActiveDocument.Range
        xRange.Find.Execute "%%data%%", , , , , , , , , "01/01/2000", True '"01/01/2000" Can Be substitute by a TextBox
'Ripristina la password                         (Put again the password)
    xWord.ActiveDocument.Protect wdAllowOnlyFormFields, , "pass"
'Per visualizze il documento                    (Show The Document)
    xWord.Visible = True
    xWord.WindowState = wdWindowStateMaximize
    xWord.Application.Activate
'Per visualizzare l'anteprima di stampa         (Print Preview)
    'xWord.ActiveDocument.PrintPreview
'Per inviare via email il documento             (Send By email)
    'xWord.ActiveDocument.SendMail
'Per inviare via fax                            (Send A fax)
    'xWord.ActiveDocument.SendFax
'Per stampare il documento                      (Print)
    'xWord.ActiveDocument.PrintOut
'Per salvarlo in una directory                  (Save In A directory)
    'xWord.ActiveDocument.SaveAs App.Path & "\Docs\" & "MyDoc.doc"
End Sub
