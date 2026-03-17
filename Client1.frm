VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Client1 
   Caption         =   "Fiche Client"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Client1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Client1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnValider_Click()
Sheets("Client(2)").Activate 'effacer la plage de cellule  ou les informations vont être rentrées
Range("I1:P100").Clear

Dim elmt_select As String
Dim n_commande As Integer
Dim n_client As Integer
Dim statut(100) As String
Dim n As Integer
  n = 2
Dim k As Integer 'K me permet de remplir les cellules en commencant par la cellule b3
k = 2 'c'est pour ca que j'initialise a 2
elmt_select = Client1.ComboBox1.Value 'je recupere la valeur selectionner dans la chaine deroulante

    Range("I1").FormulaR1C1 = "Historique de commande de  " & elmt_select 'Permet de mettre l'entête avec ca mise en forme
    Range("I1:M1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    With Sheets("Client(2)")
    .Range("I2") = "ID_Commandes"
    .Range("J2") = "ID_Client"
    .Range("K2") = "Régularité"
    .Range("L2") = "ID_Payement"
    .Range("M2") = "Statut"
    .Range("N2") = "GrosChantier"
    
    End With
n_commande = Range("commande").Rows.Count
n_client = Range("client").Rows.Count
For i = 1 To n_client
Sheets("Client(2)").Activate
If (Range("C" & i + 1) = elmt_select) Then
    For j = 1 To n_commande
            Sheets("Commande(100)").Activate
            If Range("B" & j + 1) = Sheets("Client(2)").Range("A" & i + 1) Then
                With Sheets("Client(2)")
                .Range("I" & n + 1) = Sheets("Commande(100)").Range("A" & j + 1)
                .Range("J" & n + 1) = Sheets("Commande(100)").Range("B" & j + 1)
                .Range("K" & n + 1) = Sheets("Commande(100)").Range("G" & j + 1)
                .Range("L" & n + 1) = Sheets("Commande(100)").Range("C" & j + 1)
                .Range("M" & n + 1) = Sheets("Commande(100)").Range("D" & j + 1)
                .Range("N" & n + 1) = Sheets("Commande(100)").Range("E" & j + 1)
                n = n + 1
                End With
            End If
     Next j
End If
Next i
Unload Client1 'je charge la userform
Client1.Hide 'je cache la userform
Sheets("Client(2)").Activate 'j'active la feuilles contenant les informations
End Sub


Private Sub ComboBox1_Change()
Dim i As String
Dim j As Integer
    
        j = 2
        i = Sheets("Client(2)").Range("C" & j + 1)
        While i <> ""
        
        
        i = Sheets("Client(2)").Range("C" & j)
        Client1.ComboBox1.AddItem i
        j = j + 1
        Wend
End Sub


Private Sub UserForm_Initialize()
Dim i As String
Dim j As Integer
    
        j = 2
        i = Sheets("Client(2)").Range("C" & j + 1)
        While i <> ""
        
        
        i = Sheets("Client(2)").Range("C" & j)
        Client1.ComboBox1.AddItem i
        j = j + 1
        Wend
End Sub
