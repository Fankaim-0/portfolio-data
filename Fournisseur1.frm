VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fournisseur1 
   Caption         =   "Fiche Fournisseur"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Fournisseur1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Fournisseur1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnValider_Click()
Sheets("Fournisseur(3)").Activate 'effacer la plage de cellule  ou les informations vont être rentrées
Range("E1:M100").Clear

Dim elmt_select As String
Dim n_fournisseur As Integer
Dim n_stockage As Integer
Dim statut(100) As String
Dim n As Integer
  n = 2
Dim k As Integer 'K me permet de remplir les cellules en commencant par la cellule b3
k = 2 'c'est pour ca que j'initialise a 2
elmt_select = Fournisseur1.ComboBox1.Value 'je recupere la valeur selectionner dans la chaine deroulante

    Range("E1").FormulaR1C1 = "Historique d'approvisionnement de  " & elmt_select 'Permet de mettre l'entête avec ca mise en forme
    Range("E1:M1").Select
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
    
    With Sheets("Fournisseur(3)")
    .Range("E2") = "ID_Fournisseur"
    .Range("F2") = "NomFournisseur"
    .Range("G2") = "Quantité"
    .Range("H2") = "Seuil"
    .Range("I2") = "DateLivraisonProduit"
    .Range("J2") = "QuantitéLivraison"
    
    End With
n_fournisseur = Range("fournisseur").Rows.Count
n_stockage = Range("stockage").Rows.Count
For i = 1 To n_fournisseur
Sheets("Fournisseur(3)").Activate
If (Range("B" & i + 1) = elmt_select) Then
    For j = 1 To n_stockage
            Sheets("Stockage(6)").Activate
            If Range("C" & j + 1) = Sheets("Fournisseur(3)").Range("A" & i + 1) Then
                With Sheets("Fournisseur(3)")
                .Range("E" & n + 1) = Sheets("Fournisseur(3)").Range("A" & j + 1)
                .Range("F" & n + 1) = Sheets("Fournisseur(3)").Range("B" & j + 1)
                .Range("G" & n + 1) = Sheets("Stockage(6)").Range("D" & j + 1)
                .Range("H" & n + 1) = Sheets("Stockage(6)").Range("E" & j + 1)
                .Range("I" & n + 1) = Sheets("Stockage(6)").Range("F" & j + 1)
                .Range("I" & n + 1).NumberFormat = "m/d/yyyy"
                .Range("J" & n + 1) = Sheets("Stockage(6)").Range("G" & j + 1)
                n = n + 1
                End With
            End If
     Next j
End If
Next i
Unload Fournisseur1 'je charge la userform
Fournisseur1.Hide 'je cache la userform
Sheets("Fournisseur(3)").Activate 'j'active la feuilles contenant les informations
End Sub

Private Sub ComboBox1_Change()
Dim i As String
Dim j As Integer
    
        j = 2
        i = Sheets("Fournisseur(3)").Range("B" & j + 1)
        While i <> ""
        
        
        i = Sheets("Fournisseur(3)").Range("B" & j)
        Fournisseur1.ComboBox1.AddItem i
        j = j + 1
        Wend
End Sub


Private Sub UserForm_Initialize()
Dim i As String
Dim j As Integer
    
        j = 2
        i = Sheets("Fournisseur(3)").Range("B" & j + 1)
        While i <> ""
        
        
        i = Sheets("Fournisseur(3)").Range("B" & j)
        Fournisseur1.ComboBox1.AddItem i
        j = j + 1
        Wend
End Sub
