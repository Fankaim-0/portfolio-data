VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Stockage1 
   Caption         =   "Bilan de stocks"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Stockage1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Stockage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnValider_Click()
Sheets("Stockage(6)").Activate 'effacer la plage de cellule  ou les informations vont être rentrées
Range("I1:P100").Clear

Dim elmt_select As String
Dim n_stockage As Integer
Dim statut(100) As String
Dim n As Integer
  n = 2
Dim k As Integer 'K me permet de remplir les cellules en commencant par la cellule b3
k = 2 'c'est pour ca que j'initialise a 2
elmt_select = Stockage1.ComboBox1.Value 'je recupere la valeur selectionner dans la chaine deroulante

    Range("I1").FormulaR1C1 = "Bilan de stocks de  " & elmt_select 'Permet de mettre l'entête avec ca mise en forme
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
    With Sheets("Stockage(6)")
    .Range("I2") = "ID_Stock"
    .Range("J2") = "Quantité"
    .Range("K2") = "Seuil"
    .Range("L2") = "DateLivraisonProduit"
    .Range("M2") = "QuantitéLivraison"
    
    End With
n_stockage = Range("stockage").Rows.Count
For i = 1 To n_stockage
Sheets("Stockage(6)").Activate
If (Range("D" & i + 1) = elmt_select) Then
    For j = 1 To n_stockage
            Sheets("Stockage(6)").Activate
            If Range("D" & j + 1) < 5 Then
                With Sheets("Stockage(6)")
                .Range("I" & n + 1) = Sheets("Stockage(6)").Range("A" & j + 1)
                .Range("J" & n + 1) = Sheets("Stockage(6)").Range("D" & j + 1)
                .Range("K" & n + 1) = Sheets("Stockage(6)").Range("E" & j + 1)
                .Range("L" & n + 1) = Sheets("Stockage(6)").Range("F" & j + 1)
                .Range("L" & n + 1).NumberFormat = "m/d/yyyy"
                .Range("M" & n + 1) = Sheets("Stockage(6)").Range("G" & j + 1)
               
                n = n + 1
                End With
                Else
                  With Sheets("Stockage(6)")
                .Range("I" & n + 1) = Sheets("Stockage(6)").Range("A" & j + 1)
                .Range("J" & n + 1) = Sheets("Stockage(6)").Range("D" & j + 1)
                .Range("K" & n + 1) = Sheets("Stockage(6)").Range("E" & j + 1)
                .Range("L" & n + 1) = Sheets("Stockage(6)").Range("F" & j + 1)
                .Range("L" & n + 1).NumberFormat = "m/d/yyyy"
                .Range("M" & n + 1) = Sheets("Stockage(6)").Range("G" & j + 1)
               
                n = n + 1
                End With
                
            End If
     Next j
End If
Next i
Unload Stockage1 'je charge la userform
Stockage1.Hide 'je cache la userform
Sheets("Stockage(6)").Activate 'j'active la feuilles contenant les informations
End Sub


Private Sub ComboBox1_Change()
Dim i As String
Dim j As Integer
    
        j = 2
        i = Sheets("Stockage(6)").Range("D" & j + 1)
        While i <> ""
        
        
        i = Sheets("Stockage(6)").Range("D" & j)
        Client1.ComboBox1.AddItem i
        j = j + 1
        Wend
End Sub

