VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Statut_Com 
   Caption         =   "Histogramme des commandes"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Statut_Com.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Statut_Com"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnValider_Click()
Sheets("GESTION").Activate 'effacer la plage de cellule  ou les informations vont être rentrées
Range("A1:E100").Clear

Dim elmt_select As String
Dim n_commande As Integer
Dim statut(100) As String
Dim k As Integer 'K me permet de remplir les cellules en commencant pr la cellule b3
k = 2 'c'est pour ca que j'initialise a 2
elmt_select = Statut_Com.ComboBox1.Value 'je recupere la valeur selectionner dans la chaine deroulante

    Range("A1").FormulaR1C1 = "Commandes(100)" & elmt_select 'Permet de mettre l'entête avec ca mise en forme
    Range("A1:E1").Select
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
n_commande = Range("commande").Rows.Count
For i = 1 To n_commande
Sheets("Commande(100)").Activate

If Range("D" & i + 1) = elmt_select Then 'Je recherche les clients correspondant a la valeur dans la chaine deroulante

With Sheets("GESTION")
.Range("A2") = "ID_Commande"
.Range("B2") = "ID_Clients"
.Range("C2") = "ID_Artisan"
.Range("D2") = "GrosChantier"
.Range("E2") = "Prix_Commandes"
.Range("A" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 1, False) 'recherche ID COMMANDES
.Range("B" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("B" & i + 1), Range("TabClient"), 1, False) 'recherche ID CLIENT
.Range("D" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 5, False) 'recherche type de travaux
.Range("E" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 12, False) 'Prix
k = k + 1
End With
End If
If elmt_select = "Devis" Then 'Lorsqu'il selectionne tout les clients


With Sheets("GESTION")
.Range("A2") = "ID_COMMANDES"
.Range("B2") = "ID_Clients"
.Range("C2") = "ID_Artisan"
.Range("D2") = "GrosChantier"
.Range("E2") = "Prix_Commandes"
.Range("A" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 1, False) 'recherche ID COMMANDES
.Range("B" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("B" & i + 1), Range("TabClient"), 1, False) 'recherche nom client
.Range("D" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 5, False) 'recherche type de travaux
.Range("E" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 12, False)  'Prix
k = k + 1
End With
End If
If elmt_select = "En cours" Then 'Lorsqu'il selectionne tout les clients


With Sheets("GESTION")
.Range("A2") = "ID_COMMANDES"
.Range("B2") = "ID_Clients"
.Range("C2") = "ID_Artisan"
.Range("D2") = "GrosChantier"
.Range("E2") = "Prix_Commandes"
.Range("A" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 1, False) 'recherche ID COMMANDES
.Range("B" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("B" & i + 1), Range("TabClient"), 1, False) 'recherche nom client
.Range("D" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 5, False) 'recherche type de travaux
.Range("E" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 12, False)  'Prix
k = k + 1
End With
End If
If elmt_select = "Terminé" Then 'Lorsqu'il selectionne tout les clients


With Sheets("GESTION")
.Range("A2") = "ID_COMMANDES"
.Range("B2") = "ID_Clients"
.Range("C2") = "ID_Artisan"
.Range("D2") = "GrosChantier"
.Range("E2") = "Prix_Commandes"
.Range("A" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 1, False) 'recherche ID COMMANDES
.Range("B" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("B" & i + 1), Range("TabClient"), 1, False) 'recherche nom client
.Range("D" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 5, False) 'recherche type de travaux
.Range("E" & k + 1) = Application.WorksheetFunction.VLookup(Sheets("Commande(100)").Range("A" & i + 1), Range("TabCommande"), 12, False)  'Prix
k = k + 1
End With
End If
Next i

i = 2
With Sheets("GESTION")
If CheckBox1.Value = True Then
flag = True
    While flag = True

       'If i = 1 Then
        'If Sheets("Commande(100)").Range("D" & i + 1) = Application.WorksheetFunction.VLookup(Sheets("GESTION").Range("A" & i + 1), Range("TabCommande"), 11, False) Then
        If (Application.WorksheetFunction.VLookup(.Cells(i + 1, 1), Range("TabCommande"), 11, False) = "1" And Application.WorksheetFunction.VLookup(.Cells(i + 1, 1), Range("TabCommande"), 9, False) >= 2000) Or (Application.WorksheetFunction.VLookup(.Cells(i + 1, 1), Range("TabCommande"), 11, False) = "0" And Application.WorksheetFunction.VLookup(.Cells(i + 1, 1), Range("TabCommande"), 9, False) >= 4000) Then
        k = 1
        While (k < 6)
        .Cells(i + 1, k).Interior.Color = RGB(174, 240, 194)
        k = k + 1
        Wend
            'MsgBox "ok"
        End If
         If .Cells(i + 2, 1) = "" Then
         flag = False
        End If
        i = i + 1
       
        
    Wend
End If
End With

Unload Statut_Com 'je charge la userform
Statut_Com.Hide 'je cache la userform
Sheets("GESTION").Activate 'j'active la feuilles contenant les informations
End Sub


Private Sub CheckBox1_Click()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub UserForm_Initialize()
With Statut_Com.ComboBox1
.AddItem "DEVIS"
.AddItem "EN COURS"
.AddItem "TERMINE"
End With

End Sub
