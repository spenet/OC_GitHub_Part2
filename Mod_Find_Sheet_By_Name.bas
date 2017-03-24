Attribute VB_Name = "Mod_Find_Sheet_By_Name"
'________________________________________________________________________________
'
' Author : Sylvain P.
' Date   : 12/10/2010
' Rev    : 1.0
' Awacs  : N/A
'________________________________________________________________________________

Option Explicit

' Cette fonction renvoie le numero de feuille d'apres son nom.
' Si la feuille n'est pas trouvée, la fonction renvoie 0 (feuille invalide)

Function f_Find_Sheet_By_Name(s_Sheet_Name As String) As Integer
 Dim index_feuille As Integer
 
  f_Find_Sheet_By_Name = 0
  
  ' Recherche des feuilles pour bosser
  For index_feuille = 1 To Sheets.Count
    If (s_Sheet_Name = Sheets(index_feuille).Name) Then
      f_Find_Sheet_By_Name = index_feuille
    End If
  Next index_feuille
End Function


Sub test_FindSheetByName()
  Dim result As Integer
  result = f_Find_Sheet_By_Name("Feuil2")
End Sub
