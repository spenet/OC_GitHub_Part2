Attribute VB_Name = "Mod_Template_Func"
'________________________________________________________________________________
'
' Author      : Sylvain P.
' Date        : dd/mm/yyyy
' Rev         : 1.0
' Description :
'
'
'________________________________________________________________________________
Option Explicit

'________________________________________________________________________________
'
' Sub_Name    : Template_fonction
' Date        : dd/mm/yyyy
' Description : C'est quoi qu'elle fait la fonction ?
' Input       : liste des paramètres d'entrée, type, range et rôle
'
' Output      : paramètre de sortie, type, range et rôle (en cas d'erreur qu'est ce qui est renvoyé ?)
'
'________________________________________________________________________________
Function Template_function() As Variant ' Modifier le type de sortie suivant le besoin
  On Error GoTo ManageError


  Template_function = 0 ' Penser à renvoyer la valeur de retour avant de quitter
  Exit Function
ManageError:
  MsgBox ("Erreur dans la fonction 'Template_function' : " & Err.Description) ' Penser à changer le nom de la fonction
  Resume Next
End Function

