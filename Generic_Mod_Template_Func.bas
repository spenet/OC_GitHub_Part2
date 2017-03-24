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
' Input       : liste des param�tres d'entr�e, type, range et r�le
'
' Output      : param�tre de sortie, type, range et r�le (en cas d'erreur qu'est ce qui est renvoy� ?)
'
'________________________________________________________________________________
Function Template_function() As Variant ' Modifier le type de sortie suivant le besoin
  On Error GoTo ManageError


  Template_function = 0 ' Penser � renvoyer la valeur de retour avant de quitter
  Exit Function
ManageError:
  MsgBox ("Erreur dans la fonction 'Template_function' : " & Err.Description) ' Penser � changer le nom de la fonction
  Resume Next
End Function

