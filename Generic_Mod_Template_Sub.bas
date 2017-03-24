Attribute VB_Name = "Mod_Template_Sub_Func"
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
' Sub_Name    : Template_sub
' Date        : dd/mm/yyyy
' Description : C'est quoi qu'elle fait la procédure ?
' Input       : liste des paramètres d'entrée, type, range et rôle
'
'________________________________________________________________________________
Sub Template_sub()
  On Error GoTo ManageError


  Exit Sub
ManageError:
  MsgBox ("Erreur dans la fonction 'Template_sub' : " & Err.Description) ' Penser à changer le nom de la fonction
  Resume Next
End Sub