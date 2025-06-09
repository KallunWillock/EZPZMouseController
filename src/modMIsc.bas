Attribute VB_Name = "modMIsc"
Public Sub EZPZMouseWheelDemo()
  
  frmMouse_Wheel.Show vbModeless

End Sub

Public Sub EZPZMouseWindowlessControlsDemo()
  
  frmMouse_Windowless.Show vbModeless

End Sub

Public Function AddMSInkReference() As Long
  
  ' Name:       Microsoft Tablet PC Type Library 1.0
  ' Location:   C:\Program Files\Common Files\Microsoft Shared\ink\InkObj.dll
  '             C:\Program Files (x86)\Common Files\Microsoft Shared\ink\InkObj.dll
  
  On Error GoTo ErrHandler
  
  Dim hadAccess As Boolean: hadAccess = VBTrustedAccess
  If Not hadAccess Then
    Dim result As TDBUTTONS_RETURN_CODES
    result = TaskBox("Please enable access", "Please enable 'Trust access to the VBA project object model' " & _
                        "so that a reference to 'Microsoft Tablet PC Type Library 1.0' can be added" & _
                        vbNewLine & vbNewLine & "Would you like to visit the settings window?", _
                        "Settings", TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON, TD_SHIELD_GRADIENT_ICON)
    If result = IDYES Then
      ShowMacroSecurityDialog
    ElseIf result = IDNO Then
      Exit Function
    End If
  End If
  
  Const MSINKAUTLibGUID As String = "{7D868ACD-1A5D-4A47-A247-F39741353012}"
  ThisWorkbook.VBProject.References.AddFromGuid MSINKAUTLibGUID, 1, 0
  
ErrHandler:
  ' Error# &H802D indicates that the library reference is already set
  AddMSInkReference = Err.Number

  If Not hadAccess And VBTrustedAccess Then
    TaskBox "Thank you!", "You can now revoke 'Trust access to the VBA project object model'.", "Settings", TDCBF_OK_BUTTON, TD_SHIELD_OK_ICON
    ShowMacroSecurityDialog
  End If
  
End Function
  
Public Sub Setup()
  
  Dim result As Long
  result = AddMSInkReference

  Select Case result
    Case 0:         TaskBox "Success!", "You have added a reference to Microsoft Tablet PC Type Library 1.0", "Reference added", TDCBF_OK_BUTTON, TD_SHIELD_OK_ICON
    Case &H802D&:   TaskBox "Success!", "You already have a reference set to Microsoft Tablet PC Type Library 1.0. You're good to go!", "Reference already added", TDCBF_OK_BUTTON, TD_SHIELD_OK_ICON
    Case Else
      Dim Response As TDBUTTONS_RETURN_CODES
      Response = TaskBox("ERROR", "Error " & Err.Number & " - " & Err.Description & vbNewLine & vbNewLine & "Unable to add reference to Microsoft Tablet PC Type Library 1.0", "Failed to add reference", TDCBF_OK_BUTTON Or TDCBF_RETRY_BUTTON, TD_SHIELD_ERROR_ICON)
      If Response = IDRETRY Then Setup
  End Select

End Sub

' Many thanks to Cristian Buse for his suggested amends and his code below/above

Public Function VBTrustedAccess() As Boolean
  VBTrustedAccess = IsAccessToVBProjectsOn(ThisWorkbook)
End Function

Public Sub ShowMacroSecurityDialog()
  Application.CommandBars.ExecuteMso "MacroSecurity"
End Sub

Private Function IsAccessToVBProjectsOn(ByVal Target As Workbook) As Boolean
  Dim DummyProject As Object
  On Error Resume Next
  Set DummyProject = Target.VBProject.VBComponents
  IsAccessToVBProjectsOn = (Err.Number = 0)
  On Error GoTo 0
End Function
