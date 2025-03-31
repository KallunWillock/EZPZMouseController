Attribute VB_Name = "modMIsc"
Public Sub EZPZMouseWheelDemo()
  
  frmMouse_Wheel.Show vbModeless

End Sub

Public Sub EZPZMouseWindowlessControlsDemo()
  
  frmMouse_Windowless.Show vbModeless

End Sub

Public Function AddMSInkReference() As Long
  
  ' Name:       Microsoft Tablet PC Type Library 1.0
  ' Location:   C:\Program Files (x86)\Common Files\Microsoft Shared\ink\InkObj.dll
  
  On Error GoTo ErrHandler
  
  Const MSINKAUTLibGUID As String = "{7D868ACD-1A5D-4A47-A247-F39741353012}"
  ThisWorkbook.VBProject.References.AddFromGuid MSINKAUTLibGUID, 1, 0
  
ErrHandler:
  ' Error# &H802D indicates that the library reference is already set
  AddMSInkReference = Err.Number

End Function

Public Sub Setup()
  
  Dim Result As Long
  Result = AddMSInkReference

  Select Case Result
    Case 0:         TaskBox "Success!", "You have added a reference to Microsoft Tablet PC Type Library 1.0", "Reference added", TDCBF_OK_BUTTON, TD_SHIELD_OK_ICON
    Case &H802D&:   TaskBox "Success!", "You already have a reference set to Microsoft Tablet PC Type Library 1.0. You're good to go!", "Reference already added", TDCBF_OK_BUTTON, TD_SHIELD_OK_ICON
    Case Else
      Dim Response As TDBUTTONS_RETURN_CODES
      Response = TaskBox("ERROR", "Error " & Err.Number & " - " & Err.Description & vbNewLine & vbNewLine & "Unable to add reference to Microsoft Tablet PC Type Library 1.0", "Failed to add reference", TDCBF_OK_BUTTON Or TDCBF_RETRY_BUTTON, TD_SHIELD_ERROR_ICON)
      If Response = IDRETRY Then Setup
  End Select

End Sub



