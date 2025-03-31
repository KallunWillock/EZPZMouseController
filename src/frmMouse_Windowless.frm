VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMouse_Windowless 
   Caption         =   "UserForm1"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "frmMouse_Windowless.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMouse_Windowless"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'    |||||||||||||    EZPZ MOUSECONTROLLER - DEMO 2      |||||||||||||
'    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'
'    AUTHOR:   Kallun Willock
'    NOTES:    This demonstrates how the InkController can be used with
'              windowless MSForms controls. It relies on attaching to
'              the UserForm's hWnd. Note that with the UserForm, you must
'              use the hWnd of the Client Area and not the UserForm
'              proper (as set out in the code below).
'
'              - The technique requires a reference to be set to
'                Microsoft Tablet PC Type Library, version 1.0.
'                "C:\Users\YourUserName\AppData\Roaming\Microsoft
'
'    VERSION:  1.0        31/03/2025         Uploaded to Github

Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
#Else
  Private Enum LongPtr
  [_]
  End Enum
  Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
  Private Declare Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
#End If
 
Private WithEvents IC As MSINKAUTLib.InkCollector
Attribute IC.VB_VarHelpID = -1
Private TargetControl As msforms.Control

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Initialize()
  SetupMouseWheel
  Label1.Picture = New StdPicture
  Me.TextBox1.SelStart = 0
End Sub

Private Sub SetupMouseWheel()
  Dim hWnd As LongPtr, TemphWnd As LongPtr
  Call IUnknown_GetWindow(Me, VarPtr(hWnd))
  Const GW_CHILD = 5
  TemphWnd = GetWindow(hWnd, GW_CHILD)
  Set IC = New MSINKAUTLib.InkCollector
  With IC
    .hWnd = TemphWnd                                ' The InkCollector requires an 'anchor' hWnd
    .SetEventInterest ICEI_MouseWheel, True         ' This sets event that you want to listen for
    .MousePointer = IMP_Arrow                       ' If this is not set, the mouse pointer disappears
    .DynamicRendering = False                       ' I suggest turning this off = less overhead
    .DefaultDrawingAttributes.Transparency = 255    ' And making the drawing fullly transparent
    .Enabled = True                                 ' This must be set last
  End With
End Sub

' When the mouse cursor moves over these controls, this will set the control as the target of the mousewheel event.

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Not Label1 Is TargetControl Then
    Set TargetControl = Label1
  End If
End Sub

Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Not TextBox1 Is TargetControl Then
    Set TargetControl = TextBox1
  End If
End Sub

' The MouseWheel event selects what type of control it is dealing with and then executes the custom actions accoringly.
' Here, I use CallByName to adjust the controls properties to avoid the headaches associated with the limitations found
' in the the generic MSForms.Control control.

Private Sub IC_MouseWheel(ByVal Button As MSINKAUTLib.InkMouseButton, ByVal Shift As MSINKAUTLib.InkShiftKeyModifierFlags, ByVal Delta As Long, ByVal X As Long, ByVal Y As Long, Cancel As Boolean)
  Select Case TypeName(TargetControl)
    Case "Label"
      CallByName TargetControl, "Caption", VbLet, "Delta: " & Delta
    Case "TextBox"
      Dim CurrentLine As Long
      CurrentLine = CallByName(TargetControl, "CurLine", VbGet)
      If CurrentLine = TextBox1.LineCount - 1 And Delta < 0 Then Exit Sub
      If CurrentLine = 0 And Delta > 0 Then Exit Sub
      CallByName TargetControl, "CurLine", VbLet, IIf(Delta > 0, CurrentLine - 1, CurrentLine + 1)
  End Select
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Set IC = Nothing
End Sub
