VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMouse_Wheel 
   Caption         =   "A minimal example"
   ClientHeight    =   6330
   ClientLeft      =   30
   ClientTop       =   90
   ClientWidth     =   11100
   OleObjectBlob   =   "frmMouse_Wheel.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMouse_Wheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents IC As MSINKAUTLib.InkCollector
Attribute IC.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
  Frame1.ScrollHeight = 20000
  SetupMouseWheel
End Sub

Private Sub SetupMouseWheel()
  Set IC = New MSINKAUTLib.InkCollector
  With IC
    .hWnd = Me.Frame1.[_GethWnd]                    ' The InkCollector requires an 'anchor' hWnd
    .SetEventInterest ICEI_MouseWheel, True         ' This sets event that you want to listen for
    .MousePointer = IMP_Arrow                       ' If this is not set, the mouse pointer disappears
    .DynamicRendering = False                       ' I suggest turning this off
    .DefaultDrawingAttributes.Transparency = 255    ' And making the drawing fullly transparent
    .Enabled = True                                 ' This must be set last
  End With
End Sub

Private Sub IC_MouseWheel(ByVal Button As MSINKAUTLib.InkMouseButton, ByVal Shift As MSINKAUTLib.InkShiftKeyModifierFlags, ByVal Delta As Long, ByVal X As Long, ByVal Y As Long, Cancel As Boolean)
  If Delta > 0 Then
    Frame1.ScrollTop = Frame1.ScrollTop - 20 - (2 * Shift)
  Else
    Frame1.ScrollTop = Frame1.ScrollTop + 20 + (2 * Shift)
  End If
End Sub
