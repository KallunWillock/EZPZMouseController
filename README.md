# EZPZMouseController

**A VBA library for capturing MouseWheel and other mouse events in UserForms and controls using the InkCollector object, without subclassing or complex Win32 API calls.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) <!-- Optional: Add a license badge -->

## 🐭 Background & Motivation

Support for the `MouseWheel` event in VBA, particularly within UserForms, is inconsistent and often non-existent for many standard controls. While solutions involving Windows subclassing or API hooks exist, they can be complex, fragile, and/or difficult to implement correctly.

This library leverages the `InkCollector` object, part of the **Microsoft Tablet PC Type Library** (typically included with Windows). While primarily designed for ink and pen input, `InkCollector` can be attached to (almost) any window handle (`hWnd`) and configured to listen for standard mouse events, including `MouseMove`, `MouseDown`, `MouseUp`, `DblClick`, and more.

This method provides a relatively simple approach to:

*   Reliably detect mouse wheel movements over UserForm controls.
*   Capture other standard mouse events with detailed information (button state, shift keys, coordinates).
*   Implement custom responses like scrolling, zooming, or other interactions.
*   Avoid the pitfalls of subclassing and complex API manipulations.

## 🖱️ Key Features

*   **Mouse Wheel Event Handling:** Captures `MouseWheel` events, providing direction (`Delta`), modifier keys (`Shift`), and coordinates.
*   **Standard Mouse Events:** Captures `MouseDown`, `MouseUp`, `MouseMove`, `and `DblClick`.
*   **Cursor Events:** Captures `CursorInRange` and `CursorOutOfRange` (though their precise triggers require further investigation).
*   **No Subclassing/API Hooks:** Relies solely on the built-in `InkCollector` object. As such, this method is **Bitness Agnositc**. Also, it is **also compatible with TwinBasic** (with a slight adjustment to the code.)
*   **Simple Integration:** Requires only a reference to the Tablet PC library and half-dozen lines of code.
*   **Customisable:** Event handlers in your UserForm code allow for flexible, custom actions.
*   **Control Targeting:** Can be attached to specific controls that possess a window handle (`hWnd`). This includes API-generated controls!

## ⚠️ Requirements & Considerations

1.  **Microsoft Tablet PC Type Library Reference:** Your VBA project **must** have a reference set to this library.
    *   Typically located at: `C:\Program Files (x86)\Common Files\Microsoft Shared\ink\InkObj.dll
    *   *Avoid* referencing `.exd` files directly from user profiles (like `MSINKAUTLib.exd`); reference the primary `InkObj.dll` or the library name in the VBA References dialog.
2.  **Window Handle (`hWnd`) Requirement:** `InkCollector` needs to attach to a window handle (`hWnd`).
    *   Standard VBA UserForm controls like `Frame`, `ListBox`, `MultiPage`, and the `UserForm` itself *do* have an `hWnd`.
    *   Controls like `Label`, `Image`, `CommandButton` *do not* typically have their own `hWnd`.
    *   **Workaround:** To capture events over non-windowed controls, there are (at present) two potential solutions:
    (1) place them inside a windowed container (like a `Frame`) and attach the `InkCollector` to the container's `hWnd`; or,
    (2) attach the InkCollector to the UserForm itself and then use the mouse coordinates (`X`, `Y`) within the event handler to determine interaction with the child controls if needed. There is a demonstration of this second solution in the demo workbook.
3.  **Performance:** While testing indicates good stability, attaching many `InkCollector` instances simultaneously might have performance implications. Consider managing instances carefully, potentially binding/unbinding dynamically as focus changes.
4.  **Event Specificity:** While `MouseWheel` is the primary target, remember `InkCollector` *can* capture ink strokes if not configured correctly.

## 🐁 Setup
The workbook housing the code will need to have a reference added to the MS Tablet PC Type Library 1.0. This can be done either manually or programmatically.
1.  **The manual route**:
    *   Go to `Tools` > `References...`.
    *   Scroll down and check **`Microsoft Tablet PC Type Library 1.0`** (or similar version).
    *   If not listed, click `Browse...` and navigate to `C:\Program Files (x86)\Common Files\Microsoft Shared\ink\InkObj.dll` (adjust path if needed). Click `OK`.
2.  **The programmatic route**: The following will attempt to add a reference to the type library. It is included in the demo workbook code.

```vba
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
```

## 🚀 Usage Example

It's a silly example but the following demonstrates one way of handling the `MouseWheel` event to move a button around a frame.

1.  **Create a UserForm:** Add a new UserForm to your project.
2.  **Add Controls:** Add a `Frame` control (`Frame1`) to the UserForm, and in the centre of that, draw a `CommandButton` (`CommandButton1`). 
3.  **Add the following to UserForm:**

```vba
' --- In the UserForm's code module ---

Private WithEvents IC As MSINKAUTLib.InkCollector

Private Sub UserForm_Initialize()
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
  Dim TargetProperty As String, CurrentValue As Long
  If Shift Then TargetProperty = "Left" Else TargetProperty = "Top"
  CurrentValue = CallByName(CommandButton1, TargetProperty, VbGet)
  CallByName CommandButton1, TargetProperty, VbLet, IIf(Delta > 0, CurrentValue - 5, CurrentValue + 5)
End Sub
```
