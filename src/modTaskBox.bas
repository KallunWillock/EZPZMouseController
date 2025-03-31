Attribute VB_Name = "modTaskBox"
                                                                                             
'    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'    |||||||||||||             TASKBOX (v1.3)            |||||||||||||
'    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'
'    AUTHOR:    Kallun Willock
'    PURPOSE:   A basic implementation of the TaskDialog.
'    URL:       https://github.com/KallunWillock/JustMoreVBA/blob/main/Boxes/modBox_TaskBox.bas
'    REFERENCE: http://vbnet.mvps.org/index.html?code/comdlg/taskdialog.htm
'    NOTES:     See TaskDialogIndirect for a more feature packed alternative
'               https://github.com/fafalone/cTaskDialog64
'    VERSION:   1.3        19/09/2025         Various corrections to code

  Option Explicit
  
  Public Enum TDICONS
      TD_NO_ICON = 0                              '  No icon - MainInstruction and Contents against a white background
      IDI_APPLICATION = 32512                     '  Generic icon of an application - imageres.dll - index 11
      
      TD_WARNING_ICON = -1                        '  vbExclamation equivalent
      TD_ERROR_ICON = -2                          '  vbCritical equivalent
      TD_INFORMATION_ICON = -3                    '  vbInformation equivalent
      IDI_QUESTION = 32512                        '  vbQuestion equivalent
  
      TD_SHIELD_ICON = -4                         '  Icon of a security shield
      TD_SHIELD_GRADIENT_ICON = -5                '  Icon of a security shield against a gradient blue/teal colour bar - default setting
      TD_SHIELD_WARNING_ICON = -6                 '  Exclamation point in shield icon against a gradient orange/yellow colour bar
      TD_SHIELD_ERROR_ICON = -7                   '  X in shield icon against a gradient red colour bar
      TD_SHIELD_OK_ICON = -8                      '  Tick in shield icon against a gradient green colour bar
      TD_SHIELD_GRAY_ICON = -9                    '  Icon of a security shield against a grey colour bar
  End Enum
  
  '   The Task Dialog allows for any combination from the common button set: OK, Yes, No, Cancel, Retry, Close
  
  Public Enum TDBUTTONS
      TDCBF_OK_BUTTON = &H1&                      '  Return: 1 (IDOK)
      TDCBF_YES_BUTTON = &H2&                     '  Return: 6 (IDYES)
      TDCBF_NO_BUTTON = &H4&                      '  Return: 7 (IDNO)
      TDCBF_CANCEL_BUTTON = &H8&                  '  Return: 2 (IDCANCEL)
      TDCBF_RETRY_BUTTON = &H10&                  '  Return: 4 (IDRETRY)
      TDCBF_CLOSE_BUTTON = &H20&                  '  Return: 8 (IDCLOSE)
  End Enum
  
  Public Enum TDBUTTONS_RETURN_CODES
      IDOK = 1
      IDCANCEL = 2
      IDRETRY = 4
      IDYES = 6
      IDNO = 7
      IDCLOSE = 8
  End Enum
  
  #If VBA7 Then
      Private Declare PtrSafe Function TaskDialog Lib "comctl32.dll" (ByVal hWndParent As LongPtr, ByVal hInstance As LongPtr, ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, ByVal dwCommonButtons As Long, ByVal pszIcon As LongPtr, pnButton As Long) As Long
  #Else
      Private Enum LongPtr
          [_]
      End Enum
      Private Declare Function TaskDialog Lib "comctl32.dll" (ByVal hwndParent As LongPtr, ByVal hInstance As LongPtr, ByVal pszWindowTitle As LongPtr, ByVal pszMainInstruction As LongPtr, ByVal pszContent As LongPtr, ByVal dwCommonButtons As Long, ByVal pszIcon As LongPtr, pnButton As Long) As Long
  #End If
  
  Public Function TaskBox(ByVal TaskBoxMainInstruction As String, Optional TaskBoxContent As String = "", _
                          Optional ByVal TaskBoxTitle As String = " ", _
                          Optional ByVal dwButtons As TDBUTTONS = TDCBF_OK_BUTTON, _
                          Optional ByVal lIcon As TDICONS = TD_SHIELD_GRADIENT_ICON, Optional ByVal hWndParent As LongPtr = -1) As TDBUTTONS
  
      Dim dwIcon              As LongPtr
      Dim pnButton            As Long
      Dim Result              As TDBUTTONS_RETURN_CODES
      
      Const IDPROMPT          As Long = &HFFFF&
      
      '  Make the IntResource
      dwIcon = IDPROMPT And lIcon
      
      '  From MSDN: "If you create a task dialog while a dialog box is present, use a handle to the dialog box as the hWndParent parameter.
      '              The hWndParent parameter should not identify a child window, such as a control in a dialog box."
      If hWndParent = -1 Then hWndParent = Application.hWnd
      Result = TaskDialog(hWndParent, 0&, StrPtr(TaskBoxTitle), StrPtr(TaskBoxMainInstruction), StrPtr(TaskBoxContent), dwButtons, dwIcon, pnButton)
  
      '   From VBNET MVPS: "The value of the button the user pressed is not returned as a result of the function call but rather as a parameter passed
      '   ByRef to the function. The return value of the call now represents success (0) or OUTOFMEMORY, INVALIDARG, or simply "FAIL"."
      TaskBox = pnButton
  
  End Function
  
  
