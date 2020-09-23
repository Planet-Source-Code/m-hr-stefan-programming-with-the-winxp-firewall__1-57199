Attribute VB_Name = "modMain"
Option Explicit

'** Initialize for XP-Style through Ressource File

Public Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" _
    (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()

    On Error Resume Next
    
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
    frmMain.Show
    
    On Error GoTo 0
    
End Sub




