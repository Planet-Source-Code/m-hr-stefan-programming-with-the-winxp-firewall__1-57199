VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFirewall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ICSSC_DEFAULT = 0
Const CONNECTION_PUBLIC = 0
Const CONNECTION_PRIVATE = 1
Const CONNECTION_ALL = 2

Const NET_FW_IP_PROTOCOL_UDP = 17
Const NET_FW_IP_PROTOCOL_TCP = 6

Const NET_FW_SCOPE_ALL = 0
Const NET_FW_SCOPE_LOCAL_SUBNET = 1

Private oNetShareMgr As Object

'--> Return the Firewall Status
Public Function FirewallStatus() As Boolean

Dim bolStatus As Boolean
Dim oProfile As Object

On Error GoTo errHandler
    
    Set oNetShareMgr = CreateObject("HNetCfg.FwMgr")
    Set oProfile = oNetShareMgr.LocalPolicy.CurrentProfile

    If oProfile.FirewallEnabled = False Then
      bolStatus = False
    Else
      bolStatus = True
    End If
              
  FirewallStatus = bolStatus
  
  Exit Function
  
errHandler:
  FirewallStatus = False
  MsgBox "Error: " & Err.Description
  Err.Clear

End Function

'--> Enable Firewall
Public Sub EnableFirewall()

Dim oProfile As Object

  On Error GoTo ErrorHandler
        
  Set oNetShareMgr = CreateObject("HNetCfg.FwMgr")
  Set oProfile = oNetShareMgr.LocalPolicy.CurrentProfile

  If oProfile.FirewallEnabled = False Then
    oProfile.FirewallEnabled = True
  End If
  
  Set oProfile = Nothing
  Set oNetShareMgr = Nothing
    
  Exit Sub
    
ErrorHandler:
  MsgBox Err.Description
  Err.Clear

End Sub

'--> Disable Firewall
Public Sub DisableFirewall()

Dim oProfile As Object

  On Error GoTo ErrorHandler
        
  Set oNetShareMgr = CreateObject("HNetCfg.FwMgr")
  Set oProfile = oNetShareMgr.LocalPolicy.CurrentProfile

  If oProfile.FirewallEnabled = True Then
    oProfile.FirewallEnabled = False
  End If
  
  Set oProfile = Nothing
  Set oNetShareMgr = Nothing
    
  Exit Sub
    
ErrorHandler:
  MsgBox Err.Description
  Err.Clear

End Sub

'--> Add a new port to the Firewall Configuration
Public Sub AddPortToFirewall(ByVal strPortName As String, ByVal strPortProtocol As String, ByVal intPortNumber As Integer)

Dim oProfile As Object
Dim port As Object

On Error GoTo errHandler


  Set oNetShareMgr = CreateObject("HNetCfg.FwMgr")
  Set oProfile = oNetShareMgr.LocalPolicy.CurrentProfile
  Set port = CreateObject("HNetCfg.FWOpenPort")
  
  port.Name = strPortName
  If LCase(strPortProtocol) = "UDP" Then
    port.Protocol = NET_FW_IP_PROTOCOL_UDP
  Else
    port.Protocol = NET_FW_IP_PROTOCOL_TCP
  End If
  
  port.port = intPortNumber

  port.Scope = NET_FW_SCOPE_ALL

  port.Enabled = True

  oProfile.GloballyOpenPorts.Add port
  
  Set oProfile = Nothing
  Set port = Nothing
  Set oNetShareMgr = Nothing

  Exit Sub
  
errHandler:
  MsgBox Err.Description
  Err.Clear

End Sub

Public Sub AllowIncomingICMP(ByVal bolAllow As Boolean)

Dim oProfile As Object

On Error GoTo errHandler

  Set oNetShareMgr = CreateObject("HNetCfg.FwMgr")
  Set oProfile = oNetShareMgr.LocalPolicy.CurrentProfile
  oProfile.IcmpSettings.AllowInboundEchoRequest = bolAllow
  
  Set oProfile = Nothing
  Set oNetShareMgr = Nothing

Exit Sub

errHandler:
  MsgBox Err.Description
  Err.Clear

End Sub





