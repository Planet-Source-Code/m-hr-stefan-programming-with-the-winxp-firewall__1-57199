VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Windows - XP - Firewall"
   ClientHeight    =   7530
   ClientLeft      =   2835
   ClientTop       =   2625
   ClientWidth     =   11010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11010
   Begin VB.CheckBox chkAllowIncomingEcho 
      Caption         =   "Allow Incoming Echo Request"
      Height          =   375
      Left            =   5700
      TabIndex        =   8
      Top             =   960
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   11025
      TabIndex        =   7
      Top             =   0
      Width           =   11025
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Programming with the XP Fireall with VB6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   180
         Width           =   6705
      End
   End
   Begin VB.CommandButton cmdAddSelected 
      Caption         =   "&Add Selected Ports to Firewall"
      Height          =   465
      Left            =   60
      TabIndex        =   6
      Top             =   6960
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Caption         =   "Firewall Status"
      Height          =   1635
      Left            =   60
      TabIndex        =   0
      Top             =   870
      Width           =   5445
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'Kein
         Height          =   825
         Left            =   4350
         ScaleHeight     =   825
         ScaleWidth      =   1005
         TabIndex        =   3
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton cmdDisable 
         Caption         =   "&Disable Firewall"
         Height          =   465
         Left            =   1950
         TabIndex        =   2
         Top             =   870
         Width           =   1695
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "&Enable Firewall"
         Height          =   465
         Left            =   150
         TabIndex        =   1
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   4965
      End
   End
   Begin MSComctlLib.ListView lsvPort 
      Height          =   4305
      Left            =   60
      TabIndex        =   5
      Top             =   2580
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   7594
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--> Enable and Disable Firewall from WinXP SP2 Via NetCon 1.0 Typelibrary
'--> You need to have XPSP2 installed to use this

Private oFirewall As clsFirewall
Private lngPortCounter As Long

Private Sub chkAllowIncomingEcho_Click()
  If Me.chkAllowIncomingEcho.Value = True Then
    Call oFirewall.AllowIncomingICMP(True)
  Else
    Call oFirewall.AllowIncomingICMP(False)
  End If
End Sub

Private Sub cmdAddSelected_Click()

Dim itm As ListItem

  For Each itm In Me.lsvPort.ListItems
    If itm.Checked = True Then
      Call oFirewall.AddPortToFirewall(itm.SubItems(2), itm.SubItems(1), itm.Text)
    End If
  Next

End Sub

Private Sub cmdDisable_Click()
  
  oFirewall.DisableFirewall
  Call GetFirewallStatus
  
End Sub

Private Sub cmdEnable_Click()

  oFirewall.EnableFirewall
  Call GetFirewallStatus

End Sub

Private Sub Form_Load()

  Call initListview
  Set oFirewall = New clsFirewall
  Call GetFirewallStatus
  Call LoadPortList
  
End Sub

Private Sub GetFirewallStatus()

  If oFirewall.FirewallStatus = True Then
    Me.lblStatus.Caption = "Firewall enabled"
    Me.picStatus.Picture = LoadResPicture(101, vbResBitmap)
  Else
    Me.lblStatus.Caption = "Firewall disabled"
    Me.picStatus.Picture = LoadResPicture(102, vbResBitmap)
  End If
  
End Sub

Private Sub initListview()

  With Me.lsvPort
    .View = lvwReport
    .ColumnHeaders.Add Text:="Port", Width:=1000, Alignment:=AlignmentConstants.vbLeftJustify
    .ColumnHeaders.Add Text:="Type", Width:=1000, Alignment:=vbCenter
    .ColumnHeaders.Add Text:="Keyword", Width:=2000, Alignment:=vbCenter
    .ColumnHeaders.Add Text:="Description", Width:=2000, Alignment:=vbCenter
    .ColumnHeaders.Add Text:="Trojan Info", Width:=8000, Alignment:=vbCenter
    .Checkboxes = True
  End With

End Sub

Private Sub LoadPortList()

'--> Load Portlist via Filesystem Object

Dim oFso As FileSystemObject
Dim stream As TextStream
Dim strArray() As String
Dim oItm As ListItem

On Error GoTo errHandler

  Set oFso = New FileSystemObject
  Set stream = oFso.OpenTextFile(App.Path & "\Portlist.txt")

  While Not stream.AtEndOfStream
    lngPortCounter = lngPortCounter + 1
    strArray = Split(stream.ReadLine, "|")
    Set oItm = Me.lsvPort.ListItems.Add(Text:=strArray(0))
    oItm.SubItems(1) = strArray(1)
    oItm.SubItems(2) = strArray(2)
    oItm.SubItems(3) = strArray(3)
  Wend

  stream.Close
  Set stream = Nothing
  Set oFso = Nothing

Exit Sub

errHandler:
  MsgBox Err.Description
  Err.Clear
  If Not oFso Is Nothing Then Set oFso = Nothing
  If Not stream Is Nothing Then stream = Nothing

End Sub
