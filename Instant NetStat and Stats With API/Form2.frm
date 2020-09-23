VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Instant NetStat"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form2"
   ScaleHeight     =   3555
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   150
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update List"
      Height          =   375
      Left            =   6345
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Local IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Remote IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5265
      Picture         =   "Form2.frx":0000
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "Total In List:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error Resume Next
  Dim pTcpTable As MIB_TCPTABLE
  Dim pdwSize As Long
  Dim bOrder As Long
  Dim nRet As Long
  Dim i As Integer, s As String
  ListView1.ListItems.Clear
  DoEvents
  nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
  nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
  For i = 0 To pTcpTable.dwNumEntries - 1
    If pTcpTable.table(i).dwState - 1 <> MIB_TCP_STATE_LISTEN Then
    Set Item = ListView1.ListItems.Add(, , c_ip(pTcpTable.table(i).dwLocalAddr))
    Item.SubItems(1) = c_port(pTcpTable.table(i).dwLocalPort)
    Item.SubItems(2) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    Item.SubItems(3) = c_port(pTcpTable.table(i).dwRemotePort)
    Item.SubItems(4) = c_state(pTcpTable.table(i).dwState - 1)
    'Item.EnsureVisible
    Else
    Set Item = ListView1.ListItems.Add(, , c_ip(pTcpTable.table(i).dwLocalAddr))
    Item.SubItems(1) = c_port(pTcpTable.table(i).dwLocalPort)
    Item.SubItems(2) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    Item.SubItems(3) = "0"
    Item.SubItems(4) = c_state(pTcpTable.table(i).dwState - 1)
    'Item.EnsureVisible
    End If
  Next
  DoEvents
    Me.MousePointer = vbNormal
    Label2.Caption = "Netstat status as of: " & Date & " " & Time
    Command1.Enabled = True
End Sub

Private Sub Form_Load()
Call Command1_Click
End Sub

Private Sub Timer1_Timer()
Label1.Caption = "Total In List: " & ListView1.ListItems.Count
End Sub
