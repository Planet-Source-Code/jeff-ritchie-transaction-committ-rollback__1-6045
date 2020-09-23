VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmRollback 
   Caption         =   "Rollback Demonstration"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Rollback.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Status Bar"
      Top             =   2475
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "13/02/2000"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "1:00"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add GST Component"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Picture         =   "Rollback.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Update the Product Price list with GST"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblDescript 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Rollback.frx":1FD4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmRollback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables must be declared
Option Explicit
'
Private db As Database
Private rs As Recordset
Private ws As Workspace
Private sPath, sMsg, sHeading As String
'
Private Sub cmdAdd_Click()
    '
    On Error GoTo AddError
    '
    ' Are there any records to Process
    '
    If rs.EOF And rs.BOF Then
        sMsg = "No Product Records to Process"
        MsgBox sMsg, vbInformation, sHeading
        Exit Sub
    End If
    '
    ' Begin Transaction Focus, i.e all updates are in the one
    ' focus therefore can completely roll back or commit all
    ' individual transactions.
    '
    Dim cPrice, cTax As Currency
    Dim sAnswer As String
    Screen.MousePointer = vbHourglass
    StatusBar.Panels(1) = "Updating..."
    '
    ws.BeginTrans
    '
    rs.MoveFirst ' In case Record pointer is not at BOF
    Do
        rs.Edit
        cPrice = rs!Price
        cTax = cPrice * 0.1
        cPrice = cPrice + cTax
        rs!Price = cPrice
        rs!GST = cTax
        rs.Update ' Workspace not the actual database
        rs.MoveNext
        If rs.EOF Then Exit Do
    Loop
    '
    ' Does the Operator want to committ the modifications
    '
    sAnswer = MsgBox("Update the Price List with GST", _
        vbYesNoCancel + vbQuestion, sHeading)
    '
    Select Case sAnswer
        Case vbYes
            '
            ' Committ changes to database
            '
            ws.CommitTrans
        Case vbNo
            '
            ' Rollback changes to database
            '
            ws.Rollback
        Case vbCancel
            '
            ' Exit process
            '
            ws.Rollback
            Unload Me
    End Select
    Screen.MousePointer = vbNormal
    StatusBar.Panels(1) = "Complete..."
    Exit Sub
    '
AddError:
    Screen.MousePointer = vbNormal
    sMsg = Str(Err) & " : " & Err.Description
    MsgBox sMsg, vbCritical, sHeading
    Unload Me
    '
End Sub
'
Private Sub Form_Load()
    '
    On Error GoTo LoadError
    sHeading = "Database Rollback Demonstration"
    '
    ' Center form in Screen
    '
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    '
    ' Open Database
    '
    sPath = App.Path & "\Product.mdb"
    Set db = DBEngine.OpenDatabase(sPath)
    Set ws = DBEngine.Workspaces(0)
    Set rs = db.OpenRecordset("Pricing")
    '
    ' Ok inform the Operator everything is ok
    '
    StatusBar.Font.Bold = True
    StatusBar.Panels(1).Text = "Ready..."
    Exit Sub
    '
LoadError:
    sMsg = Str(Err) & " : " & Err.Description
    MsgBox sMsg, vbCritical, sHeading
    End
    '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    '
    ws.Close
    '
End Sub
