VERSION 5.00
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Begin VB.Form frmPrintLabel 
   Caption         =   "Label Printing"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin TextPrinter.JustPrinter JustPrinter1 
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   7095
      Left            =   2280
      ScaleHeight     =   7035
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   -240
      Width           =   8535
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   2400
         TabIndex        =   20
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtSwitchName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   2400
         TabIndex        =   14
         Top             =   6120
         Width           =   3495
      End
      Begin VB.TextBox txtLineCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   2400
         TabIndex        =   13
         Top             =   5400
         Width           =   3495
      End
      Begin VB.TextBox txtVendorCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   2400
         TabIndex        =   12
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtIndexAR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   2400
         TabIndex        =   11
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox txtPartNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   2400
         TabIndex        =   10
         Top             =   3240
         Width           =   3495
      End
      Begin VB.ComboBox CboModelName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtDatePr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   2400
         TabIndex        =   5
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtCopyNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   3600
         TabIndex        =   4
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         Height          =   975
         Left            =   6840
         Picture         =   "frmPrintLabel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   1035
         Left            =   6840
         Picture         =   "frmPrintLabel.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5520
         Width           =   1485
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Switch Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   21
         Top             =   6240
         Width           =   1605
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Line Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   19
         Top             =   5520
         Width           =   1485
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   18
         Top             =   4800
         Width           =   1605
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Index AR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   17
         Top             =   4080
         Width           =   1485
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   16
         Top             =   2640
         Width           =   1725
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   15
         Top             =   3360
         Width           =   1485
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Manual Print Screen"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   8295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPrintLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrintModelName As String

Private Sub Check1_Click()

'    If Check1.Value = "1" Then
'        Printtype = "2D"
'    Else
'        Printtype = "1D"
'    End If

End Sub

Private Sub CmdClose_Click()
    CopyLabel = False
    frmmenu.Show
    Unload Me
End Sub

Private Sub CmdPrint_Click()
If ValidEntry(1, 9999999, txtCopyNo) = False Then Exit Sub
    
    CopyLabel = True
    PrintLabel JustPrinter1
End Sub

Private Sub Form_Load()
    frmPrintLabel.WindowState = 2
    Picture1.BackColor = RGB(142, 167, 190)
    txtDatePr.Text = Date
    'txtDatePr.Locked = True
    LoadModelCombo CboModelName
    PrintModelName = GetSetting(App.Title, "PrintLastModel", "PrintLastModel")
    LastModel PrintModelNameModelName, CboModelName
    LoadSettingsData

End Sub
Private Sub CboModelName_Click()

PrintModelName = CboModelName.Text
SaveSetting App.Title, "PrintLastModel", "PrintLastModel", PrintModelName
'ModelPicture Image1, ModelName
End Sub

Private Sub LoadModelCombo(Combo As ComboBox)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    Combo.Clear
    Sql = "Select * from Model_Set order by ModelName"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Do While Rs.EOF = False
        Combo.AddItem Rs("ModelName")
        Rs.MoveNext
    Loop
    
End Sub

Private Sub LastModel(ByVal Model As String, Combo As ComboBox)
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * from Model_Set where ModelName='" & Model & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If Rs.EOF = False Then
        Combo.Text = Model
    Else
        Combo.ListIndex = 0
    End If

End Sub

Private Sub LoadSettingsData()
On Error GoTo Error
Dim Str() As String
Dim Rs As ADODB.Recordset
Dim Sql As String


    Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
        
    txtModelDesc = Rs("ModelDesc")
    PartNo = Rs("PrintPartNo")
    BarcodeLength = Rs("BarcodeLength")
    HardwareNo = Rs("HardwareNo")
    SerialStartingtxt = Rs("SerialStartingtxt")
    VendorId = Rs("VendorId")

    PrintSwitchName = Rs("PrintSwitchName")
    PrintLineCode = Rs("PrintLineCode")
    
    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    PrinterName = Rs("PrinterName1")
    
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadSettingsData"
Resume Next
End Sub
Private Function ValidEntry(Min, Max As Double, Text As TextBox) As Boolean

    If Trim(Text) = "" Or (Val(Text) < Min Or Val(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max), vbInformation
        Text.SetFocus
'        Text.BackColor = vbRed
        ValidEntry = False
    Else
'        Text.BackColor = vbWhite
        ValidEntry = True
    End If

End Function

Private Sub Label3_Click()

End Sub
