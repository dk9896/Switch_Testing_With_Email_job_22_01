VERSION 5.00
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Begin VB.Form frmPrintLabel 
   Caption         =   "Label Printing"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   12150
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
      Height          =   7695
      Left            =   2400
      ScaleHeight     =   7635
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   2040
      Width           =   5655
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
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   3015
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
         Left            =   2160
         TabIndex        =   7
         Top             =   3720
         Width           =   3015
      End
      Begin VB.PictureBox Picture3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         ScaleHeight     =   1035
         ScaleWidth      =   4395
         TabIndex        =   4
         Top             =   120
         Width           =   4455
         Begin VB.ComboBox CboModelName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Model Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         Height          =   1335
         Left            =   480
         Picture         =   "frmPrintLabel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   1275
         Left            =   3000
         Picture         =   "frmPrintLabel.frx":BE86
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6120
         Width           =   2085
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "S.N0."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1800
         Width           =   855
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
