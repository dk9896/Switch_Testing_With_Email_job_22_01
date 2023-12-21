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
      BackColor       =   &H00FFFFC0&
      Height          =   8415
      Left            =   4080
      ScaleHeight     =   8355
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   840
      Width           =   8535
      Begin VB.Frame FramePrint2 
         Caption         =   "Printer Detail"
         ForeColor       =   &H000040C0&
         Height          =   4815
         Left            =   240
         TabIndex        =   30
         Top             =   3120
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtPiaggioPartNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   39
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox txtSupplierCode 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   38
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txtRevisionNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   37
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtSupplierPartNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   36
            Top             =   2400
            Width           =   3495
         End
         Begin VB.TextBox txtFinalApproval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   35
            Top             =   2880
            Width           =   3495
         End
         Begin VB.TextBox txtBatchNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   34
            Top             =   3360
            Width           =   3495
         End
         Begin VB.TextBox txtOtherInfo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   33
            Top             =   3840
            Width           =   3495
         End
         Begin VB.TextBox txtSwitchName 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txtCOO 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   31
            Top             =   4320
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Piaggio Part No"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   1845
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Code"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   113
            Left            =   120
            TabIndex        =   47
            Top             =   1920
            Width           =   1845
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Rev P/N"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   114
            Left            =   120
            TabIndex        =   46
            Top             =   1560
            Width           =   1845
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier PartNo / Rev"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   115
            Left            =   120
            TabIndex        =   45
            Top             =   2400
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Final Approval"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   119
            Left            =   120
            TabIndex        =   44
            Top             =   2880
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch No"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   81
            Left            =   120
            TabIndex        =   43
            Top             =   3360
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Other Info"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   120
            Left            =   120
            TabIndex        =   42
            Top             =   3840
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   6000
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Switch Name"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   122
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1845
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Country Of Origin"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   121
            Left            =   120
            TabIndex        =   40
            Top             =   4320
            Width           =   1605
            WordWrap        =   -1  'True
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmPrintLabel.frx":0000
         Left            =   2400
         List            =   "frmPrintLabel.frx":000D
         TabIndex        =   25
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Check to print multiple in serial"
         Height          =   375
         Left            =   6840
         TabIndex        =   23
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   360
         Left            =   7440
         TabIndex        =   22
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtStartString 
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
         ForeColor       =   &H80000009&
         Height          =   480
         Left            =   2400
         TabIndex        =   20
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtQRPartNo 
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
         Height          =   480
         Left            =   2400
         TabIndex        =   14
         Top             =   6120
         Width           =   3495
      End
      Begin VB.TextBox txtBalPartNo 
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
         Height          =   480
         Left            =   4080
         TabIndex        =   11
         Top             =   3960
         Width           =   1815
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
         Height          =   480
         Left            =   2880
         TabIndex        =   10
         Top             =   3240
         Width           =   3015
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
         Top             =   720
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
         Height          =   480
         Left            =   3600
         TabIndex        =   4
         Text            =   "0"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         Height          =   975
         Left            =   6840
         Picture         =   "frmPrintLabel.frx":0033
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   1035
         Left            =   6840
         Picture         =   "frmPrintLabel.frx":069D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5520
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Code"
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
         TabIndex        =   29
         Top             =   6240
         Width           =   1605
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         TabIndex        =   28
         Top             =   5520
         Width           =   1845
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode Type"
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
         Left            =   360
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inc ++"
         Height          =   255
         Left            =   6840
         TabIndex        =   24
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "QR Part No"
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
         Caption         =   "Bal Part No"
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
         Caption         =   "Index AR/Revision No"
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
         Width           =   3045
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
         Width           =   1845
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
         Top             =   840
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
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   5520
      TabIndex        =   27
      Top             =   4200
      Width           =   1215
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
If Check1.Value = 1 Then
    For i = 0 To Val(Text1.Text)
        CopyLabel = True
        PrintLabel JustPrinter1
        txtCopyNo.Text = Val(txtCopyNo.Text) + 1
    Next
Else
    CopyLabel = True
    PrintLabel JustPrinter1
    txtCopyNo.Text = Val(txtCopyNo.Text) + 1
End If
End Sub

Private Sub Combo1_Change()
    barcodeType = Combo1.ListIndex
    If barcodeType = 1 Then
    txtDatePr.Text = GetCurrentDate2
    txtStartString.Visible = False
    txtQRPartNo.Visible = True
    txtBalPartNo.Visible = True
    txtIndexAR.Visible = True
    Label7.Visible = True
    Label9.Visible = False
    Label10.Visible = False
    Label12.Visible = True
    Label14.Visible = True
    Label5.Caption = "Part Number"
    ElseIf barcodeType = 0 Then
    txtDatePr.Text = GetCurrentDate
    txtStartString.Visible = True
    txtQRPartNo.Visible = True
    txtBalPartNo.Visible = True
    txtIndexAR.Visible = True
    Label7.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Label12.Visible = False
    Label14.Visible = False
    Label5.Caption = "Part Number"
    Else
    txtDatePr.Text = Format(Now, "ddMMyy")
    txtBalPartNo.Visible = True
    Label12.Visible = True
    txtIndexAR.Visible = True
    Label7.Visible = True
    txtPartNumber.Visible = True
    Label5.Visible = True
    Label5.Caption = "Bal Part No"
    
    End If
End Sub

Private Sub Combo1_Click()
    barcodeType = Combo1.ListIndex
    If barcodeType = 1 Then
    txtDatePr.Text = GetCurrentDate2
    txtStartString.Visible = False
    txtQRPartNo.Visible = True
    txtBalPartNo.Visible = True
    txtIndexAR.Visible = True
    Label7.Visible = True
    Label9.Visible = False
    Label10.Visible = False
    Label12.Visible = True
    Label14.Visible = True
    Label8.Visible = True
    txtVendorCode.Visible = True
    Label5.Caption = "Part Number"
    ElseIf barcodeType = 0 Then
    Label8.Visible = True
    txtVendorCode.Visible = True
    txtDatePr.Text = GetCurrentDate
    txtStartString.Visible = True
    txtQRPartNo.Visible = True
    txtBalPartNo.Visible = True
    txtIndexAR.Visible = True
    Label7.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Label12.Visible = False
    Label14.Visible = False
    Label5.Caption = "Part Number"
    Else
    txtDatePr.Text = Format(Now, "ddMMyy")
    txtBalPartNo.Visible = True
    Label12.Visible = True
    txtIndexAR.Visible = True
    Label7.Visible = True
    txtPartNumber.Visible = True
    Label5.Visible = True
    Label5.Caption = "Bal Part No"
    Label9.Visible = False
    txtStartString.Visible = False
    Label8.Visible = False
    txtVendorCode.Visible = False
    Label14.Visible = False
    txtQRPartNo.Visible = False
    Label10.Visible = False
    End If
End Sub

Private Sub Form_Load()
    frmPrintLabel.WindowState = 2
    Picture1.BackColor = RGB(142, 167, 190)
    txtDatePr.Text = GetCurrentDate
    'txtDatePr.Locked = True
    LoadModelCombo cbomodelname
    PrintModelName = GetSetting(App.Title, "PrintLastModel", "PrintLastModel")
    LastModel PrintModelNameModelName, cbomodelname
    LoadSettingsData

End Sub
Private Sub CboModelName_Click()

PrintModelName = cbomodelname.Text
SaveSetting App.Title, "PrintLastModel", "PrintLastModel", PrintModelName
ModelName = cbomodelname.Text
LoadSettingsData
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
        
    txtPartNumber.Text = Rs("PrintPartNo")
    txtIndexAR.Text = Rs("HardwareNo")
    txtStartString.Text = Format(Now, "YYWW")
    txtVendorCode.Text = Rs("VendorId")
    txtBalPartNo.Text = Rs("SerialStartingtxt")
    txtQRPartNo.Text = Rs("QRPartNo")
    barcodeType = Val(Rs("BarcodeType"))
    Combo1.ListIndex = barcodeType
    If barcodeType = 1 Then
    txtBalPartNo.Text = Rs("SupplierCode")
    txtQRPartNo.Text = GetShiftInAsc(getShift)
    'txtDatePr.Text = GetCurrentDate2
    'txtStartString.Visible = False
    'txtQRPartNo.Visible = False
    'txtBalPartNo.Visible = False
    'txtIndexAR.Visible = False
    'Label7.Visible = False
    'Label9.Visible = False
    'Label10.Visible = False
    ElseIf barcodeType = 2 Then
    txtBalPartNo.Text = Rs("SupplierCode")
    txtPartNumber.Text = Rs("SerialStartingtxt")
    
    'txtDatePr.Text = GetCurrentDate
    'txtStartString.Visible = True
    'txtQRPartNo.Visible = True
    'txtBalPartNo.Visible = True
    'txtIndexAR.Visible = True
    'Label7.Visible = True
    'Label9.Visible = True
    'Label10.Visible = True
    End If
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

