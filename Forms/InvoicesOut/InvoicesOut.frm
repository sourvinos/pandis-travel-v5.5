VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{89D94A1E-DB65-4469-AFB5-D54C6F6B7639}#1.1#0"; "QRCodeAX.ocx"
Begin VB.Form InvoicesOut 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   11550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11550
   ScaleWidth      =   18105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   0
      TabIndex        =   98
      Top             =   10125
      Width           =   10365
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "����������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   6
         Left            =   8775
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8421631
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "��������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "����������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   3
         Left            =   4500
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "��������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   4
         Left            =   5925
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   5
         Left            =   7350
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "�����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   3075
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "��������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6840
      Left            =   12525
      TabIndex        =   53
      Top             =   675
      Width           =   4515
      Begin VB.TextBox txtInvoiceIsActiveID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   3075
         Width           =   780
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   109
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceIsActiveID"
         Top             =   3075
         Width           =   3540
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   97
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceTrnID"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceTrnID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceDateIn"
         Top             =   2700
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceDateIn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   2700
         Width           =   780
      End
      Begin VB.TextBox txtInvoiceMasterRefersTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceMasterRefersTo"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   91
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceSecondaryRefersTo"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceSecondaryRefersTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txtShipRegistryNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   6075
         Width           =   780
      End
      Begin VB.CheckBox chkCodeHandID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   5700
         Width           =   780
      End
      Begin VB.TextBox txtCodeLastDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   4950
         Width           =   780
      End
      Begin VB.TextBox txtCodeLastNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   4575
         Width           =   780
      End
      Begin VB.CheckBox chkPaymentTermCreditID 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   6450
         Width           =   780
      End
      Begin VB.TextBox txtInvoiceCodeID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1950
         Width           =   780
      End
      Begin VB.TextBox txtVATPercent 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox txtInvoiceID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox txtCodePersonsPlusOrMinus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   5325
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   70
         TabStop         =   0   'False
         Text            =   "Codes.CodeCustomers"
         Top             =   5325
         Width           =   3540
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   "Settings.VAT"
         Top             =   825
         Width           =   3540
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   67
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceCodeID"
         Top             =   1950
         Width           =   3540
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   "Codes.CodeLastNo"
         Top             =   4575
         Width           =   3540
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   "Codes.CodeLastDate"
         Top             =   4950
         Width           =   3540
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "Ships.ShipRegistryNo"
         Top             =   6075
         Width           =   3540
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "Codes.CodeHandID"
         Top             =   5700
         Width           =   3540
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "PaymentTerms.PaymentTermCredit"
         Top             =   6450
         Width           =   3540
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "Invoices.InvoicePersonID"
         Top             =   2325
         Width           =   3540
      End
      Begin VB.TextBox txtInvoicePersonID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2325
         Width           =   780
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "InvoicesOut.InvoiceOutDestinationID"
         Top             =   3450
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceOutDestinationID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3450
         Width           =   780
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "InvoicesOut.InvoiceOutShipID"
         Top             =   3825
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceOutShipID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3825
         Width           =   780
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   75
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "InvoicesOut.InvoiceOutPaymentTermID"
         Top             =   4200
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceOutPaymentTermID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3675
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4200
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ��������� ���� "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2640
      Left            =   11175
      TabIndex        =   52
      Tag             =   "SameColorAsBackground"
      Top             =   5250
      Width           =   2115
      Begin UserControls.newFloat mskDirectAmount 
         Height          =   465
         Left            =   450
         TabIndex        =   18
         Top             =   1950
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   8
         Left            =   1650
         Top             =   1875
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   7
         Left            =   0
         Top             =   1950
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " �� �������� "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2640
      Index           =   0
      Left            =   2175
      TabIndex        =   42
      Tag             =   "SameColorAsBackground"
      Top             =   5250
      Width           =   2940
      Begin UserControls.newInteger mskAdultsWithTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   8
         Top             =   375
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskKidsWithTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   10
         Top             =   900
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskFreeWithTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   12
         Top             =   1425
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskTotalPersonsWithTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1950
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   8421376
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskAdultsAmountWithTransfer 
         Height          =   465
         Left            =   1275
         TabIndex        =   9
         Top             =   375
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskKidsAmountWithTransfer 
         Height          =   465
         Left            =   1275
         TabIndex        =   11
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskTotalAmountWithTransfer 
         Height          =   465
         Left            =   1275
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1950
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   8421376
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   1
         Left            =   0
         Top             =   1125
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   5
         Left            =   2475
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ����� �������� "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2640
      Left            =   5175
      TabIndex        =   39
      Tag             =   "SameColorAsBackground"
      Top             =   5250
      Width           =   2940
      Begin UserControls.newInteger mskAdultsWithoutTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   13
         Top             =   375
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskKidsWithoutTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   15
         Top             =   900
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskFreeWithoutTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   17
         Top             =   1425
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskTotalPersonsWithoutTransfer 
         Height          =   465
         Left            =   450
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1950
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   8421376
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskAdultsAmountWithoutTransfer 
         Height          =   465
         Left            =   1275
         TabIndex        =   14
         Top             =   375
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskKidsAmountWithoutTransfer 
         Height          =   465
         Left            =   1275
         TabIndex        =   16
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         ForeColor       =   0
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskTotalAmountWithoutTransfer 
         Height          =   465
         Left            =   1275
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1950
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   8421376
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   3
         Left            =   2475
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   2
         Left            =   0
         Top             =   975
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.CheckBox chkAgreement 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "������ �� ��������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   2175
      TabIndex        =   7
      Top             =   4875
      Width           =   2640
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ������ "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2640
      Left            =   8175
      TabIndex        =   31
      Tag             =   "SameColorAsBackground"
      Top             =   5250
      Width           =   2940
      Begin UserControls.newInteger mskTotalAdults 
         Height          =   465
         Left            =   450
         TabIndex        =   32
         Top             =   375
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskTotalKids 
         Height          =   465
         Left            =   450
         TabIndex        =   33
         Top             =   900
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskTotalFree 
         Height          =   465
         Left            =   450
         TabIndex        =   34
         Top             =   1425
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newInteger mskTotalPersons 
         Height          =   465
         Left            =   450
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1950
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   8421376
         MaxLength       =   6
         Text            =   "99.999"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskAdultsAmountTotal 
         Height          =   465
         Left            =   1275
         TabIndex        =   36
         Top             =   375
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskKidsAmountTotal 
         Height          =   465
         Left            =   1275
         TabIndex        =   37
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   0
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newFloat mskTotalAmount 
         Height          =   465
         Left            =   1275
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1950
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Enabled         =   0   'False
         Alignment       =   1
         ForeColor       =   8421376
         Text            =   "99.999,99"
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   6
         Left            =   2475
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   4
         Left            =   0
         Top             =   375
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin UserControls.newDate mskDateIssue 
      Height          =   465
      Left            =   2175
      TabIndex        =   0
      Top             =   1125
      Width           =   1455
      _ExtentX        =   2672
      _ExtentY        =   820
      ForeColor       =   0
      Text            =   "01/01/2017"
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtCodeShortDescriptionA 
      Height          =   465
      Left            =   2175
      TabIndex        =   2
      Top             =   2175
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   3
      Text            =   "���"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtCustomerDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   4
      Top             =   3225
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   100
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtDestinationDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   5
      Top             =   3750
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "����������������������������������������"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtShipDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   6
      Top             =   4275
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtRemarks 
      Height          =   465
      Left            =   2175
      TabIndex        =   19
      Top             =   8100
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtPaymentTermDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   20
      Top             =   8625
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   2925
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   2175
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   3375
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   2175
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7200
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   3225
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   7200
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   3750
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":10CE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   6
      Left            =   7200
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   4275
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":1668
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   8
      Left            =   7200
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8625
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":1C02
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   7650
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   3225
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":219C
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   7650
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   3750
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":2736
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   7
      Left            =   7650
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   4275
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":2CD0
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   9
      Left            =   7650
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   8625
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":326A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtInvoiceNo 
      Height          =   465
      Left            =   2175
      TabIndex        =   3
      Top             =   2700
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   10
      Text            =   "9999999999"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newDate mskDateRefersTo 
      Height          =   465
      Left            =   2175
      TabIndex        =   1
      Top             =   1650
      Width           =   1455
      _ExtentX        =   2672
      _ExtentY        =   820
      ForeColor       =   0
      Text            =   "01/01/2017"
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtInvoiceIsActiveDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   21
      Top             =   9150
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      Text            =   "���"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   10
      Left            =   2850
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   9150
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   820
      BackColor       =   14742518
      ButtonShape     =   3
      ButtonStyle     =   8
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      PicNormal       =   "InvoicesOut.frx":3804
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin QRCodeAX.QR imgQRCode 
      Height          =   1815
      Left            =   10725
      TabIndex        =   111
      Top             =   8175
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3201
      DataString      =   "iosadfi3i2fhsdahlkhsaklfsdiosadfi3i2fhsdahlkhsaklfsd"
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "������ �������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   450
      TabIndex        =   108
      Top             =   9225
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   450
      TabIndex        =   106
      Top             =   1725
      Width           =   1290
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   9
      Left            =   0
      Top             =   9525
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   2100
      Top             =   9600
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   13
      Left            =   2850
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   4800
      Top             =   10800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   13275
      Top             =   6150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   0
      Left            =   1725
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   12
      Left            =   0
      Top             =   2325
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblHand 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "�������������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   3825
      TabIndex        =   51
      Top             =   2400
      Width           =   1350
   End
   Begin VB.Label lblCodeBatch 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "����� �"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   5250
      TabIndex        =   50
      Top             =   2400
      Width           =   585
   End
   Begin VB.Label lblCodeDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   3825
      TabIndex        =   49
      Top             =   2175
      Width           =   4200
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   13
      Left            =   450
      TabIndex        =   48
      Top             =   7275
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   12
      Left            =   450
      TabIndex        =   47
      Top             =   6750
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   11
      Left            =   450
      TabIndex        =   46
      Top             =   6225
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   450
      TabIndex        =   45
      Top             =   5700
      Width           =   1290
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   720
      Left            =   225
      TabIndex        =   30
      Top             =   75
      Width           =   2250
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "�����������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   450
      TabIndex        =   29
      Top             =   2250
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "�� ������������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   450
      TabIndex        =   28
      Top             =   2775
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   450
      TabIndex        =   27
      Top             =   3825
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   450
      TabIndex        =   26
      Top             =   4350
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   450
      TabIndex        =   25
      Top             =   8175
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "���� ��������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   450
      TabIndex        =   24
      Top             =   8700
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "��������������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   450
      TabIndex        =   23
      Top             =   3300
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   450
      TabIndex        =   22
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   0
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "InvoicesOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim blnCancel As Boolean
Dim blnPrinterHasBeenSelected As Boolean
Dim lngTrnID As Long
Dim IsError As Boolean

Private Function AddAmounts()

    On Error GoTo ErrTrap

    'Variables
    Dim cutTotalAmountWithTransfer As Currency
    Dim curTotalAmountWithoutTransfer As Currency
    Dim curAdultsAmountTotal As Currency
    Dim curKidsAmountTotal As Currency
    Dim curTotalAmount As Currency
    
    '������
    cutTotalAmountWithTransfer = CCur(mskAdultsAmountWithTransfer.text) + CCur(mskKidsAmountWithTransfer.text)
    curTotalAmountWithoutTransfer = CCur(mskAdultsAmountWithoutTransfer.text) + CCur(mskKidsAmountWithoutTransfer.text)
    
    '���������
    curAdultsAmountTotal = CCur(mskAdultsAmountWithTransfer.text) + CCur(mskAdultsAmountWithoutTransfer.text)
    curKidsAmountTotal = CCur(mskKidsAmountWithTransfer.text) + CCur(mskKidsAmountWithoutTransfer.text)
    curTotalAmount = curAdultsAmountTotal + curKidsAmountTotal
    
    '�������� - ������
    mskTotalAmountWithTransfer.text = format(cutTotalAmountWithTransfer, "#,##0.00")
    mskTotalAmountWithoutTransfer.text = format(curTotalAmountWithoutTransfer, "#,##0.00")
    
    '�������� - ���������
    mskAdultsAmountTotal.text = format(curAdultsAmountTotal, "#,##0.00")
    mskKidsAmountTotal.text = format(curKidsAmountTotal, "#,##0.00")
    mskTotalAmount.text = format(curTotalAmount, "#,##0.00")
    
    Exit Function
    
ErrTrap:
    If Err.Number = 13 Then
        Resume Next
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
End Function

Private Function AddPersons()

    '������
    Dim lngTotalPersonsWithTransfer As Long
    Dim lngTotalPersonsWithoutTransfer As Long
    
    '���������
    Dim lngTotalAdults As Long
    Dim lngTotalKids As Long
    Dim lngTotalFree As Long
    Dim lngTotalPersons As Long
    
    '������
    lngTotalPersonsWithTransfer = CCur(mskAdultsWithTransfer.text) + CCur(mskKidsWithTransfer.text) + CCur(mskFreeWithTransfer.text)
    lngTotalPersonsWithoutTransfer = CCur(mskAdultsWithoutTransfer.text) + CCur(mskKidsWithoutTransfer.text) + CCur(mskFreeWithoutTransfer.text)
    
    '���������
    lngTotalAdults = CCur(mskAdultsWithTransfer.text) + CCur(mskAdultsWithoutTransfer.text)
    lngTotalKids = CCur(mskKidsWithTransfer.text) + CCur(mskKidsWithoutTransfer.text)
    lngTotalFree = CCur(mskFreeWithTransfer.text) + CCur(mskFreeWithoutTransfer.text)
    lngTotalPersons = lngTotalAdults + lngTotalKids + lngTotalFree
    
    '��������
    mskTotalPersonsWithTransfer.text = format(lngTotalPersonsWithTransfer, "#,##0")
    mskTotalPersonsWithoutTransfer.text = format(lngTotalPersonsWithoutTransfer, "#,##0")
    mskTotalAdults.text = format(lngTotalAdults, "#,##0")
    mskTotalKids.text = format(lngTotalKids, "#,##0")
    mskTotalFree.text = format(lngTotalFree, "#,##0")
    mskTotalPersons.text = format(lngTotalPersons, "#,##0")

End Function

Private Sub AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            blnCancel = True
            
            ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtInvoiceOutDestinationID, txtInvoiceOutShipID, txtInvoiceOutPaymentTermID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID, txtShipRegistryNo, chkPaymentTermCreditID, txtInvoiceIsActiveID
            ClearFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, lblCodeDescription, lblCodeBatch, lblHand, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
            
            DisableFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9), cmdIndex(10)
            
            UpdateButtons Me, 6, 1, 0, 0, 0, IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), 0, 1
        End If
        Exit Sub
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Sub

Private Function CheckForTheLastInvoice()

    '�� ��� ����� �� ��������� ����������� � �� ����� ����������
    If txtCodeLastNo.text = Int(txtInvoiceNo.text) Or chkCodeHandID.Value = 1 Then
        CheckForTheLastInvoice = True '����� �� ��������
    Else
        CheckForTheLastInvoice = False '��� ����� �� ��������
    End If

End Function

Private Function ClearInvoiceFields()

    With rptInvoice
        .lblAdultsWithTransfer.Caption = ""
        .lblAdultAmountWithTransfer.Caption = ""
        .lblAdultsTotalWithTransfer.Caption = ""
        .lblKidsWithTransfer.Caption = ""
        .lblKidAmountWithTransfer.Caption = ""
        .lblKidsTotalWithTransfer.Caption = ""
        .lblAdultsWithoutTransfer.Caption = ""
        .lblAdultAmountWithoutTransfer.Caption = ""
        .lblAdultsTotalWithoutTransfer.Caption = ""
        .lblKidsWithoutTransfer.Caption = ""
        .lblKidAmountWithoutTransfer.Caption = ""
        .lblKidsTotalWithoutTransfer.Caption = ""
        .lblFree.Caption = ""
        .lblTotalPersons.Caption = ""
        .lblTotalAmount.Caption = ""
    End With

End Function

Private Function CreateQRCode()

    imgQRCode.DataString = strQRCodeDescription

End Function

Private Sub DeleteRecord()
    
    BeginTrans
    
    If MainDeleteRecord("CommonDB", "Invoices", strApplicationName, "InvoiceID", txtInvoiceID.text, True) Then
        If MainDeleteRecord("CommonDB", "InvoicesOut", strApplicationName, "InvoiceOutTrnID", txtInvoiceTrnID.text, False) Then
            CommitTrans
            blnCancel = True
            ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtInvoiceOutDestinationID, txtInvoiceOutShipID, txtInvoiceOutPaymentTermID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID, txtShipRegistryNo, chkPaymentTermCreditID, txtInvoiceIsActiveID
            ClearFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, lblCodeDescription, lblCodeBatch, lblHand, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, , mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
            DisableFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9), cmdIndex(10)
            UpdateButtons Me, 6, 1, 0, 0, 0, IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), 0, 1
        Else
            Rollback
        End If
    Else
        Rollback
    End If
    
End Sub

Public Function DoPostFoundJobs(rstRecordset As Recordset)

    On Error GoTo ErrTrap

    blnStatus = False

    DisableFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, _
        chkAgreement, _
        mskAdultsWithTransfer, mskAdultsAmountWithTransfer, mskKidsWithTransfer, mskKidsAmountWithTransfer, mskFreeWithTransfer, _
        mskAdultsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsWithoutTransfer, mskKidsAmountWithoutTransfer, mskFreeWithoutTransfer, _
        mskDirectAmount, _
        txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
    blnCancel = False
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9), cmdIndex(10)
    PopulateFields rstRecordset
    AddAmounts
    AddPersons
    EnableOrDisableFields
    'Temp use only!
    UpdateButtons Me, 6, 0, 1, 0, IIf(CheckForTheLastInvoice = 0, 0, 1), IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), 1, 0
    'UpdateButtons Me, 6, IIf(chkCodeHandID.Value = 1, 0, 1), IIf(chkCodeHandID.Value = 1, 1, 0), IIf(chkCodeHandID.Value = 1, 0, 1), IIf(CheckForTheLastInvoice = 0, 0, 1), IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), IIf(chkCodeHandID.Value = 1, 1, 0), IIf(chkCodeHandID.Value = 1, 0, 1)
        
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function

Private Function HideQRCode()

    rptInvoice.shpQRCode.Visible = False
    rptInvoice.imgQRCode.Visible = False

End Function

Private Function InsertQRCode()

    Set rptInvoice.imgQRCode = LoadPicture(strReportsPathName & "QRCode.bmp")

End Function

Private Sub NewRecord()
    
    Dim tmpRecordset As Recordset
    
    blnStatus = True
    blnCancel = False
    
    ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtInvoiceOutDestinationID, txtInvoiceOutShipID, txtInvoiceOutPaymentTermID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID, txtShipRegistryNo, chkPaymentTermCreditID
    ClearFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, lblCodeDescription, lblCodeBatch, lblHand, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
    
    EnableFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, txtCodeLastNo, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, _
        mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, _
        mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, _
        mskDirectAmount, _
        txtRemarks, _
        txtPaymentTermDescription, _
        txtInvoiceIsActiveDescription, _
        cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9), cmdIndex(10)
    
    InitializeFields mskDateIssue, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount
    
    mskDateIssue.text = format(Date, "dd/mm/yyyy")
    txtInvoiceDateIn.text = Date
    txtInvoiceOutShipID.text = IIf(txtInvoiceSecondaryRefersTo.text = "1", "0", "")
    txtRemarks.text = strUsualRemarks
    txtVATPercent.text = intVAT
    
    '������� ���� ��������
    Set tmpRecordset = CheckForMatch("CommonDB", "PaymentTerms", "PaymentTermID", "Numeric", intUsualPaymentTermID)
    If tmpRecordset.RecordCount > 0 Then
        txtInvoiceOutPaymentTermID.text = tmpRecordset.Fields(0)
        txtPaymentTermDescription.text = tmpRecordset.Fields(1)
    End If
    
    txtInvoiceIsActiveID.text = "1"
    txtInvoiceIsActiveDescription.text = "���"
    
    UpdateButtons Me, 6, 0, 1, 0, 0, 0, 1, 0
    
    mskDateIssue.SetFocus
    
End Sub

Private Function AskToPrintInvoice()

    Dim arrDummy()
    
    '������� ��� �������� �� ��������� ��� �� ����� �� ��� �������
    'Temp use only!
    'If chkCodeHandID.Value = 0 And blnStatus Then
    If chkCodeHandID.Value = 0 Then
        If MyMsgBox(2, strApplicationName, strAppMessages(7), 2) Then
            ProcessSelectedInvoicesForPrinting txtInvoiceTrnID.text, "Print", arrDummy, arrDummy
        End If
    End If

End Function

Private Function PopulateFields(rstRecordset As Recordset)

    With rstRecordset
    
        txtInvoiceMasterRefersTo.text = !InvoiceMasterRefersTo
        txtInvoiceSecondaryRefersTo.text = !InvoiceSecondaryRefersTo
        txtVATPercent.text = !InvoiceOutVATPercent
        txtInvoiceID.text = !InvoiceID
        txtInvoiceTrnID.text = !InvoiceTrnID
        txtInvoiceCodeID.text = !InvoiceCodeID
        txtInvoicePersonID.text = !InvoicePersonID
        txtInvoiceDateIn.text = !InvoiceDateIn
        txtInvoiceIsActiveID.text = !InvoiceIsActiveID
        txtInvoiceOutDestinationID.text = !InvoiceOutDestinationID
        txtInvoiceOutShipID.text = !InvoiceOutShipID
        txtInvoiceOutPaymentTermID.text = !InvoiceOutPaymentTermID
        txtCodeLastNo.text = !CodeLastNo
        txtCodeLastDate.text = !codeLastDate
        txtCodePersonsPlusOrMinus.text = !CodeCustomers
        chkCodeHandID.Value = !CodeHandID
        txtShipRegistryNo.text = !ShipRegistryNo
        chkPaymentTermCreditID.Value = !PaymentTermCreditID
        
        mskDateIssue.text = format(!InvoiceDateIssue, "dd/mm/yyyy")
        mskDateRefersTo.text = format(!InvoiceDateRefersTo, "dd/mm/yyyy")
        txtCodeShortDescriptionA.text = !CodeShortDescriptionA
        lblCodeDescription.Caption = !CodeDescription
        lblCodeBatch.Caption = !CodeBatch
        lblHand.Caption = IIf(!CodeHandID, "����������", "�������������")
        txtInvoiceNo.text = !InvoiceNo
        txtCustomerDescription.text = !Description
        txtDestinationDescription.text = !DestinationDescription
        txtShipDescription.text = !ShipDescription
        
        chkAgreement.Value = IIf(!InvoiceOutAgreement, 1, 0)
        
        mskAdultsWithTransfer.text = format(!InvoiceOutAdultsWithTransfer, "#,##0")
        mskKidsWithTransfer.text = format(!InvoiceOutKidsWithTransfer, "#,##0")
        mskFreeWithTransfer.text = format(!InvoiceOutFreeWithTransfer, "#,##0")
        mskAdultsAmountWithTransfer.text = format(!InvoiceOutAdultsAmountWithTransfer, "#,##0.00")
        mskKidsAmountWithTransfer.text = format(!InvoiceOutKidsAmountWithTransfer, "#,##0.00")
        
        mskAdultsWithoutTransfer.text = format(!InvoiceOutAdultsWithoutTransfer, "#,##0")
        mskKidsWithoutTransfer.text = format(!InvoiceOutKidsWithoutTransfer, "#,##0")
        mskFreeWithoutTransfer.text = format(!InvoiceOutFreeWithoutTransfer, "#,##0")
        mskAdultsAmountWithoutTransfer.text = format(!InvoiceOutAdultsAmountWithoutTransfer, "#,##0.00")
        mskKidsAmountWithoutTransfer.text = format(!InvoiceOutKidsAmountWithoutTransfer, "#,##0.00")
        
        mskDirectAmount.text = format(!InvoiceOutDirectAmount, "#,##0.00")
        
        txtRemarks.text = !InvoiceOutReason
        txtPaymentTermDescription.text = !PaymentTermDescription
        
        txtInvoiceIsActiveID.text = !InvoiceIsActiveID
        txtInvoiceIsActiveDescription.text = !YesOrNoDescription
        
    End With

End Function

Private Function PrintThisInvoice(blnPreview As Boolean, printOrPDF As String, strInvoiceNo As String, strCustomerName)

    Dim intLoop As Integer
    Dim pdf As New ARExportPDF
    
    If printOrPDF = "PDF" Then
        rptInvoice.Run False
        pdf.AcrobatVersion = 2
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = strReportsPathName & UCase(CommonMain.lblCompany.Caption) & " " & "InvoiceNo" & Right("00000" & strInvoiceNo, 5) & " " & SanitizeString(strCustomerName) & ".pdf"
        pdf.Export rptInvoice.Pages
    Else
        For intLoop = 1 To intInvoiceCopies + 1
            rptInvoice.Restart
            If blnPreview Then
                rptInvoice.Zoom = -2
                rptInvoice.Printer.ColorMode = vbPRCMMonochrome
                rptInvoice.WindowState = vbMaximized
                rptInvoice.Show 1
                Exit For
            Else
                rptInvoice.Printer.DeviceName = strPrinterName
                rptInvoice.PrintReport False
                rptInvoice.Run True
            End If
        Next intLoop
    End If

End Function

Private Function SaveInvoice()

    If blnStatus Then txtInvoiceTrnID.text = AddOneToTheLastRecord("Invoices")
    
    If MainSaveRecord("CommonDB", "Invoices", blnStatus, strApplicationName, "InvoiceID", _
            txtInvoiceID.text, _
            txtInvoiceTrnID.text, _
            txtInvoiceMasterRefersTo.text, _
            txtInvoiceSecondaryRefersTo.text, _
            mskDateIssue.text, _
            txtInvoiceDateIn.text, _
            mskDateRefersTo.text, _
            mskDateIssue.text, _
            txtInvoiceCodeID.text, _
            txtInvoiceNo.text, _
            txtInvoicePersonID.text, _
            txtInvoiceIsActiveID.text, _
            strCurrentUser) <> 0 Then
        IsError = False
    Else
        IsError = True
    End If

End Function

Private Function SaveInvoiceOut()

    If IsError Then Exit Function
    
    If MainSaveRecord("CommonDB", "InvoicesOut", blnStatus, strApplicationName, "InvoiceOutTrnID", txtInvoiceTrnID.text, txtInvoiceTrnID.text, _
        txtInvoiceOutDestinationID.text, _
        txtInvoiceOutShipID.text, _
        chkAgreement.Value, _
        mskAdultsWithTransfer.text, _
        mskKidsWithTransfer.text, _
        mskFreeWithTransfer.text, _
        mskAdultsWithoutTransfer.text, _
        mskKidsWithoutTransfer.text, _
        mskFreeWithoutTransfer.text, _
        mskAdultsAmountWithTransfer.text, _
        mskKidsAmountWithTransfer.text, _
        mskAdultsAmountWithoutTransfer.text, _
        mskKidsAmountWithoutTransfer.text, _
        mskDirectAmount.text, _
        txtVATPercent.text, _
        txtInvoiceOutPaymentTermID.text, _
        txtRemarks.text) <> 0 Then
        IsError = False
    Else
        IsError = True
    End If

End Function

Private Function SaveQRCode()

    SavePicture imgQRCode.Picture, strReportsPathName & "QRCode.bmp"

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    If Not ValidatePersonsAndAmounts Then Exit Function
    
    txtInvoiceOutShipID.text = IIf(txtInvoiceSecondaryRefersTo.text = "2", 0, txtInvoiceOutShipID.text)
    
    BeginTrans
    
    SaveInvoice
    SaveInvoiceOut
    UpdateCodes
    
    If IsError Then
        DisplayErrorMessage True, strStandardMessages(5)
        Rollback
        Exit Function
    Else
        CommitTrans
        blnCancel = True
        AskToPrintInvoice
        ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtInvoiceOutDestinationID, txtInvoiceOutShipID, txtInvoiceOutPaymentTermID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID, txtShipRegistryNo, chkPaymentTermCreditID, txtInvoiceIsActiveID
        ClearFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, lblCodeDescription, lblCodeBatch, lblHand, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
        DisableFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9), cmdIndex(10)
        UpdateButtons Me, 6, 1, 0, 0, 0, IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), 0, 1
    End If
    
End Function

Private Function UpdateCodes()

    If IsError Then Exit Function
    
    Dim rsCodes As Recordset
    
    If blnStatus And chkCodeHandID.Value = 0 Then
        Set rsCodes = CommonDB.OpenRecordset("Codes")
        With rsCodes
            .index = "CodeID"
            .Seek "=", txtInvoiceCodeID.text
            If Not .NoMatch Then
                .Edit
                !CodeLastNo = Int(txtInvoiceNo.text)
                !codeLastDate = mskDateIssue.text
                .Update
            End If
        End With
    End If

End Function

Private Function UpdateInvoiceFieldsWithData(rstRecordset As Recordset)

    On Error GoTo ErrTrap
            
    'Local variables
    Dim strVAT As String
    
    '�����
    Dim intFree As Integer
    Dim intPersons As Integer
    
    '����
    Dim curNet As Currency
    Dim curVAT As Currency
    Dim curGross As Currency
    
    '�����
    intFree = rstRecordset!InvoiceOutFreeWithTransfer + rstRecordset!InvoiceOutFreeWithoutTransfer
    intPersons = rstRecordset!InvoiceOutAdultsWithTransfer + rstRecordset!InvoiceOutAdultsWithoutTransfer + rstRecordset!InvoiceOutKidsWithTransfer + rstRecordset!InvoiceOutKidsWithoutTransfer + rstRecordset!InvoiceOutFreeWithTransfer + rstRecordset!InvoiceOutFreeWithoutTransfer
    
    '����
    strVAT = "1." & rstRecordset!InvoiceOutVATPercent
    curNet = (rstRecordset!InvoiceOutAdultsAmountWithTransfer + rstRecordset!InvoiceOutAdultsAmountWithoutTransfer + rstRecordset!InvoiceOutKidsAmountWithTransfer + rstRecordset!InvoiceOutKidsAmountWithoutTransfer + rstRecordset!InvoiceOutDirectAmount) / Val(strVAT)
    curGross = rstRecordset!InvoiceOutAdultsAmountWithTransfer + rstRecordset!InvoiceOutAdultsAmountWithoutTransfer + rstRecordset!InvoiceOutKidsAmountWithTransfer + rstRecordset!InvoiceOutKidsAmountWithoutTransfer + rstRecordset!InvoiceOutDirectAmount
    curVAT = curGross - curNet
    
    With rptInvoice
    
        .Restart
        
        .lblCompanyData.Caption = arrCompanyData(1) & Chr(13) & arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6)
        
        .lblDate.Caption = rstRecordset!InvoiceDateIssue
        .lblInvoiceDescription.Caption = rstRecordset!CodeDescription
        .lblInvoiceNo.Caption = "�� " & Right("00000" & rstRecordset!InvoiceNo, 5)
        .lblCodeBatch.Caption = rstRecordset!CodeBatch
        
        .lblCustomerDescription.Caption = rstRecordset!Description
        .lblCompanyProfession.Caption = rstRecordset!Profession
        .lblCompanyAddress.Caption = rstRecordset!Address
        .lblCompanyTaxNo.Caption = rstRecordset!TaxNo
        .lblTaxOfficeDescription.Caption = rstRecordset!TaxOfficeDescription
        
        .lblRemarks.Caption = rstRecordset!InvoiceOutReason
        .lblDestination.Caption = rstRecordset!DestinationDescription
        .lblShip.Caption = rstRecordset!ShipDescription
        .lblTripDate.Caption = rstRecordset!InvoiceDateRefersTo
        
        .lblAdultsWithTransfer.Caption = IIf(rstRecordset!InvoiceOutAdultsWithTransfer <> 0, format(rstRecordset!InvoiceOutAdultsWithTransfer, "#,##0"), "")
        .lblAdultAmountWithTransfer.Caption = IIf(rstRecordset!InvoiceOutAdultsAmountWithTransfer <> 0, format(rstRecordset!InvoiceOutAdultsAmountWithTransfer / rstRecordset!InvoiceOutAdultsWithTransfer, "#,##0.00"), "")
        .lblAdultsTotalWithTransfer.Caption = IIf(rstRecordset!InvoiceOutAdultsAmountWithTransfer <> 0, format(rstRecordset!InvoiceOutAdultsAmountWithTransfer, "#,##0.00"), "")
        
        .lblKidsWithTransfer.Caption = IIf(rstRecordset!InvoiceOutKidsWithTransfer <> 0, format(rstRecordset!InvoiceOutKidsWithTransfer, "#,##0"), "")
        .lblKidAmountWithTransfer.Caption = IIf(rstRecordset!InvoiceOutKidsAmountWithTransfer <> 0, format(rstRecordset!InvoiceOutKidsAmountWithTransfer / rstRecordset!InvoiceOutKidsWithTransfer, "#,##0.00"), "")
        .lblKidsTotalWithTransfer.Caption = IIf(rstRecordset!InvoiceOutKidsAmountWithTransfer <> 0, format(rstRecordset!InvoiceOutKidsAmountWithTransfer, "#,##0.00"), "")

        .lblAdultsWithoutTransfer.Caption = IIf(rstRecordset!InvoiceOutAdultsWithoutTransfer <> 0, format(rstRecordset!InvoiceOutAdultsWithoutTransfer, "#,##0"), "")
        .lblAdultAmountWithoutTransfer.Caption = IIf(rstRecordset!InvoiceOutAdultsAmountWithoutTransfer <> 0, format(rstRecordset!InvoiceOutAdultsAmountWithoutTransfer / rstRecordset!InvoiceOutAdultsWithoutTransfer, "#,##0.00"), "")
        .lblAdultsTotalWithoutTransfer.Caption = IIf(rstRecordset!InvoiceOutAdultsAmountWithoutTransfer <> 0, format(rstRecordset!InvoiceOutAdultsAmountWithoutTransfer, "#,##0.00"), "")

        .lblKidsWithoutTransfer.Caption = IIf(rstRecordset!InvoiceOutKidsWithoutTransfer <> 0, format(rstRecordset!InvoiceOutKidsWithoutTransfer, "#,##0"), "")
        .lblKidAmountWithoutTransfer.Caption = IIf(rstRecordset!InvoiceOutKidsAmountWithoutTransfer <> 0, format(rstRecordset!InvoiceOutKidsAmountWithoutTransfer / rstRecordset!InvoiceOutKidsWithoutTransfer, "#,##0.00"), "")
        .lblKidsTotalWithoutTransfer.Caption = IIf(rstRecordset!InvoiceOutKidsAmountWithoutTransfer <> 0, format(rstRecordset!InvoiceOutKidsAmountWithoutTransfer, "#,##0.00"), "")
        
        .lblFree.Caption = IIf(intFree <> 0, format(rstRecordset!InvoiceOutFreeWithTransfer + rstRecordset!InvoiceOutFreeWithoutTransfer, "#,##0"), "")
        
        .lblTotalPersons.Caption = IIf(intPersons <> 0, format(intPersons, "#,##0"), "")
        .lblTotalAmount.Caption = format(curGross, "#,##.00")
        
        .lblVATPercent.Caption = "������ " & rstRecordset!InvoiceOutVATPercent & "%"
        .lblTotalNet1.Caption = format(curNet, "#,##0.00")
        .lblTotalVAT1.Caption = format(curVAT, "#,##0.00")
        .lblTotal1.Caption = format(curGross, "#,##0.00")
        
        .lblGrandTotalNet.Caption = .lblTotalNet1.Caption
        .lblGrandTotalVAT.Caption = .lblTotalVAT1.Caption
        .lblGrandTotal.Caption = format(CCur(.lblGrandTotalNet.Caption) + CCur(.lblGrandTotalVAT.Caption), "#,##0.00")
        
        .lblTotalGross.Caption = .lblTotalNet1.Caption
        .lblTotalVAT.Caption = .lblGrandTotalVAT.Caption
        .lblTotal.Caption = format(CCur(.lblGrandTotalNet.Caption) + CCur(.lblGrandTotalVAT.Caption), "#,##0.00")
        
        .lblPaymentTerm.Caption = rstRecordset!PaymentTermDescription
        
        'QR Code
        If intPrintQRCodeID = 1 Then
            CreateQRCode
            SaveQRCode
            InsertQRCode
        Else
            HideQRCode
        End If

    End With
    
    Exit Function

ErrTrap:
    If Err.Number = 6 Then
        Resume Next
    Else
        DisplayErrorMessage True, Err.Description
    End If

End Function
Private Function ValidateFields()

    ValidateFields = False
    
    '����������
    If Not CheckDate(mskDateIssue.text, strApplicationName) Then
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    '���������� ��������
    If Not CheckDate(mskDateRefersTo.text, strApplicationName) Then
        mskDateRefersTo.SetFocus
        Exit Function
    End If
    
    '���������� �� ����������� ���������� ��� ��� ����������
    If (IsDate(txtCodeLastDate.text) And chkCodeHandID.Value = 0) And blnStatus Then
        If CDate(txtCodeLastDate.text) > CDate(mskDateIssue.text) Then
            If MyMsgBox(4, strApplicationName, strAppMessages(4) & txtCodeLastDate.text & ".", 1) Then
            End If
            mskDateIssue.SetFocus
            Exit Function
        End If
    End If
    
    '���������� �� ���������� ���������� ��� ������
    If CDate(mskDateIssue.text) > Date Then
        If MyMsgBox(4, strApplicationName, strAppMessages(5), 1) Then
        End If
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    '���������� �������� ���������� ��� ������
    If CDate(mskDateRefersTo.text) > Date Then
        If MyMsgBox(4, strApplicationName, strAppMessages(5), 1) Then
        End If
        mskDateRefersTo.SetFocus
        Exit Function
    End If
    
    '����� �����������
    If Len(mskDateIssue.text) <> 10 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    '��������
    If Len(txtInvoiceCodeID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtCodeShortDescriptionA.SetFocus
        Exit Function
    End If
    
    '�� ������������
    If txtInvoiceNo.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    '�� ������������ = �������
    If Not IsNumeric(txtInvoiceNo.text) Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    '������������� ��������. �� ���� ������� ������ �� ����� ��� �� �� ���� ��� ���������� ������������
    If blnStatus Then
        If chkCodeHandID.Value = 0 Then
            If Not CheckForValidYear(txtCodeLastDate.text, mskDateIssue.text) Then
                If MyMsgBox(4, strApplicationName, strStandardMessages(22), 1) Then
                End If
                txtCodeShortDescriptionA.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '������������� �������� ��� ������������: ������ �� �� ������� ��� ��������� ������� ��� ���� �����
    'Temp use only!
    If blnStatus Then
        If chkCodeHandID.Value = 0 Then
            If Not CheckForDuplicateInvoice(mskDateIssue.text, txtInvoiceCodeID.text, txtInvoiceNo.text) Then
                If MyMsgBox(4, strApplicationName, strStandardMessages(22), 1) Then
                End If
                txtCodeShortDescriptionA.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '������������� ��������: ��� ����������� �� ������������� ����������
    'Temp use only!
    'If chkCodeHandID.Value = 0 Then
    If chkCodeHandID.Value = 0 And blnStatus Then
        If CDate(mskDateIssue.text) < CDate(txtCodeLastDate) Then
            If MyMsgBox(4, strApplicationName, strAppMessages(4) & txtCodeLastDate.text, 1) Then
            End If
            mskDateIssue.SetFocus
            Exit Function
        End If
    End If
    
    '�������
    If Len(txtInvoicePersonID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtCustomerDescription.SetFocus
        Exit Function
    End If
    
    '����������
    If Len(txtInvoiceOutDestinationID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtDestinationDescription.SetFocus
        Exit Function
    End If
    
    '�����
    If txtInvoiceOutShipID.text = "0" And txtInvoiceSecondaryRefersTo.text = "1" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtShipDescription.SetFocus
        Exit Function
    End If
    
    '���� ��������
    If Len(txtInvoiceOutPaymentTermID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPaymentTermDescription.SetFocus
        Exit Function
    End If
    
    '�� ��� ���� ��� ����� ��� ��� ����� ��� ��������� ������
    If (Val(mskTotalAmountWithTransfer.text) <> 0 Or Val(mskTotalAmountWithoutTransfer.text)) And Val(mskDirectAmount.text) <> 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskDirectAmount.SetFocus
        Exit Function
    End If
    
    '������ �������
    If Len(txtInvoiceIsActiveID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtInvoiceIsActiveDescription.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function ValidatePersonsAndAmounts()

    ValidatePersonsAndAmounts = False
    
    '������ �������� �� ��������
    If mskAdultsWithTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskAdultsWithTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskAdultsWithTransfer.text) > 9999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskAdultsWithTransfer.SetFocus
        Exit Function
    End If
    
    '���� �������� �� ��������
    If mskAdultsAmountWithTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskAdultsAmountWithTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskAdultsAmountWithTransfer.text) > 99999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskAdultsAmountWithTransfer.SetFocus
        Exit Function
    End If
    
    '���� �������� �� �������� > 0 ��� ������ �������� �� �������� = 0
    If CCur(mskAdultsAmountWithTransfer.text) > 0 And Val(mskAdultsWithTransfer.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskAdultsWithTransfer.SetFocus
        Exit Function
    End If
    
    '������ �������� ����� ��������
    If mskAdultsWithoutTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskAdultsWithoutTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskAdultsWithoutTransfer.text) > 9999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskAdultsWithoutTransfer.SetFocus
        Exit Function
    End If
    
    '���� �������� ����� ��������
    If mskAdultsAmountWithoutTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskAdultsAmountWithoutTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskAdultsAmountWithoutTransfer.text) > 99999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskAdultsAmountWithoutTransfer.SetFocus
        Exit Function
    End If
    
    '���� �������� ����� �������� > 0 ��� ������ �������� ����� �������� = 0
    If CCur(mskAdultsAmountWithoutTransfer.text) > 0 And Val(mskAdultsWithoutTransfer.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskAdultsWithoutTransfer.SetFocus
        Exit Function
    End If
    
    '������ ������� �� ��������
    If mskKidsWithTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskKidsWithTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskKidsWithTransfer.text) > 9999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskKidsWithTransfer.SetFocus
        Exit Function
    End If
    
    '���� ������� �� ��������
    If mskKidsAmountWithTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskKidsAmountWithTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskKidsAmountWithTransfer.text) > 99999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskKidsAmountWithTransfer.SetFocus
        Exit Function
    End If
    
    '���� ������� �� �������� > 0 ��� ������ ������� �� �������� = 0
    If CCur(mskKidsAmountWithTransfer.text) > 0 And Val(mskKidsWithTransfer.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskKidsWithTransfer.SetFocus
        Exit Function
    End If
    
    '������ ������� ����� ��������
    If mskKidsWithoutTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskKidsWithoutTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskKidsWithoutTransfer.text) > 9999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskKidsWithoutTransfer.SetFocus
        Exit Function
    End If
    
    '���� ������� ����� ��������
    If mskKidsAmountWithoutTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskKidsAmountWithoutTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskKidsAmountWithoutTransfer.text) > 99999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskKidsAmountWithoutTransfer.SetFocus
        Exit Function
    End If
    
    '���� ������� ����� �������� > 0 ��� ������ ������� ����� �������� = 0
    If CCur(mskKidsAmountWithoutTransfer.text) > 0 And Val(mskKidsWithoutTransfer.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskKidsWithoutTransfer.SetFocus
        Exit Function
    End If
    
    '������ �� ��������
    If mskFreeWithTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskFreeWithTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskFreeWithTransfer.text) > 9999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskFreeWithTransfer.SetFocus
        Exit Function
    End If
    
    '������ ����� ��������
    If mskFreeWithoutTransfer.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskFreeWithoutTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskFreeWithoutTransfer.text) > 9999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskFreeWithoutTransfer.SetFocus
        Exit Function
    End If
    
    '��������� ������
    If mskDirectAmount.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskDirectAmount.SetFocus
        Exit Function
    End If
    If CCur(mskDirectAmount.text) > 99999 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskDirectAmount.SetFocus
        Exit Function
    End If
    
    ValidatePersonsAndAmounts = True

End Function

Private Sub chkAgreement_Click()

    If chkAgreement.Value = 0 Then
        '�������������. ��� ����� �� ������ ��� ��������
        mskAdultsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithTransfer.text, 2), "#,##0.00")
        mskKidsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithTransfer.text, 3), "#,##0.00")
        mskAdultsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithoutTransfer.text, 4), "#,##0.00")
        mskKidsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithoutTransfer.text, 5), "#,##0.00")
        '��������� �� ������
        AddAmounts
        '������������ �� �����
        DisableFields mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer
    Else
        'If blnStatus Or (chkAgreement.Value = 1 And chkCodeHandID.Value = 1) Then
        '    '���������� �� �����
        '    EnableFields mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer
        '    '��������� �� ������
        '    AddAmounts
        'End If
        'Temp use only!
        '���������� �� �����
        EnableFields mskAdultsWithTransfer, mskAdultsAmountWithTransfer, mskKidsWithTransfer, mskKidsAmountWithTransfer, mskFreeWithTransfer
        EnableFields mskAdultsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsWithoutTransfer, mskKidsAmountWithoutTransfer, mskFreeWithoutTransfer
        '��������� �� ������
        AddAmounts
    End If

End Sub

Private Sub chkAgreement_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

End Sub

Private Sub chkAgreement_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub cmdButton_Click(index As Integer)

    Dim arrDummy()
    
    Select Case index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            ProcessSelectedInvoicesForPrinting txtInvoiceTrnID.text, "Print", arrDummy, arrDummy 'Called when the print button is clicked
        Case 3
            DeleteRecord
        Case 4
            FindRecords
        Case 5
            AbortProcedure False
        Case 6
            AbortProcedure True
    End Select

End Sub

Private Function FindRecords()

    With InvoicesOutIndex
        .Tag = "True"
        .txtInvoiceMasterRefersTo.text = txtInvoiceMasterRefersTo.text
        .txtInvoiceSecondaryRefersTo.text = txtInvoiceSecondaryRefersTo.text
        .Show 1, Me
    End With

End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            '����������� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Codes", "CodeShortDescriptionA, CodeMasterRefersTo", "String, String", txtCodeShortDescriptionA.text, txtInvoiceMasterRefersTo.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 8, 0, 3, 5, 6, 7, 8, 10, 11, "ID", "����. �'", "���������", "�����", "����������", "�������", "��������� ��", "����������", 0, 6, 40, 6, 10, 0, 0, 0, 1, 1, 0, 1, 1, 1, 1, 1)
                txtInvoiceCodeID.text = tmpTableData.strCode
                txtCodeShortDescriptionA.text = tmpTableData.strFirstField
                lblCodeDescription.Caption = tmpTableData.strSecondField
                lblCodeBatch.Caption = IIf(txtInvoiceCodeID.text <> "" And tmpTableData.strThirdField <> "", " ����� " & tmpTableData.strThirdField, "")
                chkCodeHandID.Value = IIf(tmpTableData.strFourthField = "1", 1, 0)
                lblHand.Caption = IIf(tmpTableData.strFourthField = "1", "����������", "�������������")
                txtCodePersonsPlusOrMinus.text = tmpTableData.strFifthField
                txtInvoiceNo.Locked = IIf(chkCodeHandID.Value = 1, False, True)
                txtCodeLastNo.text = tmpTableData.strSixthField
                txtCodeLastDate.text = format(tmpTableData.strSeventhField, "dd/mm/yyyy")
                If txtInvoiceCodeID.text <> "" And chkCodeHandID.Value = 0 Then txtInvoiceNo.text = Val(txtCodeLastNo.text) + 1
            End If
        Case 1
            '����������� - F5
            With TablesCodes
                .Tag = "True"
                .txtCodeMasterRefersTo.text = txtInvoiceMasterRefersTo.text
                .txtCodeSecondaryRefersTo.text = "0"
                .Show 1, Me
            End With
        Case 2
            '������� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtCustomerDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 10, "ID", "��������", "�.�.�.", 0, 40, 15, 1, 0, 1)
                txtInvoicePersonID.text = tmpTableData.strCode
                txtCustomerDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            '������� - F5
            With Persons
                .Tag = "True"
                .Show 1, Me
            End With
        Case 4
            '���������� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription, ShowInList", "String, Numeric", txtDestinationDescription.text, txtInvoiceSecondaryRefersTo.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "���������", 0, 40, 1, 0)
                txtInvoiceOutDestinationID.text = tmpTableData.strCode
                txtDestinationDescription.text = tmpTableData.strFirstField
            End If
        Case 5
            '���������� - F5
            With TablesDestinations
                .Tag = "True"
                .txtShowInList.text = txtInvoiceSecondaryRefersTo.text
                .Show 1, Me
            End With
        Case 6
            '����� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipDescription", "String", txtShipDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 4, "ID", "���������", "��������", 0, 40, 0, 1, 0, 0)
                txtInvoiceOutShipID.text = tmpTableData.strCode
                txtShipDescription.text = tmpTableData.strFirstField
                txtShipRegistryNo.text = tmpTableData.strSecondField
            End If
        Case 7
            '����� - F5
            With TablesShips
                .Tag = "True"
                .Show 1, Me
            End With
        Case 8
            '���� �������� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "PaymentTerms", "PaymentTermDescription", "String", txtPaymentTermDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 2, "ID", "���������", "�������", 0, 40, 0, 1, 0, 0)
                txtInvoiceOutPaymentTermID.text = tmpTableData.strCode
                txtPaymentTermDescription.text = tmpTableData.strFirstField
                chkPaymentTermCreditID.Value = IIf(tmpTableData.strSecondField <> "", 1, 0)
            End If
        Case 9
            '���� �������� - F5
            With TablesPaymentTerms
                .Tag = "True"
                .Show 1, Me
            End With
        Case 10
            '������ ������� - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtInvoiceIsActiveDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 2, "ID", "���������", "�������", 0, 40, 0, 1, 0, 0)
                txtInvoiceIsActiveID.text = tmpTableData.strCode
                txtInvoiceIsActiveDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Public Function SeekRecord(lngTrnID)

    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim rstRecordset As Recordset
    
    strSQL = "SELECT " _
        & "Invoices.InvoiceID, Invoices.InvoiceTrnID, Invoices.InvoiceMasterRefersTo, Invoices.InvoiceSecondaryRefersTo, Invoices.InvoiceDateIssue, Invoices.InvoiceDateIn, Invoices.InvoiceDateRefersTo, Invoices.InvoiceCodeID, Invoices.InvoiceNo, Invoices.InvoicePersonID, Invoices.InvoiceIsActiveID, " _
        & "InvoicesOut.InvoiceOutAgreement, InvoicesOut.InvoiceOutAdultsWithTransfer, InvoicesOut.InvoiceOutKidsWithTransfer, InvoicesOut.InvoiceOutFreeWithTransfer, InvoicesOut.InvoiceOutAdultsWithoutTransfer, InvoicesOut.InvoiceOutKidsWithoutTransfer, InvoicesOut.InvoiceOutFreeWithoutTransfer, InvoicesOut.InvoiceOutAdultsAmountWithTransfer, InvoicesOut.InvoiceOutKidsAmountWithTransfer, InvoicesOut.InvoiceOutAdultsAmountWithoutTransfer, InvoicesOut.InvoiceOutKidsAmountWithoutTransfer, InvoicesOut.InvoiceOutDirectAmount, InvoicesOut.InvoiceOutVATPercent, InvoicesOut.InvoiceOutReason, InvoicesOut.InvoiceOutDestinationID, InvoicesOut.InvoiceOutShipID, InvoicesOut.InvoiceOutPaymentTermID, " _
        & "Codes.CodeShortDescriptionA, Codes.CodeDescription, Codes.CodeBatch, Codes.CodeHandID, Codes.CodeCustomers, Codes.CodeLastNo, Codes.CodeLastDate, " _
        & "Customers.Description, Customers.Profession, Customers.Address, Customers.TaxNo, " _
        & "Ships.ShipDescription, Ships.ShipRegistryNo, " _
        & "PaymentTerms.PaymentTermCreditID, PaymentTerms.PaymentTermDescription, " _
        & "Destinations.DestinationDescription, " _
        & "TaxOffices.TaxOfficeDescription, " _
        & "YesOrNo.YesOrNoDescription " _
        & "FROM ((((((((Invoices " _
        & "INNER JOIN InvoicesOut ON Invoices.InvoiceTrnID = InvoicesOut.InvoiceOutTrnID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID) " _
        & "INNER JOIN Ships ON InvoicesOut.InvoiceOutShipID = Ships.ShipID) " _
        & "INNER JOIN PaymentTerms ON InvoicesOut.InvoiceOutPaymentTermID = PaymentTerms.PaymentTermID) " _
        & "INNER JOIN Destinations ON InvoicesOut.InvoiceOutDestinationID = Destinations.DestinationID) " _
        & "INNER JOIN TaxOffices ON Customers.TaxOfficeID = TaxOffices.TaxOfficeID) " _
        & "INNER JOIN YesOrNo ON Invoices.InvoiceIsActiveID = YesOrNo.YesOrNoID)"
        
    'InvoiceTrnID
    strThisParameter = "lngInvoiceTrnID long"
    strThisQuery = "Invoices.InvoiceTrnID = lngInvoiceTrnID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = lngTrnID

    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    TempQuery.SQL = strSQL & strOrder
    
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    Set SeekRecord = rstRecordset
    
    Exit Function

UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return

End Function

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim ShiftDown, AltDown, CtrlDown
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyP And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyF3 And cmdButton(3).Enabled, vbKeyD And CtrlDown And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyF7 And cmdButton(4).Enabled, vbKeyF And CtrlDown And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyEscape
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
            If cmdButton(6).Enabled Then cmdButton_Click 6
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    PositionControls Me, False
    ColorizeControls Me, False, False
    blnCancel = True
    ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtInvoiceOutDestinationID, txtInvoiceOutShipID, txtInvoiceOutPaymentTermID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID, txtShipRegistryNo, chkPaymentTermCreditID, txtInvoiceIsActiveID
    ClearFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, lblCodeDescription, lblCodeBatch, lblHand, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
    DisableFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskTotalPersonsWithTransfer, mskAdultsAmountWithTransfer, mskKidsAmountWithTransfer, mskTotalAmountWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer, mskTotalPersonsWithoutTransfer, mskAdultsAmountWithoutTransfer, mskKidsAmountWithoutTransfer, mskTotalAmountWithoutTransfer, mskTotalAdults, mskTotalKids, mskTotalFree, mskTotalPersons, mskAdultsAmountTotal, mskKidsAmountTotal, mskTotalPersons, mskTotalAmount, mskDirectAmount, txtRemarks, txtPaymentTermDescription, txtInvoiceIsActiveDescription
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9), cmdIndex(10)
    UpdateButtons Me, 6, 1, 0, 0, 0, IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), 0, 1

End Sub

Private Sub mskAdultsAmountWithoutTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then AddAmounts

End Sub

Private Sub mskAdultsAmountWithTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then AddAmounts

End Sub

Private Sub mskAdultsWithoutTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskAdultsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithoutTransfer.text, 4), "#,##0.00")
        End If
        AddPersons
        AddAmounts
    End If

End Sub

Private Sub mskAdultsWithTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskAdultsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithTransfer.text, 2), "#,##0.00")
        End If
        AddPersons
        AddAmounts
    End If

End Sub

Private Sub mskDateIssue_Validate(Cancel As Boolean)

    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskAdultsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithTransfer.text, 2), "#,##0.00")
            mskKidsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithTransfer.text, 3), "#,##0.00")
            mskAdultsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithoutTransfer.text, 4), "#,##0.00")
            mskKidsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithoutTransfer.text, 5), "#,##0.00")
            AddAmounts
        End If
    End If

End Sub

Private Sub mskDateRefersTo_GotFocus()

    If mskDateIssue.text = "" Then
        ClearFields mskDateRefersTo
    End If

    If mskDateRefersTo.text = "" Then
        mskDateRefersTo.text = mskDateIssue.text
    End If

End Sub


Private Sub mskFreeWithoutTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then
        AddPersons
        AddAmounts
    End If

End Sub

Private Sub mskFreeWithTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then AddPersons

End Sub

Private Sub mskKidsAmountWithoutTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then AddAmounts

End Sub

Private Sub mskKidsAmountWithTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then AddAmounts
    
End Sub

Private Sub mskKidsWithoutTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskKidsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithoutTransfer.text, 5), "#,##0.00")
        End If
        AddPersons
        AddAmounts
    End If

End Sub

Private Sub mskKidsWithTransfer_Validate(Cancel As Boolean)

    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskKidsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithTransfer.text, 3), "#,##0.00")
        End If
        AddPersons
        AddAmounts
    End If

End Sub

Private Sub txtCodeShortDescriptionA_Change()

    If txtCodeShortDescriptionA.text = "" Then
        ClearFields txtInvoiceCodeID, lblCodeDescription, lblCodeBatch, lblHand, txtCodeLastDate, txtCodeLastNo, txtInvoiceNo, chkCodeHandID, txtCodePersonsPlusOrMinus
    End If
    
End Sub

Private Sub txtCodeShortDescriptionA_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtCodeShortDescriptionA_Validate(Cancel As Boolean)

    If txtInvoiceCodeID.text = "" And txtCodeShortDescriptionA.text <> "" Then cmdIndex_Click 0: If txtInvoiceCodeID.text = "" Then Cancel = True

End Sub

Private Sub txtInvoiceIsActiveDescription_Change()

    If txtInvoiceIsActiveDescription.text = "" Then
        ClearFields txtInvoiceIsActiveID
    End If

End Sub

Private Sub txtInvoiceIsActiveDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 10

End Sub


Private Sub txtInvoiceIsActiveDescription_Validate(Cancel As Boolean)

    If txtInvoiceIsActiveID.text = "" And txtInvoiceIsActiveDescription.text <> "" Then cmdIndex_Click 10: If txtInvoiceIsActiveID.text = "" Then Cancel = True
    
End Sub

Private Sub txtPaymentTermDescription_Change()

    If txtPaymentTermDescription.text = "" Then
        ClearFields txtInvoiceOutPaymentTermID, chkPaymentTermCreditID
    End If

End Sub

Private Sub txtPaymentTermDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 8
    If KeyCode = vbKeyF5 Then cmdIndex_Click 9

End Sub

Private Sub txtPaymentTermDescription_Validate(Cancel As Boolean)

    If txtInvoiceOutPaymentTermID.text = "" And txtPaymentTermDescription.text <> "" Then cmdIndex_Click 8: If txtInvoiceOutPaymentTermID.text = "" Then Cancel = True

End Sub

Private Sub txtShipDescription_Change()

    If txtShipDescription.text = "" Then
        txtInvoiceOutShipID.text = "0"
        ClearFields txtShipRegistryNo
    End If
    
End Sub

Private Sub txtShipDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 6
    If KeyCode = vbKeyF5 Then cmdIndex_Click 7

End Sub

Private Sub txtShipDescription_Validate(Cancel As Boolean)

    If txtInvoiceOutShipID.text = "0" And txtShipDescription.text <> "" Then cmdIndex_Click 6: If txtInvoiceOutShipID.text = "0" Then Cancel = True
    
End Sub

Private Sub txtCustomerDescription_Change()

    If txtCustomerDescription.text = "" Then
        ClearFields txtInvoicePersonID
    End If

End Sub

Private Sub txtCustomerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If KeyCode = vbKeyF5 Then cmdIndex_Click 3
    
    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskAdultsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithTransfer.text, 2), "#,##0.00")
            mskKidsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithTransfer.text, 3), "#,##0.00")
            mskAdultsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithoutTransfer.text, 4), "#,##0.00")
            mskKidsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithoutTransfer.text, 5), "#,##0.00")
            AddAmounts
        End If
    End If

End Sub

Private Sub txtCustomerDescription_Validate(Cancel As Boolean)

    If txtInvoicePersonID.text = "" And txtCustomerDescription.text <> "" Then cmdIndex_Click 2: If txtInvoicePersonID.text = "" Then Cancel = True
    
    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskAdultsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithTransfer.text, 2), "#,##0.00")
            mskKidsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithTransfer.text, 3), "#,##0.00")
            mskAdultsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithoutTransfer.text, 4), "#,##0.00")
            mskKidsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithoutTransfer.text, 5), "#,##0.00")
            AddAmounts
        End If
    End If

End Sub

Private Sub txtDestinationDescription_Change()
                                                                
    If txtDestinationDescription.text = "" Then
        ClearFields txtInvoiceOutDestinationID
    End If

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    If KeyCode = vbKeyF5 Then cmdIndex_Click 5
    
    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskAdultsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithTransfer.text, 2), "#,##0.00")
            mskKidsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithTransfer.text, 3), "#,##0.00")
            mskAdultsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithoutTransfer.text, 4), "#,##0.00")
            mskKidsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithoutTransfer.text, 5), "#,##0.00")
            AddAmounts
        End If
    End If
    
End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtInvoiceOutDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 4: If txtInvoiceOutDestinationID.text = "" Then Cancel = True
    
    If Not blnCancel Then
        If chkAgreement.Value = 0 Then
            mskAdultsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithTransfer.text, 2), "#,##0.00")
            mskKidsAmountWithTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithTransfer.text, 3), "#,##0.00")
            mskAdultsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskAdultsWithoutTransfer.text, 4), "#,##0.00")
            mskKidsAmountWithoutTransfer.text = format(CalculateExcursionCharges(txtInvoicePersonID.text, txtInvoiceOutDestinationID.text, mskDateIssue.text, mskKidsWithoutTransfer.text, 5), "#,##0.00")
            AddAmounts
        End If
    End If

End Sub

Private Function EnableOrDisableFields()

    'Temp use only!
    'If chkCodeHandID.Value = 1 Then
        EnableFields mskDateIssue, mskDateRefersTo, txtCodeShortDescriptionA, txtCodeLastNo, txtInvoiceNo, txtCustomerDescription, txtDestinationDescription, txtShipDescription, chkAgreement, _
            mskDirectAmount, _
            txtRemarks, _
            txtPaymentTermDescription, _
            txtInvoiceIsActiveDescription, _
            cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9), cmdIndex(10)
        If chkAgreement.Value = 1 Then EnableFields mskAdultsWithTransfer, mskKidsWithTransfer, mskFreeWithTransfer, mskAdultsWithoutTransfer, mskKidsWithoutTransfer, mskFreeWithoutTransfer
    'End If

End Function

Public Function ProcessSelectedInvoicesForPrinting(strInvoiceTrnID, whatToDo, arrInvoicesTrnID(), arrCustomerNames())

    Dim intLoop As Integer
    Dim rstRecordset As Recordset
    Dim printOrPDF As String
    
    printOrPDF = whatToDo
    
    If printOrPDF = "Print" Then
        If Not SelectPrinter("PrinterPrintsReports") Then Exit Function
        If Not PrinterExists(strPrinterName) Then Exit Function
    End If
    
    If strInvoiceTrnID <> "" Then
        ReDim arrInvoicesTrnID(0)
        arrInvoicesTrnID(0) = strInvoiceTrnID
    End If
    
    For intLoop = 0 To UBound(arrInvoicesTrnID)
        Set rstRecordset = SeekRecord(arrInvoicesTrnID(intLoop))
        If rstRecordset.RecordCount = 0 Then MyMsgBox 4, strApplicationName, strStandardMessages(9), 1: Exit Function
        ClearInvoiceFields
        UpdateInvoiceFieldsWithData rstRecordset
        PrintThisInvoice blnPreviewInvoices, printOrPDF, rstRecordset!InvoiceNo, arrCustomerNames(intLoop) 'False = Do not preview, True = Create PDF instead of print
    Next intLoop
    
    If printOrPDF = "PDF" Then MyMsgBox 1, strApplicationName, strStandardMessages(8), 1

End Function

