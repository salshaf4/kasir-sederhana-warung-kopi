VERSION 5.00
Begin VB.Form FormPerhitungan 
   Caption         =   "WARKOP SI UCOQ"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   52
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   51
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   50
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Frame GHarga 
      Caption         =   "Harga"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   42
      Top             =   6000
      Width           =   5055
      Begin VB.TextBox txtPajak 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   49
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   47
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtSubtotal 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   46
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label LTotal 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label LPajak 
         Caption         =   "Pajak 10%"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Label LSubtotal 
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame GMinuman 
      Caption         =   "Minuman"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   6360
      TabIndex        =   21
      Top             =   1560
      Width           =   5295
      Begin VB.CheckBox CMinuman20 
         Caption         =   "Lemon Tea Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   41
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman19 
         Caption         =   "Nutrisari Susu Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   40
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman18 
         Caption         =   "Nutrisari Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   39
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman17 
         Caption         =   "White Coffee Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   38
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman16 
         Caption         =   "Indocafe Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   37
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman15 
         Caption         =   "Good Day Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   36
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman14 
         Caption         =   "ABC Susu/Mocca"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   35
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman13 
         Caption         =   "Kopi Susu"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   34
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman12 
         Caption         =   "Kopi Hitam"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   33
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman11 
         Caption         =   "Jeruk Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   32
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman10 
         Caption         =   "Soda Susu"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman9 
         Caption         =   "Teh Tarik Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman8 
         Caption         =   "Extra Joss/Kuku Bima"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman7 
         Caption         =   "Susu Coklat/Putih Panas/ Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CheckBox CMinuman6 
         Caption         =   "Ovaltine Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman5 
         Caption         =   "Milo Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman4 
         Caption         =   "Teh Susu Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman3 
         Caption         =   "Teh Manis Panas/Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman2 
         Caption         =   "Teh Tawar Es"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox CMinuman1 
         Caption         =   "Teh Tawar Panas"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame GMakanan 
      Caption         =   "Makanan"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   5295
      Begin VB.CheckBox CMakanan20 
         Caption         =   "Pancong Coklat Kacang Susu"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   20
         Top             =   3720
         Width           =   2535
      End
      Begin VB.CheckBox CMakanan19 
         Caption         =   "Pancong Coklat Susu"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   19
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CheckBox CMakanan18 
         Caption         =   "Pancong Milo"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   18
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan17 
         Caption         =   "Pancong Oreo"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan16 
         Caption         =   "Pancong Keju"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan15 
         Caption         =   "Pancong Tiramisu"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   15
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan14 
         Caption         =   "Pancong Green Tea"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   14
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan13 
         Caption         =   "Pancong Kacang"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan12 
         Caption         =   "Pancong Coklat"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan11 
         Caption         =   "Pancong Polos"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan10 
         Caption         =   "Mie Tektek Ultimate"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan9 
         Caption         =   "Mie Tektek Spesial"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CheckBox CMakanan8 
         Caption         =   "Mie Nyemek Spesial"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox CMakanan7 
         Caption         =   "Mie Dokdok Spesial"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox CMakanan6 
         Caption         =   "Indomie Kuah Sosis"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox CMakanan5 
         Caption         =   "Indomie Kuah Kornet"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox CMakanan4 
         Caption         =   "Indomie Kuah Telor Rebus"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox CMakanan3 
         Caption         =   "Indomie Goreng Sosis"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox CMakanan2 
         Caption         =   "Indomie Goreng Kornet"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox CMakanan1 
         Caption         =   "Indomie Goreng Telor"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARKOP SI UCOQ"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   48
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "FormPerhitungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UpdateTotalAndSubtotal()
Dim subtotal As Long
Dim pajak As Double
Dim total As Double

If CMakanan1.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan2.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan3.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan4.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan5.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan6.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan7.Value = 1 Then
    subtotal = subtotal + 14000
End If

If CMakanan8.Value = 1 Then
    subtotal = subtotal + 10000
End If

If CMakanan9.Value = 1 Then
    subtotal = subtotal + 14000
End If

If CMakanan10.Value = 1 Then
    subtotal = subtotal + 17000
End If

If CMakanan11.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMakanan12.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan13.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan14.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan15.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan16.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan17.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan18.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan19.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan20.Value = 1 Then
    subtotal = subtotal + 9000
End If

If CMinuman1.Value = 1 Then
    subtotal = subtotal + 1000
End If

If CMinuman2.Value = 1 Then
    subtotal = subtotal + 2000
End If

If CMinuman3.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman4.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman5.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman6.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman7.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman8.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman9.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman10.Value = 1 Then
    subtotal = subtotal + 10000
End If

If CMinuman11.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman12.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman13.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman14.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman15.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman16.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman17.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman18.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman19.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman20.Value = 1 Then
    subtotal = subtotal + 5000
End If

txtSubtotal.Text = subtotal

pajak = subtotal * 0.1
txtPajak.Text = pajak

total = subtotal + pajak
' txtTotal.Text = total

End Sub
Private Sub CMakanan1_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMakanan10_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMakanan11_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan12_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan13_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan14_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan15_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan16_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan17_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan18_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan19_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan2_Click()
UpdateTotalAndSubtotal
End Sub



Private Sub CMakanan20_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMakanan3_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan4_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan5_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan6_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan7_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan8_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMakanan9_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub cmdKeluar_Click()
Unload Me
End Sub

Private Sub cmdReset_Click()

Dim i As Integer

For i = 1 To 20
    Controls("CMakanan" & i).Value = 0
Next i

For i = 1 To 20
    Controls("CMinuman" & i).Value = 0
Next i

End Sub

Private Sub cmdTotal_Click()

Dim subtotal As Long
Dim pajak As Double
Dim total As Double

If CMakanan1.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan2.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan3.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan4.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan5.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan6.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan7.Value = 1 Then
    subtotal = subtotal + 14000
End If

If CMakanan8.Value = 1 Then
    subtotal = subtotal + 10000
End If

If CMakanan9.Value = 1 Then
    subtotal = subtotal + 14000
End If

If CMakanan10.Value = 1 Then
    subtotal = subtotal + 17000
End If

If CMakanan11.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMakanan12.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan13.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan14.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan15.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan16.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan17.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan18.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMakanan19.Value = 1 Then
    subtotal = subtotal + 8000
End If

If CMakanan20.Value = 1 Then
    subtotal = subtotal + 9000
End If

If CMinuman1.Value = 1 Then
    subtotal = subtotal + 1000
End If

If CMinuman2.Value = 1 Then
    subtotal = subtotal + 2000
End If

If CMinuman3.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman4.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman5.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman6.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman7.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman8.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman9.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman10.Value = 1 Then
    subtotal = subtotal + 10000
End If

If CMinuman11.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman12.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman13.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman14.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman15.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman16.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman17.Value = 1 Then
    subtotal = subtotal + 5000
End If

If CMinuman18.Value = 1 Then
    subtotal = subtotal + 4000
End If

If CMinuman19.Value = 1 Then
    subtotal = subtotal + 7000
End If

If CMinuman20.Value = 1 Then
    subtotal = subtotal + 5000
End If

pajak = subtotal * 0.1

total = subtotal + pajak
txtTotal.Text = total

End Sub



Private Sub CMinuman1_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMinuman10_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman11_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman12_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman13_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman14_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman15_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman16_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman17_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman18_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman19_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman2_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMinuman20_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman3_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMinuman4_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMinuman5_Click()
UpdateTotalAndSubtotal
End Sub

Private Sub CMinuman6_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMinuman7_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMinuman8_Click()
UpdateTotalAndSubtotal
End Sub


Private Sub CMinuman9_Click()
UpdateTotalAndSubtotal
End Sub


