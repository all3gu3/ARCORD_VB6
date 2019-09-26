VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   6150
   ClientTop       =   2565
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   9015
   Begin VB.CommandButton Command7 
      Caption         =   "Cambiar a coordenadas decimales"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   120
      TabIndex        =   53
      Top             =   3840
      Width           =   2385
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4560
      TabIndex        =   52
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4560
      TabIndex        =   51
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4560
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4560
      TabIndex        =   49
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cambiar a coordenadas hexagesimales"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   120
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CheckBox Check8 
      Caption         =   "W"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   47
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check7 
      Caption         =   "E"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   46
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check6 
      Caption         =   "S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   45
      Top             =   2400
      Width           =   615
   End
   Begin VB.CheckBox Check5 
      Caption         =   "N"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   44
      Top             =   2400
      Width           =   615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "W"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   43
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "E"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   42
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   41
      Top             =   840
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "N"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   40
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4680
      TabIndex        =   39
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   38
      Top             =   6120
      Width           =   2400
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      TabIndex        =   29
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5400
      TabIndex        =   28
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4680
      TabIndex        =   27
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   26
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      TabIndex        =   25
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5400
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4680
      TabIndex        =   23
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   22
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   2400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calcular distancia"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Introducir valores de la segunda coordenada"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   2145
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "Introducir valores de la primera coordenada"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   2145
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      TabIndex        =   12
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5400
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4680
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   9
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5400
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Cálculo de distancia entre dos puntos"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   37
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      TabIndex        =   36
      Top             =   3960
      Width           =   5400
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   """"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6480
      TabIndex        =   35
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   """"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6480
      TabIndex        =   34
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5160
      TabIndex        =   33
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5160
      TabIndex        =   32
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   31
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   30
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "LONGITUD"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   21
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   " LATITUD"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   20
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   """"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6480
      TabIndex        =   15
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5160
      TabIndex        =   14
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   13
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   """"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6480
      TabIndex        =   8
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5160
      TabIndex        =   7
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   6
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   960
      Width           =   15
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "LONGITUD"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " LATITUD"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   225
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L As String 'L es un valor para latitud: norte o sur
Dim Lo As String 'Lo es un valor para longitud: este u oeste
Dim L2 As String 'L2 es un valor para latitud2: norte o sur
Dim Lo2 As String 'Lo2 es un valor para longitud2: este u oeste
Dim Latit1 As Double
Dim Latit2 As Double
Dim Longit1 As Double
Dim Longit2 As Double
Dim LatMa As Double
Dim LatMe As Double
Dim LongMa As Double
Dim LongMe As Double
Dim myAnswer As Integer
Private Function DeshabTodo() 'DESHABILITAR TODO
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False

Command1.Enabled = False
Command2.Enabled = False
Command6.Enabled = False
Command7.Enabled = False

End Function

Private Function HabCord1Hex() 'HABILITAR COORDENADA UNO HEXAGESIMAL
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True   'HABILITA TEXTBOXES Y CHECBOXES
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Text1.SetFocus


Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False
                          'DESHABILITA TEXTBOXES Y CHECBOXES
Text17.Enabled = False
Text18.Enabled = False 'decimales
Text19.Enabled = False
Text20.Enabled = False

End Function
Private Function HabCord2Hex() 'HABILITAR COORDENADA DOS HEXAGESIMAL
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True   'HABILITA TEXTBOXES Y CHECBOXES
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
Check8.Enabled = True
Text9.SetFocus


Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
                         'DESHABILITA TEXTBOXES Y CHECBOXES
Text17.Enabled = False
Text18.Enabled = False 'decimales
Text19.Enabled = False
Text20.Enabled = False
                         
                         

End Function
Private Function HabCord1Dec() 'HABILITAR COORDENADA UNO DECIMAL
Text1.Enabled = True
Text5.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Check1.Enabled = True    'HABILITA TEXTBOXES Y CHECBOXES
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Text1.SetFocus


Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text19.Enabled = False
Text20.Enabled = False   'DESHABILITA TEXTBOXES Y CHECBOXES
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False


End Function
Private Function HabCord2Dec() 'HABILITAR COORDENADA DOS DECIMAL
Text9.Enabled = True
Text13.Enabled = True
Text19.Enabled = True
Text20.Enabled = True
Check5.Enabled = True    'HABILITA TEXTBOXES Y CHECBOXES
Check6.Enabled = True
Check7.Enabled = True
Check8.Enabled = True



Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text2.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False   'DESHABILITA TEXTBOXES Y CHECBOXES
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
End Function
Private Function ValErr()
If (Text1 <= 0) Or (Text9 <= 0) Then 'LATITUD NEGATIVA
    MsgBox "No es posible realizar los cálculos si usted introdice coordenadas negativas. En caso de que sus coordenadas sean negativas, probablemente su latitud es Sur. Porfavor, seleccione el cuadro con -S- e inserte las coordenadas con valores positivos", vbExclamation
  
  ElseIf (Text5 > 90) Or (Text13 > 90) Then 'LATITUD MAYOR A 90
    MsgBox "No es posible realizar los cálculos porque esas coordenadas no existen. La latitud no puede ser mayor a 90°", vbExclamation
  
  ElseIf (Text1 <= 0) Or (Text9 <= 0) Then 'LONGITUD NEGATIVA
    MsgBox "No es posible realizar los cálculos si usted introdice coordenadas negativas. En caso de que sus coordenadas sean negativas, probablemente su longitud es Oeste. Porfavor, seleccione el cuadro con -W- e inserte las coordenadas con valores positivos", vbExclamation
  
  ElseIf (Text5 > 180) Or (Text13 > 180) Then 'LONGITUD MAYOR A 180
    MsgBox "No es posible realizar los cálculos porque esas coordenadas no existen. La longitud no puede ser mayor a 180°", vbExclamation
  
  ElseIf (Text2 > 60) Or (Text3 > 60) Or (Text4 > 60) Or (Text6 > 60) Or (Text7 > 60) Or (Text8 > 60) Or (Text10 > 60) Or (Text11 > 60) Or (Text12 > 60) Or (Text14 > 60) Or (Text15 > 60) Or (Text16 > 60) Then
    MsgBox "Esas coordenadas no son válidas. Los minutos (') y segundos ('') de sus coordenadas deben ser valores de 0 a 60", vbExclamation
  'MINUTOS Y SEGUNDOS INVÁLIDOS (MAYORES A 60)
  
  ElseIf (Text2 < 0) Or (Text3 < 0) Or (Text4 < 0) Or (Text6 < 0) Or (Text7 < 0) Or (Text8 < 0) Or (Text10 < 0) Or (Text11 < 0) Or (Text12 < 0) Or (Text14 < 0) Or (Text15 <> 0) Or (Text16 < 0) Then
    MsgBox "Esas coordenadas no son válidas. Los minutos (') y segundos ('') de sus coordenadas deben ser mayores o iguales a 0", vbExclamation
  'MINUTOS Y SEGUNDOS INVÁLIDOS (VALORES NEGATIVOS)
  
  ElseIf (Text17 < 0) Or (Text18 < 0) Or (Text19 < 0) Or (Text20 < 0) Then
    MsgBox "Esas coordenadas no son válidas. Todas deben ser positivas"
  
  Else
     MsgBox "Las coordenadas deben ser números; porfavor, no inserte otras cosas.", vbExclamation
End If
End Function

Private Function CalcularDistDec()
'OBTENCIÓN DE VALORES
Latit1 = Val((Text1) & "." & (Text17))
Longit1 = Val((Text5) & "." & (Text18))

Latit2 = Val((Text9) & "." & (Text19))
Longit2 = Val((Text13) & "." & (Text20))

'COMPARACIÓN DE VALORES
If Latit1 > Latit2 Then
 LatMa = Latit1
 LatMe = Latit2
Else
 LatMa = Latit2
 LatMe = Latit1
End If

If Longit1 > Longit2 Then
 LongMa = Longit1
 LongMe = Longit2
Else
 LongMa = Longit2
 LongMe = Longit1
End If

PromLat = ((111 * (Cos(LatMa))) + (111 * (Cos(LatMe)))) / 2
'PromLat es el promedio de la medida de un grado en la posición de las coordenadas

FracLat = LatMa - LatMe
FracLong = LongMa - LongMe

CaV = (FracLat * PromLat)
CaH = (FracLong * 111)

Dist = (Sqr((CaV ^ 2) + (CaH ^ 2))) * 1000

Label18.Caption = "La distancia entre los puntos es de: " & Dist & " Metros"

End Function

Private Function CalcularDist()
'CONVERSIÓN A DECIMALES
Latit1 = Val(Val(Text1) + (Val(Text2 / 60) + Val(Val((Text3) & "." & (Text4)) / 3600)))
Longit1 = Val(Val(Text5) + (Val(Text6 / 60) + Val(Val((Text7) & "." & (Text8)) / 3600)))

Latit2 = Val(Val(Text9) + (Val(Text10 / 60) + Val(Val((Text11) & "." & (Text12)) / 3600)))
Longit2 = Val(Val(Text13) + (Val(Text14 / 60) + Val(Val((Text15) & "." & (Text16)) / 3600)))

'COMPARACIÓN DE VALORES
If Latit1 > Latit2 Then
 LatMa = Latit1
 LatMe = Latit2
Else
 LatMa = Latit2
 LatMe = Latit1
End If

If Longit1 > Longit2 Then
 LongMa = Longit1
 LongMe = Longit2
Else
 LongMa = Longit2
 LongMe = Longit1
End If

PromLat = ((111 * (Cos(LatMa))) + (111 * (Cos(LatMe)))) / 2
'PromLat es el promedio de la medida de un grado en la posición de las coordenadas

FracLat = LatMa - LatMe
FracLong = LongMa - LongMe

CaV = (FracLat * PromLat)
CaH = (FracLong * 111)

Dist = (Sqr((CaV ^ 2) + (CaH ^ 2))) * 1000

Label18.Caption = "La distancia entre los puntos es de: " & Dist & " Metros"




'MUESTRA LAS COORDENADAS ORGANIZADAS Label18.Caption = "1° " & LatMa & " 1°" & LongMa & "2° " & LatMe & " 2°" & LongMe
'MUESTRA LAS CUATRO COORDENADAS Label18.Caption = "Latitud 1 = " & Latit1 & " Latitud 2 = " & Latit2 & "  Longitud 1 = " & Longit1 & " Longitud 2 = " & Longit2
'FALLIDO1 Val((Val(Text1.text) & "." & (Val(Text2.Text) / 60)) + (Val(Val(Text3.text) & "." & Val(Text4.text)) / 3600)))
'FALLIDO2 Val(Val(Text1 / 10) & (Val(Text2 / 60) + Val(Val((Text3) & "." & (Text4)) / 3600)))
'HEXAGESIMAL (Text1) & "° " & (Text2) & "'' " & (Text3) & "." & (Text4) & "'"
End Function

Private Sub Check1_Click()
If Check1.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check2.Value = 0
L = " °N"
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check1.Value = 0
L = " °S"
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check4.Value = 0
Lo = " °E"
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check3.Value = 0
Lo = " °W"
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check6.Value = 0
L2 = " °N"
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check5.Value = 0
L2 = " °S"
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check8.Value = 0
Lo2 = " °E"
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then 'No escoger dos opciones al mismo tiempo
Check7.Value = 0
Lo2 = " °W"
End If
End Sub

Private Sub Command1_Click()
If Command7.Visible = True Then 'ESTÁ EN MODO HEX
HabCord1Hex
ElseIf Command6.Visible = True Then 'ESTÁ EN MODO DEC
HabCord1Dec
End If
Command1.Enabled = False 'Anulacion mutua coord1 y coord2
Command2.Enabled = True
Command4.Enabled = True 'Nuevo disponible
End Sub

Private Sub Command2_Click()
If Command7.Visible = True Then 'ESTÁ EN MODO HEX
HabCord2Hex
ElseIf Command6.Visible = True Then 'ESTÁ EN MODO DEC
HabCord2Dec
End If
Command2.Enabled = False 'Anulacion mutua coord1 y coord2
Command1.Enabled = True
Command4.Enabled = True 'Nuevo disponible
End Sub

Private Sub Command3_Click()
If Command7.Visible = True Then 'Hexagesimal está habilitado

 If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Then
    MsgBox "Debes rellenar todos los recuadros. Si el valor que quieres insertar es 0, entonces pon un 0", vbExclamation
      'EL USUARIO OLVIDÓ RELLENAR UN RECUADRO
    
 ElseIf (Val(Text1) >= 0) And (Text1 < 181) And (Text5 >= 0) And (Text5 < 91) And (Text9 >= 0) And (Text9 < 181) And (Text13 >= 0) And (Text13 < 91) And (Text2 < 61) And (Text2 >= 0) And (Text3 < 61) And (Text3 >= 0) And (Text4 < 61) And (Text4 >= 0) And (Text6 < 61) And (Text6 >= 0) And (Text7 < 61) And (Text7 >= 0) And (Text8 < 61) And (Text8 >= 0) And (Text10 < 61) And (Text10 >= 0) And (Text11 < 61) And (Text11 >= 0) And (Text12 < 61) And (Text12 >= 0) And (Text14 < 61) And (Text14 >= 0) And (Text15 < 61) And (Text15 >= 0) And (Text16 < 61) And (Text16 >= 0) Then 'si es mayor o igual a cero y menor a 181
      'LAS COORDENADAS SON VÁLIDAS Y EXISTEN
    
    CalcularDist
    MsgBox "Cálculos realizados correctamente"
    DeshabTodo
    Command3.Enabled = False
    Command4.Enabled = True
    
 Else
    ValErr
      'ALGO ESTÁ MAL
       
 End If

ElseIf Command7.Visible = False Then 'Decimal está habilitado
 
 If Text1.Text = "" Or Text5.Text = "" Or Text9.Text = "" Or Text13.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Then
    MsgBox "Debes rellenar todos los recuadros. Si el valor que quieres insertar es 0, entonces pon un 0", vbExclamation
      'EL USUARIO OLVIDÓ RELLENAR UN RECUADRO
    
 ElseIf (Text1 >= 0) And (Text1 < 181) And (Text5 >= 0) And (Text5 < 91) And (Text9 >= 0) And (Text9 < 181) Then
      'LAS COORDENADAS SON VÁLIDAS Y EXISTEN
    
    CalcularDistDec
    MsgBox "Cálculos realizados correctamente"
    DeshabTodo
    Command3.Enabled = False
    Command4.Enabled = True
    
 Else
    ValErr
      'ALGO ESTÁ MAL
 
 End If

End If

End Sub

Private Sub Command4_Click()
Command6.Enabled = True
Command7.Enabled = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = "" '6 boton de cambiar a hexa; 7 a deci
Label18.Caption = ""


If Command7.Visible = True Then 'ESTÁ EN MODO HEX
HabCord1Hex
ElseIf Command6.Visible = True Then 'ESTÁ EN MODO DEC
HabCord1Dec
End If
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
Form1.Show
Form2.Hide
'MsgBox "Cálculos realizados correctamente", vbExclamation ' Displays with an OK button with the Exclimation icon
'myAnswer = MsgBox("Are you sure?", vbOKCancel + vbQuestion) ' Displays with an OK and Cancel button with the Questionmark icon
'myAnswer = MsgBox("Do you want to dance?", vbYesNo + vbQuestion + vbDefaultButton2) ' Displays with an Yes and No button with the Questionmark icon, and the 2nd button (the No button) is default.
End Sub

Private Sub Command6_Click() 'EL BOTON FUE PRESIONADO PARA CAMBIAR A MODO DEC
Command1.Enabled = True

Label4.Visible = True
Label7.Visible = True
Label12.Visible = True
Label13.Visible = True

Text17.Visible = False
Text18.Visible = False
Text19.Visible = False 'ESCONDE MINUTOS Y SEGUNDOS
Text20.Visible = False
                           
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text20.Enabled = False

Text1.Enabled = False
Text5.Enabled = False
Text9.Enabled = False
Text13.Enabled = False


Command7.Visible = True 'ANULACION MUTUA (VISIBILIDAD) HEX-DEC
Command6.Visible = False

End Sub
Private Sub Command7_Click() 'EL BOTON FUE PRESIONADO PARA CAMBIAR A MODO HEX
Command1.Enabled = True

Label4.Visible = False
Label7.Visible = False
Label12.Visible = False
Label13.Visible = False

Text17.Visible = True
Text18.Visible = True
Text19.Visible = True
Text20.Visible = True

Text17.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
Text20.Enabled = True

Text1.Enabled = False
Text5.Enabled = False
Text9.Enabled = False
Text13.Enabled = False


Command6.Visible = True 'ANULACION MUTUA (VISIBILIDAD) HEX-DEC
Command7.Visible = False



End Sub

