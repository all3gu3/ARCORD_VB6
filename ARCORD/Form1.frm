VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
   ScaleHeight     =   6075
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular área de un terreno"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calcular perímetro de un terreno"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular el largo de un recorrido"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Calcular distancia entre dos puntos"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AREACORD"
      BeginProperty Font 
         Name            =   "Nasalization Rg"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form1.Hide
End Sub

