VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CAT_Paquetes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paquetes"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14730
   Icon            =   "CAT_Paquetes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "CAT_Paquetes.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   5655
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   9975
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"CAT_Paquetes.frx":05A6
      End
   End
End
Attribute VB_Name = "CAT_Paquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
