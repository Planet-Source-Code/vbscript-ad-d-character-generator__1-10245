VERSION 5.00
Begin VB.Form frmViewChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Character"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmViewChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   4080
      TabIndex        =   67
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txtRecord 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   65
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   ">>|"
      Height          =   255
      Left            =   3240
      TabIndex        =   64
      Top             =   7560
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   255
      Left            =   2640
      TabIndex        =   63
      Top             =   7560
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   255
      Left            =   840
      TabIndex        =   62
      Top             =   7560
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "|<<"
      Height          =   255
      Left            =   240
      TabIndex        =   61
      Top             =   7560
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   66
      Top             =   7320
      Width           =   3735
   End
   Begin VB.TextBox txtNotes 
      BackColor       =   &H80000000&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6480
      Width           =   5295
   End
   Begin VB.Frame FraMagic 
      Caption         =   "Magic Items"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
      Begin VB.TextBox txtMagic5 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtMagic4 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtMagic3 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtMagic2 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtMagic1 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraArmor 
      Caption         =   "Armor/Weapons"
      Height          =   2415
      Left            =   2760
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
      Begin VB.TextBox txtWeapon2 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtWeapon1 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtShield 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtHelm 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtArmor 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblWeapon2 
         Alignment       =   2  'Center
         Caption         =   "Weapon2"
         Height          =   255
         Left            =   1320
         TabIndex        =   44
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblWeapon1 
         Alignment       =   2  'Center
         Caption         =   "Weapon1"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblShield 
         Alignment       =   2  'Center
         Caption         =   "Shield Carried"
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblHelm 
         Alignment       =   2  'Center
         Caption         =   "Helm Worn"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblArmor 
         Alignment       =   2  'Center
         Caption         =   "Armor Worn"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame fraSaving 
      Caption         =   "Saving Throws"
      Height          =   2055
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
      Begin VB.TextBox txtPara 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtPetr 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtSpell 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtRod 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtBreath 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblPara 
         Caption         =   "Paralyzation"
         Height          =   255
         Left            =   840
         TabIndex        =   55
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblPetr 
         Caption         =   "Petrification"
         Height          =   255
         Left            =   840
         TabIndex        =   54
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblSpell 
         Caption         =   "Spells/Rings"
         Height          =   255
         Left            =   840
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblRod 
         Caption         =   "Rods/Staves et. al"
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblBreath 
         Caption         =   "Breath Weapon"
         Height          =   255
         Left            =   840
         TabIndex        =   51
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "Statistics"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
      Begin VB.TextBox txtCha 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtCon 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtStr 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtWis 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtDex 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblCha 
         Caption         =   "Charisma"
         Height          =   255
         Left            =   840
         TabIndex        =   50
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblCon 
         Caption         =   "Constitution"
         Height          =   255
         Left            =   840
         TabIndex        =   49
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblWis 
         Caption         =   "Wisdom"
         Height          =   255
         Left            =   840
         TabIndex        =   48
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblDex 
         Caption         =   "Dexterity"
         Height          =   255
         Left            =   840
         TabIndex        =   47
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblInt 
         Caption         =   "Inelligence"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblStr 
         Caption         =   "Strength (Power)"
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraPersonal 
      Caption         =   "Personal Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtArmorClass 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHitPoint 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtAlign 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRace 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtClass 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCharName 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblArmorClass 
         Alignment       =   2  'Center
         Caption         =   "Armor Class"
         Height          =   255
         Left            =   4200
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblHitPoint 
         Alignment       =   2  'Center
         Caption         =   "Hip Points"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         Caption         =   "Level"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblAlign 
         Alignment       =   2  'Center
         Caption         =   "Alignment"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblAge 
         Alignment       =   2  'Center
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         Caption         =   "Player Name"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Caption         =   "Class"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRace 
         Alignment       =   2  'Center
         Caption         =   "Race"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         Caption         =   "Character "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmViewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmViewChar
'| Purpose: Allows user to view character sheets for each
'|          character in the database
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database          'declares database variable
    Dim rsChar As Recordset     'declares recordset variable

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'opens database
    Set rsChar = db.OpenRecordset("tblCha")         'populates recordset
    'moves to first record and displays it.
    rsChar.MoveFirst
    GetData
End Sub

Private Sub cmdClose_Click()     'closes recordset, database and form
    rsChar.Close
    db.Close
    frmViewChar.Hide
End Sub

Private Sub cmdStart_Click()    'Moves to and displays first record
    On Error Resume Next
    rsChar.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click()     'Moves to and displays previous record
    On Error Resume Next
    rsChar.MovePrevious
    GetData
End Sub

Private Sub cmdEnd_Click()      'Moves to and displays last record
    On Error Resume Next
    rsChar.MoveLast
    GetData
End Sub

Private Sub cmdNext_Click()     'Moves to and displays next record
    On Error Resume Next
    rsChar.MoveNext
    GetData
End Sub

Function GetData()
    'Populates the form with appropriate data.
    txtRecord.Text = "Character " & rsChar.Fields("ID")
    txtCharName.Text = rsChar.Fields("CharName")
    txtPlayer.Text = rsChar.Fields("Player")
    txtClass.Text = rsChar.Fields("Class")
    txtRace.Text = rsChar.Fields("Race")
    txtAlign.Text = rsChar.Fields("Alignment")
    txtLevel.Text = rsChar.Fields("Level")
    txtAge.Text = rsChar.Fields("Age")
    txtHitPoint.Text = rsChar.Fields("HitPoint")
    txtArmorClass.Text = rsChar.Fields("ArmorClass")
    txtStr.Text = rsChar.Fields("Str")
    txtInt.Text = rsChar.Fields("Int")
    txtDex.Text = rsChar.Fields("Dex")
    txtWis.Text = rsChar.Fields("Wis")
    txtCon.Text = rsChar.Fields("Con")
    txtCha.Text = rsChar.Fields("Cha")
    txtBreath.Text = rsChar.Fields("Breath")
    txtRod.Text = rsChar.Fields("Rod")
    txtSpell.Text = rsChar.Fields("Spell")
    txtPetr.Text = rsChar.Fields("Petr")
    txtPara.Text = rsChar.Fields("Para")
    txtArmor.Text = rsChar.Fields("Armor")
    txtHelm.Text = rsChar.Fields("Helm")
    txtShield.Text = rsChar.Fields("Shield")
    txtWeapon1.Text = rsChar.Fields("Weapon1")
    txtWeapon2.Text = rsChar.Fields("Weapon2")
    txtMagic1.Text = rsChar.Fields("Magic1")
    txtMagic2.Text = rsChar.Fields("Magic2")
    txtMagic3.Text = rsChar.Fields("Magic3")
    txtMagic4.Text = rsChar.Fields("Magic4")
    txtMagic5.Text = rsChar.Fields("Magic5")
    txtNotes.Text = rsChar.Fields("Notes")
End Function
