VERSION 5.00
Begin VB.Form frmEditChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Character"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmEditChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   720
      TabIndex        =   67
      Top             =   7320
      Width           =   4095
      Begin VB.CommandButton cmdEnd 
         Caption         =   ">>|"
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRecord 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<"
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "|<<"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Opperation"
      Height          =   375
      Left            =   2880
      TabIndex        =   36
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Character"
      Height          =   375
      Left            =   960
      TabIndex        =   35
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox txtNotes 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   6480
      Width           =   5295
   End
   Begin VB.Frame FraMagic 
      Caption         =   "Magic Items"
      Height          =   2055
      Left            =   120
      TabIndex        =   41
      Top             =   4320
      Width           =   2535
      Begin VB.TextBox txtMagic5 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtMagic4 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtMagic3 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtMagic2 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtMagic1 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraArmor 
      Caption         =   "Armor/Weapons"
      Height          =   2415
      Left            =   2760
      TabIndex        =   40
      Top             =   3960
      Width           =   2655
      Begin VB.TextBox txtWeapon2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtWeapon1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtShield 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtHelm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtArmor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblWeapon2 
         Alignment       =   2  'Center
         Caption         =   "Weapon2"
         Height          =   255
         Left            =   1320
         TabIndex        =   55
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblWeapon1 
         Alignment       =   2  'Center
         Caption         =   "Weapon1"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblShield 
         Alignment       =   2  'Center
         Caption         =   "Shield Carried"
         Height          =   255
         Left            =   1320
         TabIndex        =   53
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblHelm 
         Alignment       =   2  'Center
         Caption         =   "Helm Worn"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblArmor 
         Alignment       =   2  'Center
         Caption         =   "Armor Worn"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame fraSaving 
      Caption         =   "Saving Throws"
      Height          =   2055
      Left            =   2760
      TabIndex        =   39
      Top             =   1800
      Width           =   2655
      Begin VB.TextBox txtPara 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtPetr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtSpell 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtRod 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtBreath 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblPara 
         Caption         =   "Paralyzation"
         Height          =   255
         Left            =   840
         TabIndex        =   66
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblPetr 
         Caption         =   "Petrification"
         Height          =   255
         Left            =   840
         TabIndex        =   65
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblSpell 
         Caption         =   "Spells/Rings"
         Height          =   255
         Left            =   840
         TabIndex        =   64
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblRod 
         Caption         =   "Rods/Staves et. al"
         Height          =   255
         Left            =   840
         TabIndex        =   63
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblBreath 
         Caption         =   "Breath Weapon"
         Height          =   255
         Left            =   840
         TabIndex        =   62
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "Statistics"
      Height          =   2415
      Left            =   120
      TabIndex        =   38
      Top             =   1800
      Width           =   2535
      Begin VB.TextBox txtCha 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtCon 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtStr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtWis 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtDex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblCha 
         Caption         =   "Charisma"
         Height          =   255
         Left            =   840
         TabIndex        =   61
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblCon 
         Caption         =   "Constitution"
         Height          =   255
         Left            =   840
         TabIndex        =   60
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblWis 
         Caption         =   "Wisdom"
         Height          =   255
         Left            =   840
         TabIndex        =   59
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblDex 
         Caption         =   "Dexterity"
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblInt 
         Caption         =   "Inelligence"
         Height          =   255
         Left            =   840
         TabIndex        =   57
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblStr 
         Caption         =   "Strength (Power)"
         Height          =   255
         Left            =   840
         TabIndex        =   56
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraPersonal 
      Caption         =   "Personal Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   37
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox cmbAlign 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbRace 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtArmorClass 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4440
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHitPoint 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCharName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblArmorClass 
         Alignment       =   2  'Center
         Caption         =   "Armor Class"
         Height          =   255
         Left            =   4320
         TabIndex        =   50
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblHitPoint 
         Alignment       =   2  'Center
         Caption         =   "Hip Points"
         Height          =   255
         Left            =   3360
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         Caption         =   "Level"
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblAlign 
         Alignment       =   2  'Center
         Caption         =   "Alignment"
         Height          =   255
         Left            =   840
         TabIndex        =   47
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblAge 
         Alignment       =   2  'Center
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         Caption         =   "Player Name"
         Height          =   255
         Left            =   4080
         TabIndex        =   45
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Caption         =   "Class"
         Height          =   255
         Left            =   2760
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRace 
         Alignment       =   2  'Center
         Caption         =   "Race"
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         Caption         =   "Character "
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEditChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmEditChar
'| Purpose: Allows user to update and edit the character
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database          'sets up database variable
    Dim rsChar As Recordset     'declares Character recordset
    Dim rsRace As Recordset     'declares Race recordset
    Dim rsClass As Recordset    'declares Class recordset
    Dim rsAlign As Recordset    'declares Alignment recordset

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'Opens database
    Set rsChar = db.OpenRecordset("tblCha")         'populates recordsets
    Set rsRace = db.OpenRecordset("tblRace")
    Set rsClass = db.OpenRecordset("tblClass")
    Set rsAlign = db.OpenRecordset("tblAlign")
    
    'Populates the Race drop-down combo-box
    rsRace.MoveFirst
    Do Until rsRace.EOF
        cmbRace.AddItem (rsRace.Fields("Race"))
        rsRace.MoveNext
    Loop
    
    'Populates the Class drop-down combo-box
    rsClass.MoveFirst
    Do Until rsClass.EOF
        cmbClass.AddItem (rsClass.Fields("Class"))
        rsClass.MoveNext
    Loop
    
    'Populates the Alignment drop-down combo-box
    rsAlign.MoveFirst
    Do Until rsAlign.EOF
        cmbAlign.AddItem (rsAlign.Fields("Alignment"))
        rsAlign.MoveNext
    Loop
    
    rsChar.MoveFirst    'Goes to first record
    GetData             'populates the form.
End Sub

Private Sub cmdCancel_Click()     'closes recordset, database and form
    rsChar.Close
    rsRace.Close
    rsAlign.Close
    rsClass.Close
    db.Close
    frmEditChar.Hide
End Sub

Private Sub cmdStart_Click()    'Moves to and displays the first record
    On Error Resume Next
    rsChar.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click()     'Moves to and displays the next record
    On Error Resume Next
    rsChar.MovePrevious
    GetData
End Sub

Private Sub cmdEnd_Click()      'Moves to and displays the last record
    On Error Resume Next
    rsChar.MoveLast
    GetData
End Sub

Private Sub cmdNext_Click()     'Moves to and displays the next record
    On Error Resume Next
    rsChar.MoveNext
    GetData
End Sub

Function GetData()
    'Reads data from database and populates the form.
    txtRecord.Text = "Character " & rsChar.Fields("ID")
    txtCharName.Text = rsChar.Fields("CharName")
    txtPlayer.Text = rsChar.Fields("Player")
    cmbClass.Text = rsChar.Fields("Class")
    cmbRace.Text = rsChar.Fields("Race")
    cmbAlign.Text = rsChar.Fields("Alignment")
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

Private Sub cmdUpdate_Click()
    'updates the current record
    With rsChar
        .Edit
        !CharName = txtCharName.Text
        !Player = txtPlayer.Text
        !Class = cmbClass.Text
        !Race = cmbRace.Text
        !Alignment = cmbAlign.Text
        !Level = txtLevel.Text
        !Age = txtAge.Text
        !HitPoint = txtHitPoint.Text
        !ArmorClass = txtArmorClass.Text
        !Str = txtStr.Text
        !Int = txtInt.Text
        !Dex = txtDex.Text
        !Wis = txtWis.Text
        !Con = txtCon.Text
        !Cha = txtCha.Text
        !Breath = txtBreath.Text
        !Rod = txtRod.Text
        !Spell = txtSpell.Text
        !Petr = txtPetr.Text
        !Para = txtPara.Text
        !Armor = txtArmor.Text
        !Helm = txtHelm.Text
        !Shield = txtShield.Text
        !Weapon1 = txtWeapon1.Text
        !Weapon2 = txtWeapon2.Text
        !Magic1 = txtMagic1.Text
        !Magic2 = txtMagic2.Text
        !Magic3 = txtMagic3.Text
        !Magic4 = txtMagic4.Text
        !Magic5 = txtMagic5.Text
        !Notes = txtNotes.Text
        .Update
    End With
    'informs user update is complete.
    MsgBox "The character, " & txtCharName.Text & ", was updated.", vbOKOnly + vbInformation, "DB Update Successfull"
End Sub

