VERSION 5.00
Begin VB.Form frmAddChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Character"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmAddChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      Height          =   375
      Left            =   1920
      TabIndex        =   63
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Opperation"
      Height          =   375
      Left            =   3840
      TabIndex        =   32
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Character"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   7320
      Width           =   1575
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
      TabIndex        =   37
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
      TabIndex        =   36
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
         TabIndex        =   51
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblWeapon1 
         Alignment       =   2  'Center
         Caption         =   "Weapon1"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblShield 
         Alignment       =   2  'Center
         Caption         =   "Shield Carried"
         Height          =   255
         Left            =   1320
         TabIndex        =   49
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblHelm 
         Alignment       =   2  'Center
         Caption         =   "Helm Worn"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblArmor 
         Alignment       =   2  'Center
         Caption         =   "Armor Worn"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame fraSaving 
      Caption         =   "Saving Throws"
      Height          =   2055
      Left            =   2760
      TabIndex        =   35
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
         TabIndex        =   62
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblPetr 
         Caption         =   "Petrification"
         Height          =   255
         Left            =   840
         TabIndex        =   61
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblSpell 
         Caption         =   "Spells/Rings"
         Height          =   255
         Left            =   840
         TabIndex        =   60
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblRod 
         Caption         =   "Rods/Staves et. al"
         Height          =   255
         Left            =   840
         TabIndex        =   59
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblBreath 
         Caption         =   "Breath Weapon"
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "Statistics"
      Height          =   2415
      Left            =   120
      TabIndex        =   34
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
         TabIndex        =   57
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblCon 
         Caption         =   "Constitution"
         Height          =   255
         Left            =   840
         TabIndex        =   56
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblWis 
         Caption         =   "Wisdom"
         Height          =   255
         Left            =   840
         TabIndex        =   55
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblDex 
         Caption         =   "Dexterity"
         Height          =   255
         Left            =   840
         TabIndex        =   54
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblInt 
         Caption         =   "Inelligence"
         Height          =   255
         Left            =   840
         TabIndex        =   53
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblStr 
         Caption         =   "Strength (Power)"
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraPersonal 
      Caption         =   "Personal Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   33
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
         ItemData        =   "frmAddChar.frx":030A
         Left            =   1320
         List            =   "frmAddChar.frx":030C
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
         TabIndex        =   46
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblHitPoint 
         Alignment       =   2  'Center
         Caption         =   "Hip Points"
         Height          =   255
         Left            =   3360
         TabIndex        =   45
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         Caption         =   "Level"
         Height          =   255
         Left            =   2400
         TabIndex        =   44
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblAlign 
         Alignment       =   2  'Center
         Caption         =   "Alignment"
         Height          =   255
         Left            =   840
         TabIndex        =   43
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblAge 
         Alignment       =   2  'Center
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         Caption         =   "Player Name"
         Height          =   255
         Left            =   4080
         TabIndex        =   41
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Caption         =   "Class"
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRace 
         Alignment       =   2  'Center
         Caption         =   "Race"
         Height          =   255
         Left            =   1440
         TabIndex        =   39
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         Caption         =   "Character "
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAddChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmAddCharacter
'| Purpose: Adds a new character to the database.  It also will
'|          generate some data for the new character.
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database          'Sets up the Database Varriable
    Dim rsChar As Recordset     'Sets up Character table as a recordset
    Dim rsRace As Recordset     'sets up Race table as a recordset
    Dim rsClass As Recordset    'sets up Class table as a recordset
    Dim rsAlign As Recordset    'sets up Alignment table as a recorgset

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path + "\cha.mdb") 'Opens database
    Set rsChar = db.OpenRecordset("tblCha")      'populates recordsets
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

End Sub

Private Sub cmdAdd_Click()
    'This adds the new character to the database and runs the update.
    With rsChar
        .AddNew                         'Specifies that we are adding a new record.
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
    
    'Visual Bell to inform user data has been entered into the database.
    MsgBox "Your character, " & txtCharName.Text & " was added to the DB", vbOKOnly + vbInformation, "DB Add Was Successfull"
    
    'closes the add character form.
    cmdCancel_Click
End Sub
    
Private Sub cmdGen_Click()
    'Sets up local variables
    Dim Race, Align, Class, RaceCount, AlignCount, ClassCount
    
    'Initialize the random number generator
    Randomize
    
    'Get a random number based on the number of records in the table.
    Race = Int(Rnd * rsRace.RecordCount) + 1
    Align = Int(Rnd * rsAlign.RecordCount) + 1
    Class = Int(Rnd * rsClass.RecordCount) + 1
    
    'Initialize the counter variables
    RaceCount = 0
    AlignCount = 0
    ClassCount = 0
    
    'Go to the first record of the table.
    rsRace.MoveFirst
    rsAlign.MoveFirst
    rsClass.MoveFirst
    
    'Find the correct record based on the random number generated above.
    Do While RaceCount <= rsRace.RecordCount
        If rsRace.Fields("ID") = Race Then
            cmbRace.Text = rsRace.Fields("Race")
        Else
            rsRace.MoveNext
        End If
        RaceCount = RaceCount + 1
    Loop
    
    'Find the correct record based on the random number generated above.
    Do While AlignCount <= rsAlign.RecordCount
        If rsAlign.Fields("ID") = Align Then
            cmbAlign.Text = rsAlign.Fields("Alignment")
        Else
            rsAlign.MoveNext
        End If
        AlignCount = AlignCount + 1
    Loop
    
    'Find the correct record based on the random number generated above.
    Do While ClassCount <= rsClass.RecordCount
        If rsClass.Fields("ID") = Class Then
            cmbClass.Text = rsClass.Fields("Class")
        Else
            rsClass.MoveNext
        End If
        ClassCount = ClassCount + 1
    Loop
    
    'Generate the values for the attributes of the character.
    txtStr.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtInt.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtDex.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtWis.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtCon.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtCha.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtBreath.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtRod.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtSpell.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtPetr.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtPara.Text = Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1 + Int(Rnd * 6) + 1
    txtLevel.Text = Int(Rnd * 10) + 1
    txtHitPoint.Text = Int(Rnd * txtLevel.Text) + 10
    txtArmorClass.Text = Int(Rnd * 10) - 5
    
    'Inform user the opperation is complete.
    MsgBox "Your character has been generated.  Please fill in missing items and verify and hit the ADD CHARACTER button.", vbOKOnly + vbInformation, "Character Generated"
End Sub

Private Sub cmdCancel_Click()     'closes recordset, database and form
    rsClass.Close
    db.Close
    frmAddChar.Hide
End Sub
