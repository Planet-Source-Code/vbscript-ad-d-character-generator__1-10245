VERSION 5.00
Begin VB.Form frmDelChar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Character"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmDelChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Character"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Opperation"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   600
      TabIndex        =   25
      Top             =   1800
      Width           =   4095
      Begin VB.CommandButton cmdStart 
         Caption         =   "|<<"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRecord 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">>"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   ">>|"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraPersonal 
      Caption         =   "Personal Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtArmorClass 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHitPoint 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtAlign 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRace 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtClass 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCharName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblArmorClass 
         Alignment       =   2  'Center
         Caption         =   "Armor Class"
         Height          =   255
         Left            =   4200
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblHitPoint 
         Alignment       =   2  'Center
         Caption         =   "Hip Points"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         Caption         =   "Level"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblAlign 
         Alignment       =   2  'Center
         Caption         =   "Alignment"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblAge 
         Alignment       =   2  'Center
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         Caption         =   "Player Name"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Caption         =   "Class"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRace 
         Alignment       =   2  'Center
         Caption         =   "Race"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         Caption         =   "Character "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDelChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmDelChar
'| Purpose: Deletes a given character from the database.
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database      'Sets up the Database Varriable
    Dim rsChar As Recordset 'Sets up Character table as a recordset

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'opens database
    Set rsChar = db.OpenRecordset("tblCha")         'populates recordsets.
    rsChar.MoveFirst                                'moves to the first record
    GetData                                         'calls the GetRecords function
End Sub

Private Sub cmdCancel_Click()     'closes recordset, database and form
    rsChar.Close
    db.Close
    frmDelChar.Hide
End Sub

Private Sub cmdDelete_Click()
    rsChar.Delete       'deletes selected character from database
    'informs user deletion is complere
    MsgBox "Character deleted.", vbInformation + vbOKOnly, "Deletion Complete"
End Sub

Private Sub cmdStart_Click() 'Goes to the first record
    On Error Resume Next
    rsChar.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click() 'Goes back a record
    On Error Resume Next
    rsChar.MovePrevious
    GetData
End Sub

Private Sub cmdEnd_Click()  'goes to the last record
    On Error Resume Next
    rsChar.MoveLast
    GetData
End Sub

Private Sub cmdNext_Click() 'goes to the next record.
    On Error Resume Next
    rsChar.MoveNext
    GetData
End Sub

Function GetData()
    'populates the form with data from the database.
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
End Function
