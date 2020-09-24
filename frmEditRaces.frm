VERSION 5.00
Begin VB.Form frmEditRaces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Races"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmEditRaces.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   ">>|"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtRecord 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "|<<"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtRace 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditRaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmEditRaces
'| Purpose: Allows user to add, update and delete races from
'|          the table
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database      'delcares the database variable
    Dim rsRace As Recordset 'declares the recordset variable

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'opens database
    Set rsRace = db.OpenRecordset("tblRace")        'populates recordset
    rsRace.MoveFirst                                'Moves to first record
    GetData                                         'displays record
End Sub

Private Sub cmdExit_Click()     'closes recordset, database and form
    rsRace.Close
    db.Close
    frmEditRaces.Hide
End Sub

Private Sub cmdAdd_Click()
    'Adds new race and informs user opperation is complete
    With rsRace
        .AddNew
        !Race = txtRace.Text
        .Update
    End With
    MsgBox "The class, " & txtRace.Text & " was added to the DB", vbOKOnly + vbInformation, "DB Add Was Successfull"
End Sub

Private Sub cmdDel_Click()
    'deletes current record and informs user opperation is complete
    rsRace.Delete
    MsgBox "Race deleted.", vbInformation + vbOKOnly, "Deletion Complete"
End Sub

Private Sub cmdStart_Click()    'Moves to and displays first record
    On Error Resume Next
    rsRace.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click()     'Moves to and displays previous record
    On Error Resume Next
    rsRace.MovePrevious
    GetData
End Sub

Private Sub cmdNext_Click()     'Moves to and displays next record
    On Error Resume Next
    rsRace.MoveNext
    GetData
End Sub

Private Sub cmdEnd_Click()      'Moves to and displays last record
    On Error Resume Next
    rsRace.MoveLast
    GetData
End Sub

Function GetData()
    'populates form fields
    txtRecord.Text = rsRace.Fields("ID")
    txtRace.Text = rsRace.Fields("Race")
End Function
