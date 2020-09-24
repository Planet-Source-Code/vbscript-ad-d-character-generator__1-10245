VERSION 5.00
Begin VB.Form frmEditClasses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Classes"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmEditClasses.frx":0000
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
   Begin VB.TextBox txtClass 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmEditClasses
'| Purpose: Allows user to add, edit and delete classes from
'|          table
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database          'declares the database variable
    Dim rsClass As Recordset    'declares the recordset variable

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'opens database
    Set rsClass = db.OpenRecordset("tblClass")      'populates recordset
    'moves to first record and displays
    rsClass.MoveFirst
    GetData
End Sub

Private Sub cmdExit_Click()     'closes recordset, database and form
    rsClass.Close
    db.Close
    frmEditClasses.Hide
End Sub

Private Sub cmdAdd_Click()
    'Adds new class to database
    With rsClass
        .AddNew
        !Class = txtClass.Text
        .Update
    End With
    'informs user that the process is complete
    MsgBox "The class, " & txtClass.Text & " was added to the DB", vbOKOnly + vbInformation, "DB Add Was Successfull"
End Sub

Private Sub cmdDel_Click()
    'deletes record and informs user process is complete
    rsClass.Delete
    MsgBox "Class deleted.", vbInformation + vbOKOnly, "Deletion Complete"
End Sub

Private Sub cmdStart_Click()    'Moves to and displays the first record
    On Error Resume Next
    rsClass.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click()     'Moves to and displays the previous record
    On Error Resume Next
    rsClass.MovePrevious
    GetData
End Sub

Private Sub cmdNext_Click()     'Moves to and displays the next record
    On Error Resume Next
    rsClass.MoveNext
    GetData
End Sub

Private Sub cmdEnd_Click()      'Moves to and displays the last record
    On Error Resume Next
    rsClass.MoveLast
    GetData
End Sub

Function GetData()
    'Populates the form with current record
    txtRecord.Text = rsClass.Fields("ID")
    txtClass.Text = rsClass.Fields("Class")
End Function
