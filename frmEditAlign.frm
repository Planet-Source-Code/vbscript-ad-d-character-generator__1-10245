VERSION 5.00
Begin VB.Form frmEditAlign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Alignment"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmEditAlign.frx":0000
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
   Begin VB.TextBox txtAlign 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditAlign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmEditAlign
'| Purpose: Allows user to add, edit and delete records from
'|          the alignment table
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database          'Sets up the database varriable
    Dim rsAlign As Recordset    'defines the recordset name

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'opens the database
    Set rsAlign = db.OpenRecordset("tblAlign")      'populates the recordset
    rsAlign.MoveFirst                               'goes to the first record
    GetData                                         'calls GetData function
End Sub

Private Sub cmdExit_Click()     'closes recordset, database and form
    rsAlign.Close
    db.Close
    frmEditAlign.Hide
End Sub

Private Sub cmdAdd_Click()
    'adds new alignment to the table
    With rsAlign
        .AddNew
        !Alignment = txtAlign.Text
        .Update
    End With
    'informs user update is finished.
    MsgBox "The Alignment, " & txtAlign.Text & " was added to the DB", vbOKOnly + vbInformation, "DB Add Was Successfull"
End Sub

Private Sub cmdDel_Click()
    rsAlign.Delete      'deletes record from database.
    'informs user delete is finished.
    MsgBox "Alignment deleted.", vbInformation + vbOKOnly, "Deletion Complete"
End Sub

Private Sub cmdStart_Click()    'Goes to and displays the first record
    On Error Resume Next
    rsAlign.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click()     'Goes to and displays the previous record
    On Error Resume Next
    rsAlign.MovePrevious
    GetData
End Sub

Private Sub cmdNext_Click()     'goes to and displays the next record
    On Error Resume Next
    rsAlign.MoveNext
    GetData
End Sub

Private Sub cmdEnd_Click()      'goes to ans displays the last record
    On Error Resume Next
    rsAlign.MoveLast
    GetData
End Sub

Function GetData()              'populates the form.
    txtRecord.Text = rsAlign.Fields("ID")
    txtAlign.Text = rsAlign.Fields("Alignment")
End Function
