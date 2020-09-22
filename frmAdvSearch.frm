VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdvSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance Database Search"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdvSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   360
      Left            =   3000
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   645
      ItemData        =   "frmAdvSearch.frx":08CA
      Left            =   2640
      List            =   "frmAdvSearch.frx":08CC
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelOperation 
      Caption         =   "Cancel Operation"
      Height          =   360
      Left            =   6480
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   3360
   End
   Begin VB.ListBox List2 
      Height          =   2985
      ItemData        =   "frmAdvSearch.frx":08CE
      Left            =   4680
      List            =   "frmAdvSearch.frx":08D0
      TabIndex        =   11
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   360
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   990
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   5040
      TabIndex        =   9
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   360
      Left            =   3000
      TabIndex        =   6
      Top             =   2760
      Width           =   1230
   End
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   4080
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChooseDB 
      Caption         =   "Select DB"
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve"
      Height          =   360
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "frmAdvSearch.frx":08D2
      Left            =   120
      List            =   "frmAdvSearch.frx":08D4
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblDoubleClick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Double Click on List item to copy on clipboard"
      Height          =   195
      Left            =   4800
      TabIndex        =   14
      Top             =   4800
      Width           =   3345
   End
   Begin VB.Label lblTableName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name : Field Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   720
      Width           =   105
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password : "
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   4200
      Width           =   840
   End
   Begin VB.Shape Shape2 
      Height          =   135
      Left            =   120
      Top             =   4680
      Width           =   4200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   120
      Top             =   4680
      Width           =   15
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   45
   End
   Begin VB.Label lblTableNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table Names"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label lblConnectionString 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "frmAdvSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbPath As String, dbPass As String
Dim cnnString As String, oprCancel As Boolean

Private Function TableNames(ConnectionString As String) _
  As Collection

On Error GoTo errHandler
Dim oCatalog As New ADOX.Catalog
Dim oTableNames As New Collection
Dim oTables As ADOX.Tables
Dim oTable As ADOX.Table
Dim oConnection As New ADODB.Connection

oConnection.ConnectionString = ConnectionString
oConnection.Open ConnectionString
Set oCatalog.ActiveConnection = oConnection
Set oTables = oCatalog.Tables

For Each oTable In oTables
    oTableNames.Add oTable.Name
Next
Set TableNames = oTableNames

errHandler:

On Error Resume Next
If oConnection.State <> 0 Then oConnection.Close
Set oConnection = Nothing
Set oCatalog = Nothing
Set oTable = Nothing
Set oTables = Nothing

End Function


Private Sub cmdCancelOperation_Click()
    oprCancel = True
End Sub

Private Sub cmdChooseDB_Click()
    cdDialog.Filter = "Access (*.mdb) | *.mdb"
    cdDialog.ShowOpen
    dbPath = cdDialog.FileName
    If dbPath <> "" Then lblConnectionString = "Your database : " & cdDialog.FileName
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRetrieve_Click()
        '<EhHeader>
        On Error GoTo cmdRetrieve_Click_Err
        '</EhHeader>
100     oprCancel = False
        DisableObjects
116     dbPass = txtPassword
118     Label1 = ""
120     List1.Clear
122     If dbPath = "" Then
124         MsgBox "Please choose any Access database 1st", vbCritical
            Exit Sub
        End If
        
126     cnnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & dbPath & "';Jet Oledb:database password=" & dbPass & ";"
    
        Dim i As Integer, j As Integer, cnt As Integer
128     cnt = 0
130     j = TableNames(cnnString).Count / 100
132     For i = 1 To TableNames(cnnString).Count
134         If Not (Mid(TableNames(cnnString).Item(i), 1, 4)) = "MSys" Then List1.AddItem TableNames(cnnString).Item(i)
136         DoEvents
        
138         Label1.Caption = Format(i * 100 / TableNames(cnnString).Count, "#.00") & "%"
140         DoEvents
        
        
142         Shape1.Width = i * 100 / TableNames(cnnString).Count * 42
144         DoEvents
146         If oprCancel = True Then Exit For
        Next
148     If oprCancel = True Then
150         MsgBox "Operation canceled", vbInformation
152         oprCancel = False
154         cmdChooseDB.Enabled = True
156         cmdExit.Enabled = True
158         cmdRetrieve.Enabled = True
160         cmdSearch.Enabled = True
162         Timer1.Enabled = False
164         Label2.Caption = ""
            Exit Sub
        End If
166     MsgBox "Done!!", vbInformation, "Table Names from Database"
    
        EnableObjects
        
        Exit Sub

cmdRetrieve_Click_Err:
180     MsgBox "Your database may be a password protected." & vbNewLine & "Please enter valid password.", vbExclamation + vbOKOnly, "Application Error"
182     cmdChooseDB.Enabled = True
184     cmdExit.Enabled = True
186     cmdRetrieve.Enabled = True
188     cmdSearch.Enabled = True
190     Timer1.Enabled = False
192     Label2.Caption = ""
End Sub

Private Sub cmdSearch_Click()
        '<EhHeader>
        On Error GoTo cmdSearch_Click_Err
        '</EhHeader>
100     If List1.ListCount = 0 Then
102         MsgBox "Please retrive Tables from Selecting Database", vbInformation, "Database"
            Exit Sub
        End If
104     If txtSearch = "" Then
106         MsgBox "Please enter any text to search", vbInformation, "Database"
            Exit Sub
        End If
        Dim flgSelect As Boolean
        Dim i As Integer, flgFound As Boolean
108     flgSelect = False
110     For i = 0 To List1.ListCount - 1
112         If List1.Selected(i) = True Then flgSelect = True
        Next
        
114     If flgSelect = False Then
116         MsgBox "Please select any table from the list to search", vbInformation
            Exit Sub
        End If
        List3.Clear
118     For i = 0 To List1.ListCount - 1
120         If List1.Selected(i) = True Then List3.AddItem List1.List(i)
        Next
        Dim ddd
122     ddd = MsgBox("Searching may take long time. Are you sure to want to Search?", vbYesNo, "Table Names from Database")
124     If ddd = vbNo Then Exit Sub
        Dim conn As New ADODB.Connection, rs As New ADODB.Recordset
126     conn.Open cnnString
        
128     flgFound = False
130     List2.Clear
        List1.Enabled = False
        DisableObjects
146     For i = 0 To List3.ListCount - 1
148         If rs.State = 1 Then rs.Close
150         rs.Open "select * from " & List3.List(i), conn, adOpenDynamic, adLockOptimistic
152         DoEvents
154         If Not (rs.EOF And rs.BOF) Then
                Dim j As Integer
156             DoEvents
158             For j = 0 To rs.Fields.Count - 1
160                 rs.MoveFirst
162                 DoEvents
164                 While Not rs.EOF
166                     DoEvents
168                     If UCase(rs.Fields(j)) = UCase(txtSearch) Then
170                         DoEvents
172                         flgFound = True
174                         DoEvents
176                         List2.AddItem List3.List(i) & " : " & rs.Fields(j).Name
178                         DoEvents

                        End If
180                     DoEvents
182                     Label1.Caption = Format(i * 100 / TableNames(cnnString).Count, "#.00") & "%"
184                     DoEvents
186                     Shape1.Width = i * 100 / TableNames(cnnString).Count * 42
188                     DoEvents
190                     rs.MoveNext
192                     If oprCancel = True Then Exit For
                    Wend
194                 Label1.Caption = Format((i + 1) * 100 / List3.ListCount, "#.00") & "%"
196                 DoEvents
198                 If oprCancel = True Then Exit For
                Next
200             Label1.Caption = Format((i + 1) * 100 / List3.ListCount, "#.00") & "%"
202             DoEvents
            End If
204         Label1.Caption = Format((i + 1) * 100 / List3.ListCount, "#.00") & "%"
206         DoEvents
208         Shape1.Width = (i + 1) * 100 / List3.ListCount * 42
210         DoEvents
212         If oprCancel = True Then Exit For
        Next
214     If oprCancel = True Then
216         MsgBox "Operation canceled", vbInformation
218         oprCancel = False
220         cmdChooseDB.Enabled = True
222         cmdExit.Enabled = True
224         List1.Enabled = True
226         cmdRetrieve.Enabled = True
228         cmdSearch.Enabled = True
230         Timer1.Enabled = False
232         Label2.Caption = ""
            Exit Sub
        End If
234     MsgBox "Done!!", vbInformation, "Table Names from Database"
236     If List2.ListCount = 0 Then
238         MsgBox "Record not found in the database", vbInformation
        Else
240         MsgBox "Found " & List2.ListCount & " Records", vbInformation
        End If
242     oprCancel = False
        EnableObjects
        '<EhFooter>
        Exit Sub

cmdSearch_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SearchDB.Form1.cmdSearch_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        '</EhFooter>
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = True
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List2_DblClick()
    Clipboard.SetText List2.List(List2.ListIndex)
End Sub

Private Sub Timer1_Timer()
If Label2.Caption = "o/o" Then
    Label2.Caption = "o\o"
Else
    Label2.Caption = "o/o"
End If
End Sub

Private Sub DisableObjects()
    cmdSelectAll.Enabled = False
    Shape1.Width = 0
    Timer1.Enabled = True
    cmdChooseDB.Enabled = False
    cmdExit.Enabled = False
    cmdRetrieve.Enabled = False
    cmdSearch.Enabled = False
End Sub

Private Sub EnableObjects()
    cmdChooseDB.Enabled = True
    cmdExit.Enabled = True
    List1.Enabled = True
    cmdRetrieve.Enabled = True
    cmdSearch.Enabled = True
    cmdSelectAll.Enabled = True
    Timer1.Enabled = False
    Label2.Caption = ""
End Sub
