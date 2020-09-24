VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormTEST_UCData 
   Caption         =   "UCData"
   ClientHeight    =   11265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   11265
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox ChkHTMShowEmpty 
      Caption         =   "Show Empty HTML Cells"
      Height          =   375
      Left            =   9960
      TabIndex        =   52
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox chkColorHTM 
      Caption         =   "HTM colors are the same"
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   51
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CheckBox chkColorHTM 
      Caption         =   "Use Color in HTM"
      Height          =   375
      Index           =   0
      Left            =   11880
      TabIndex        =   50
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Rechts
      Height          =   285
      Index           =   3
      Left            =   11880
      MaxLength       =   4
      TabIndex        =   43
      Text            =   "0"
      Top             =   3480
      Width           =   550
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Rechts
      Height          =   285
      Index           =   2
      Left            =   11880
      MaxLength       =   4
      TabIndex        =   41
      Text            =   "0"
      Top             =   3120
      Width           =   550
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Rechts
      Height          =   285
      Index           =   1
      Left            =   11880
      MaxLength       =   4
      TabIndex        =   39
      Text            =   "0"
      Top             =   2760
      Width           =   550
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Rechts
      Height          =   285
      Index           =   0
      Left            =   11880
      MaxLength       =   4
      TabIndex        =   37
      Text            =   "0"
      Top             =   2400
      Width           =   550
   End
   Begin VB.CommandButton btnBenDB 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3480
      TabIndex        =   34
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton btnGenTimePlan 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3480
      TabIndex        =   33
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CheckBox chkShowEmptyCols 
      Caption         =   "show empty cols"
      Height          =   255
      Left            =   7920
      TabIndex        =   30
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton btnPaintNew 
      Caption         =   "(re) Paint(new)"
      Height          =   735
      Left            =   5760
      TabIndex        =   29
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox chkShowEmptyRows 
      Caption         =   "show empty rows"
      Height          =   255
      Left            =   7920
      TabIndex        =   28
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton btnGenSchedPlan 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   2160
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   13920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkCenterEntries 
      Caption         =   "Center Entries"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   2760
      Value           =   1  'Aktiviert
      Width           =   1695
   End
   Begin VB.CheckBox chkAutoFill 
      Caption         =   "Autofill Entries"
      Height          =   255
      Left            =   7920
      TabIndex        =   18
      Top             =   2400
      Value           =   1  'Aktiviert
      Width           =   1575
   End
   Begin VB.CommandButton btnGenCal 
      Caption         =   "generate"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown-Liste
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown-Liste
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin UCDataShow.UCData UCData 
      Height          =   6735
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11880
      BackColor       =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      Caption         =   "test"
      RowHeadColor    =   12648447
      ColumnHeadColor =   12640511
      HeaderColor     =   8438015
      EntryColorMain  =   14737632
      AllColsHaveSameSize=   -1  'True
      ShowEmptyRows   =   -1  'True
   End
   Begin VB.CommandButton btnBuild 
      Caption         =   "fill test"
      Height          =   495
      Left            =   8040
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   10
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "remove"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtIDRemove 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "Enter DataID"
      Top             =   720
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WBShow 
      Height          =   6735
      Left            =   7440
      TabIndex        =   3
      Top             =   4440
      Width           =   6975
      ExtentX         =   12303
      ExtentY         =   11880
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Data"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "the control"
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "webresult"
      Height          =   255
      Left            =   7440
      TabIndex        =   53
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblColHTM 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "entry (htm)"
      Height          =   255
      Index           =   3
      Left            =   12960
      TabIndex        =   49
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblColHTM 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "col (htm)"
      Height          =   255
      Index           =   2
      Left            =   12960
      TabIndex        =   48
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lblColHTM 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "row (htm)"
      Height          =   255
      Index           =   1
      Left            =   12000
      TabIndex        =   47
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblColHTM 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Head (htm)"
      Height          =   255
      Index           =   0
      Left            =   12000
      TabIndex        =   46
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblFlags 
      Caption         =   "Flags"
      Height          =   255
      Left            =   7920
      TabIndex        =   45
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblOffset 
      Caption         =   "between columns"
      Height          =   255
      Index           =   3
      Left            =   12960
      TabIndex        =   44
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblOffset 
      Caption         =   "between rows"
      Height          =   255
      Index           =   2
      Left            =   12960
      TabIndex        =   42
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblOffset 
      Caption         =   "first column"
      Height          =   255
      Index           =   1
      Left            =   12960
      TabIndex        =   40
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblOffset 
      Caption         =   "Toprow"
      Height          =   255
      Index           =   0
      Left            =   12960
      TabIndex        =   38
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblOffsets 
      Caption         =   "Offsets"
      Height          =   255
      Left            =   11880
      TabIndex        =   36
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblTCases 
      Caption         =   "Testcases :"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblTCaseDB 
      Caption         =   "Testcase : Database"
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblTTPlan 
      Caption         =   "Testcase TimePlan"
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblTTime 
      Caption         =   "TestCase : Scheduleplan"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label LabelTCal 
      Caption         =   "Testcase : Calendar"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label LblColsBase 
      Caption         =   "ChangeColors (dblCLick on label to change)"
      Height          =   255
      Left            =   10200
      TabIndex        =   24
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblColSet 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Entry"
      Height          =   255
      Index           =   3
      Left            =   11040
      TabIndex        =   23
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblColSet 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Col"
      Height          =   255
      Index           =   2
      Left            =   11040
      TabIndex        =   22
      Top             =   900
      Width           =   735
   End
   Begin VB.Label lblColSet 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Row"
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   21
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblColSet 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Head"
      Height          =   255
      Index           =   0
      Left            =   10200
      TabIndex        =   20
      Top             =   600
      Width           =   1575
   End
   Begin VB.Line Line 
      X1              =   0
      X2              =   14040
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblData 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblID 
      Caption         =   "ID-Row"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   11
      ToolTipText     =   "ID of The Line"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblID 
      Caption         =   "ID-Col"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      ToolTipText     =   "ID of the group"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblID 
      Caption         =   "ID-Entry"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      ToolTipText     =   "ID of the entry"
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "FormTEST_UCData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDoc As HTMLDocument

'//IMPORTANT
'//every ID must gave a value GREATER Null
'//If not, the Value will be ignored

Private Sub btnAdd_Click()

    Call UCData.AddEntry(Val(txtID(0)), Val(txtID(1)), Val(txtID(2)), txtEntry, "Column" & txtID(1), "Line" & txtID(2))
    UCData.PaintGrid
    
    
    Set iDoc = WBShow.Document
    iDoc.body.innerHTML = UCData.HTMLOutString
End Sub

Private Sub SetOptionsToUC()
    UCData.CenterEntry = CBool(Val(chkCenterEntries))
    UCData.FillEntryToColumn = CBool(Val(chkAutoFill))
    
    UCData.ShowEmptyRows = CBool(chkShowEmptyRows.Value)
    UCData.ShowEmptyCols = CBool(chkShowEmptyCols.Value)
    
    UCData.HTMLshowEmptyCells = CBool(ChkHTMShowEmpty.Value)
    
    
'//Offsets
    UCData.OffsetRowFirst = Val(txtOffset(0))
    UCData.OffsetColLeft = Val(txtOffset(1))
    UCData.OffsetBetweenRows = Val(txtOffset(2))
    UCData.OffsetBetweenColumns = Val(txtOffset(3))
'//Colors
    UCData.HeaderColor = lblColSet(0).BackColor
    UCData.RowHeadColor = lblColSet(1).BackColor
    UCData.ColumnHeadColor = lblColSet(2).BackColor
    UCData.EntryColorMain = lblColSet(3).BackColor
    
    UCData.HTMLuseColors = CBool(chkColorHTM(0))
    UCData.ColorHtmIsSameAsControl = CBool(chkColorHTM(1))
    UCData.ColorHTMHeader = lblColHTM(0).Tag
    UCData.ColorHTMColumn = lblColHTM(2).Tag
    UCData.ColorHTMRow = lblColHTM(1).Tag
    UCData.ColorHTMEntry = lblColHTM(3).Tag
End Sub

Private Sub btnBenDB_Click()
    Set UCData.DBConnection = ConnectDataBase(App.Path & "\test.mdb")
    UCData.ClearGrid
    UCData.ReadColumnsDB "ColData", "ID", "Value"
    UCData.ReadRowsDB "RowData", "ID", "Value"
    UCData.ReadEntriesDB "EntryData", "ID", "Value", "RowID", "ColID"
    UCData.DBConnection.Close
    UCData.PaintGrid
    UCData.GetHTMLout
    WBShow.Document.body.innerHTML = UCData.GetHTMLout()
    MsgBox "now change some data in the database and reread it", , ""
End Sub

Private Sub btnBuild_Click()
    UCData.ClearGrid
'    UCData.OffsetBetweenColumns = 200
'    UCData.OffsetBetweenRows = 100
    SetOptionsToUC
    
    Call UCData.AddEntry(1, 1, 1, "TestGJEntry-1", "TCOl", "TLine and 1000")
    Call UCData.AddEntry(2, 2, 1, "TestGJEntry--2", "TcOl", "TLine and 1000")
    Call UCData.AddEntry(3, 1, 1, "TestGJEntry-3-", "TCOl", "TLine and 1000")
    Call UCData.AddEntry(4, 3, 1, "TestGJEntry-*4d", "TcoL", "TLine and 1000")
    Call UCData.AddEntry(5, 1, 2, "TestGJEntry--(5)", "TcOL", "TLinsecond")
'    Call UCData.AddEntry(6, 4, 1, "TestGJEntry+7*", "Tcol", "TLine and 1000")
'    Call UCData.AddEntry(7, 5, 1, "TestGJEntry+8kk", "Tcol", "TLine and 1000")
'    Call UCData.AddEntry(8, 6, 1, "TestGJEntry+9pq24", "Tcol", "TLine and 1000")
    UCData.PaintGrid
    Do While iDoc Is Nothing
        Set iDoc = WBShow.Document
        DoEvents
    Loop
    
    iDoc.body.innerHTML = UCData.HTMLOutString
End Sub

Private Sub btnClear_Click()
    UCData.ClearGrid
    UCData.GetHTMLout
    WBShow.Document.body.innerHTML = UCData.HTMLOutString
End Sub

Private Sub btnGenCal_Click()
Dim lMonth As Long, lYear As Long, lRow As Long, lMax As Long, wDay As Long, wdLast As Long, n As Long
    UCData.ClearGrid
    SetOptionsToUC
    lMonth = Val(cboMonth)
    lYear = Val(cboYear)
    Select Case lMonth
        Case 1, 3, 5, 7, 8, 10, 12
            lMax = 31
        Case 2
            If lYear Mod 4 = 0 Then
                lMax = 29
            Else
                lMax = 28
            End If
        Case 4, 6, 9, 11
            lMax = 30
    End Select
    lRow = 1
    wdLast = 0
    For n = 2 To 7
        UCData.AddColumn n, Format(n, "DDDD", vbSunday)
    Next
    UCData.AddColumn 1, Format(1, "DDDD", vbSunday)
    'UCData.PaintGrid
    lRow = Format(DateSerial(lYear, lMonth, 1), "ww")
    For n = 1 To lMax
        wDay = Weekday(DateSerial(lYear, lMonth, n), vbSunday)
        UCData.AddEntry n, wDay, lRow, CStr(n), , CStr(lRow)
        If wDay < wdLast Then
            lRow = lRow + 1
        End If
        wdLast = wDay
    Next
    UCData.PaintGrid
    UCData.GetHTMLout
    WBShow.Document.body.innerHTML = UCData.HTMLOutString
End Sub



Private Sub btnGenSchedPlan_Click()
Dim n As Long, t As Long, d As Long, ts As Single
    UCData.ClearGrid
    SetOptionsToUC
    
    For n = 2 To 7
        UCData.AddColumn n, Format(n, "DDDD", vbSunday)
    Next
    UCData.AddColumn 1, Format(1, "DDDD", vbSunday)
    For n = 9 To 18
        UCData.AddRow n, Format(n, "00") & ":00"
    Next
    Randomize
    For n = 1 To 20
        ts = Rnd(n * 7) * 100
        t = ts
        If t < 9 Then t = 9
        If t > 18 Then t = 18
        ts = Rnd(n * 8) * 100
        d = ts
        If d < 1 Then d = 1
        If d > 7 Then d = 7
        UCData.AddEntry n, d, t, "Sched " & n
    Next
    UCData.PaintGrid
    UCData.GetHTMLout
    WBShow.Document.body.innerHTML = UCData.HTMLOutString
    
End Sub

Private Sub btnGenTimePlan_Click()
Dim n As Long, mEn As Long, mC As Long, mR As Long
    UCData.ClearGrid
    SetOptionsToUC
    For n = 1 To 8
        UCData.AddColumn n, "Day " & n
    Next
    For n = 1 To 8
        UCData.AddRow n, "DItem " & n
    Next
    Randomize
    For n = 1 To 36
        mC = Rnd(8) * 8
        mR = Rnd(8) * 8
        If mC < 1 Then mC = 1
        If mC > 8 Then mC = 8
        If mR < 1 Then mR = 1
        If mR > 8 Then mR = 8
        
        UCData.AddEntry n, mC, mR, "Entry " & n
    Next
    UCData.PaintGrid
    UCData.GetHTMLout
    WBShow.Document.body.innerHTML = UCData.HTMLOutString
End Sub

Private Sub btnPaintNew_Click()
    SetOptionsToUC
    WBShow.Document.body.innerHTML = ""
    UCData.PaintGrid
    UCData.GetHTMLout
    WBShow.Document.body.innerHTML = UCData.GetHTMLout
End Sub



Private Sub btnRemove_Click()
    UCData.RemoveEntry Val(txtIDRemove)
    UCData.PaintGrid
    UCData.GetHTMLout
    WBShow.Document.body.innerHTML = UCData.HTMLOutString
End Sub

Private Sub Form_Load()
Dim n As Integer
    WBShow.Navigate "about:blank"
    UCData.Caption = "Test me"
    For n = 1 To 12
        cboMonth.AddItem n
    Next
    For n = 1 To 12
        cboYear.AddItem "20" + Format(n, "00")
    Next
    cboMonth.Text = Month(Date)
    cboYear.Text = Year(Date)
    lblColSet(0).BackColor = UCData.HeaderColor
    lblColSet(1).BackColor = UCData.RowHeadColor
    lblColSet(2).BackColor = UCData.ColumnHeadColor
    lblColSet(3).BackColor = UCData.EntryColorMain
    lblColHTM(0).BackColor = vbWhite
    lblColHTM(1).BackColor = vbWhite
    lblColHTM(2).BackColor = vbWhite
    lblColHTM(3).BackColor = vbWhite
    lblColHTM(0).Tag = "#FFFFFF"
    lblColHTM(1).Tag = "#FFFFFF"
    lblColHTM(2).Tag = "#FFFFFF"
    lblColHTM(3).Tag = "#FFFFFF"
End Sub


Private Sub lblColHTM_Click(Index As Integer)
    CDlg.Color = lblColHTM(Index).BackColor
    CDlg.Action = 3
    lblColHTM(Index).BackColor = CDlg.Color
    lblColHTM(Index).Tag = UCData.OleCOlToHTM(lblColHTM(Index).BackColor)
End Sub

Private Sub lblColSet_DblClick(Index As Integer)
    CDlg.Color = lblColSet(Index).BackColor
    CDlg.Action = 3
    lblColSet(Index).BackColor = CDlg.Color
End Sub

Private Sub UCData_EntrySelected(EntryID As Long)
    Debug.Print "Entry ID", EntryID
    txtID(0) = EntryID
    txtID(2) = UCData.EntrySelectedColumnID
    txtID(1) = UCData.EntrySelectedRowID
    txtEntry = UCData.EntrySelectedText
End Sub
