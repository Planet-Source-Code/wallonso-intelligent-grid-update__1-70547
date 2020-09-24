VERSION 5.00
Begin VB.UserControl UCData 
   BackColor       =   &H80000009&
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ScaleHeight     =   3735
   ScaleWidth      =   6435
   Begin VB.HScrollBar HScroll 
      Height          =   185
      Left            =   960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2760
      Width           =   3015
   End
   Begin VB.VScrollBar VScroll 
      Height          =   2655
      Left            =   4920
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   185
   End
   Begin VB.PictureBox PicOutData 
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   5  'Transparent
      Height          =   2175
      Left            =   960
      ScaleHeight     =   2115
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   600
      Width           =   3855
      Begin VB.PictureBox PicDataEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   1995
         TabIndex        =   6
         Top             =   0
         Width           =   2055
         Begin VB.Label lblEntry 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Entry"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox PicOutRow 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   600
      Width           =   975
      Begin VB.PictureBox PicDataRow 
         BackColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   0
         ScaleHeight     =   1515
         ScaleWidth      =   795
         TabIndex        =   4
         Top             =   0
         Width           =   855
         Begin VB.Label lblRowEntry 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00C0FFFF&
            Caption         =   "Row"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
      End
   End
   Begin VB.PictureBox PicOutCols 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   240
      Width           =   3855
      Begin VB.PictureBox PicDataCol 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   2475
         TabIndex        =   3
         Top             =   0
         Width           =   2535
         Begin VB.Label lblColumn 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00C0E0FF&
            Caption         =   "Col"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
      End
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      Caption         =   "name"
      Height          =   235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "UCData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

Event EntrySelected(EntryID As Long)
Event EntryDblClick(EntryID As Long)
Event EntryRightClick(EntryID As Long)

Event RowSelected(EntryID As Long)
Event RowDblClick(EntryID As Long)
Event RowRightClick(EntryID As Long)

Event ColumnSelected(EntryID As Long)
Event ColumnDblClick(EntryID As Long)
Event ColumnRightClick(EntryID As Long)



Const m_def_MinWidthRow = 1000
Const m_def_MinWidthCol = 1000
Const m_def_OffsetColLeft = 0
Const m_def_OffsetRowFirst = 0
Const m_def_ShowEmptyCols = False
Const m_def_ShowEmptyRows = False
Const m_def_AllColsHaveSameSize = False
Const m_def_CenterEntry = True
Const m_def_FillEntryToColumn = True
Const m_def_ForeColor = 0
Const m_def_RowHeadColor = 0
Const m_def_ColumnHeadColor = 0
Const m_def_HeaderColor = 0
Const m_def_EntryColorMain = 0
Const m_def_Enabled = 0
Const m_def_BackColor = &H80000009
Const m_def_IncludeBodyHTM = True
Const m_def_ColorHTMHeader = "#FFFFFF"
Const m_def_ColorHTMColumn = "#FFFFFF"
Const m_def_ColorHTMRow = "#FFFFFF"
Const m_def_ColorHTMEntry = "#FFFFFF"
Const m_def_ColorHtmIsSameAsControl = False
Const m_def_HTMLuseColors = False
Const m_def_HTMLshowEmptyCells = False

Dim m_MinWidthRow As Single
Dim m_MinWidthCol As Single

Dim mvar_OffsetColLeft As Single
'Dim mvar_OffsetRowFirst As Single
'//Additional Vars/Flags for Painting
'Private mvarOffsetRowLeft As Single '//XPos Where the RowLabel starts
Private mvar_OffsetRowFirst As Single '//XPos Where the firstentry starts in a Row
Private mvar_OffsetRows As Single '//Delta between two Rows
'Private mvarOffsetColTop As Single '//YPos Where the ColLabel starts
'Private mvarOffsetColFirst As Single '//YPos Where the firstentry starts in a Column
Private mvar_OffsetColumns As Single '//Delta between two Columns


Dim m_BackStyle As AmbientProperties
Dim m_ShowEmptyCols As Boolean
Dim m_ShowEmptyRows As Boolean
Dim m_AllColsHaveSameSize As Boolean
Private mvarCenterEntry As Boolean
Private mvarFillEntryToColumn As Boolean

Dim m_ForeColor As OLE_COLOR
Dim m_RowHeadColor As OLE_COLOR
Dim m_ColumnHeadColor As OLE_COLOR
Dim m_HeaderColor As OLE_COLOR
Dim m_EntryColorMain As OLE_COLOR

'//Converted Olecolors to htmColors
'Dim m_ColorHTM_Head As Long
'Dim m_ColorHTM_ColHead As Long
'Dim m_ColorHTM_RowHead As Long
'Dim m_ColorHTM_Entry As Long
Dim mvar_IncludeBodyHTM As Boolean


Dim m_ColorHTMHeader As String
Dim m_ColorHTMColumn As String
Dim m_ColorHTMRow As String
Dim m_ColorHTMEntry As String
Dim m_ColorHtmIsSameAsControl As Boolean
Dim m_HTMLuseColors As Boolean
Dim m_HTMLshowEmptyCells As Boolean






Dim m_Enabled As Boolean
Dim m_Font As Font
'Dim m_BackStyle As Integer



Private Type tpLineColEntry
    colDataID As Long   '//only (DB-)index of Entry
    colDataNIDx As Long   '//only (DB-)index of Entry
    colDataLeft As Single
    colDataWidth As Single
End Type

Private Type tpLineInfo
    HasContent As Boolean
    RowID As Long
    NumCols As Long
    ColEntry() As tpLineColEntry
    Height As Single    '//Line height
End Type

Private Type ROWINFO
    ID As Long
    'sCaption As String
    'MinWidth As Single
    TextWidth As Single
    Xpos As Single
    YPos As Single
    
    NumEntries As Long
    EntriesID() As Long
    EntriesNIdx() As Long
    
'//Counter for HTML, how many Items per Row Max
    LineCount As Long '//Number of Entries in Line
    LineHeight As Single '//Heigth depending on entries in Line
    bHitInPaint As Boolean
    
    LineInfo() As tpLineInfo
    
End Type
Private Type COLINFO
    ID As Long
    'sCaption As String
    ActWidth As Single
    TextWidth As Single
    showWidth As Single
    
'    MinWidth As Single
'    MaxWidth As Single
    
    Xpos As Single
    YPos As Single
    
    NumEntries As Long  '//= ColCount
    EntriesID() As Long   '//The Id's of the Entries
    EntriesNIdx() As Long '//The CounterIndex (mostly NH)
    EntriesWidth() As Single    '//And their actual Widths
    EntriesMaxWidth As Single
    bHitInPaint As Boolean
End Type
Private Type DATAINFO
    ID As Long
    RowID As Long
    ColID As Long
    sCaption As String
    Height As Single
    
    ListWidth As Single
    ListHeight As Single
    MinWidth As Single
    MinHeight As Single
    '//These 3 show the entry
    Xpos As Single
    YPos As Single
    Width As Single
    
    NumEntries As Long
    Entries() As Long   '//The Id's of the Entries
    EntriesHeight() As Single    '//And their actual Widths
    EntriesMaxHeight As Single
    
    MaxHeightY As Single
    MaxWidthX As Single
    'numLines As Long
End Type



'//Structure with Items
Private mta_InfoRow() As ROWINFO
Private mta_InfoColumn() As COLINFO
Private mta_InfoEntry() As DATAINFO

'//Number of Items
Private mvar_NumEntrys As Long
Private mvar_NumColumns As Long
Private mvar_NumRows As Long


'//ColumnData
Private ma_lColEntryPosX() As Single
Private ma_lColEntryNumEntrys() As Long

'//RowData
Private ma_lRowEntryNumEntrys() As Long


Private mvar_RowLabelMaxWidth As Single '//Max Width of Row Labels
Private ma_bEntryHitInPaint() As Boolean    '//Flag needed for Paint


Private m_UCXwidthNeeded As Single, m_UCXHeightNeeded As Single






Private mvarAutoSizeGrid As Boolean

'//Outputstring for HTML (Porperty)
Private mvar_sHTMLOUT As String
Private mvar_RSTOut As ADODB.Recordset

Private mvarSelectedEntry_Text As String
'Private mvarSelectedEntry_Value As String
Private mvarSelectedEntry_RowID As Long
Private mvarSelectedEntry_ColID As Long


Private mvar_DBConnection As ADODB.Connection

Public Property Set DBConnection(ByVal newData As ADODB.Connection)
    Set mvar_DBConnection = newData
End Property
Public Property Get DBConnection() As ADODB.Connection
    Set DBConnection = mvar_DBConnection
End Property


Public Function ReadColumnsDB(sTable As String, sFieldID As String, sFieldText As String, Optional sOrderBy As String)
Dim strSQL As String, rst As ADODB.Recordset
    If Not mvar_DBConnection Is Nothing Then
        strSQL = "SELECT * FROM " & sTable
        If Len(sOrderBy) Then
            strSQL = strSQL & " ORDERBY " & sOrderBy
        End If
        Set rst = New ADODB.Recordset
        rst.Open strSQL, mvar_DBConnection, adOpenStatic
        Do While rst.EOF = False And rst.BOF = False
            AddColumn rst.Fields(sFieldID), rst.Fields(sFieldText)
            rst.MoveNext
        Loop
        If rst.State Then rst.Close
        Set rst = Nothing
    End If
End Function
Public Function ReadRowsDB(sTable As String, sFieldID As String, sFieldText As String, Optional sOrderBy As String)
Dim strSQL As String, rst As ADODB.Recordset
    If Not mvar_DBConnection Is Nothing Then
        strSQL = "SELECT * FROM " & sTable
        If Len(sOrderBy) Then
            strSQL = strSQL & " ORDERBY " & sOrderBy
        End If
        Set rst = New ADODB.Recordset
        rst.Open strSQL, mvar_DBConnection, adOpenStatic
        Do While rst.EOF = False And rst.BOF = False
            AddRow rst.Fields(sFieldID), rst.Fields(sFieldText)
            rst.MoveNext
        Loop
        If rst.State Then rst.Close
        Set rst = Nothing
    End If
End Function

Public Function ReadEntriesDB(sTable As String, sFieldIDEntry As String, sFieldText As String, sFieldIDRow As String, sFieldIDColumn As String, Optional sOrderBy As String)
Dim strSQL As String, rst As ADODB.Recordset
    If Not mvar_DBConnection Is Nothing Then
        strSQL = "SELECT * FROM " & sTable
        If Len(sOrderBy) Then
            strSQL = strSQL & " ORDERBY " & sOrderBy
        End If
        Set rst = New ADODB.Recordset
        rst.Open strSQL, mvar_DBConnection, adOpenStatic
        Do While rst.EOF = False And rst.BOF = False
            AddEntry rst.Fields(sFieldIDEntry), rst.Fields(sFieldIDColumn), rst.Fields(sFieldIDRow), rst.Fields(sFieldText)
            rst.MoveNext
        Loop
        If rst.State Then rst.Close
        Set rst = Nothing
    End If
End Function

Public Property Get EntrySelectedRowID() As Long
    EntrySelectedRowID = mvarSelectedEntry_RowID
End Property
Public Property Get EntrySelectedColumnID() As Long
    EntrySelectedColumnID = mvarSelectedEntry_ColID
End Property
Public Property Get EntrySelectedText() As String
    EntrySelectedText = mvarSelectedEntry_Text
End Property

Public Function GetRowIDFromEntry(EntryID As Long) As Long
Dim n As Long
    For n = 1 To mvar_NumEntrys
        If mta_InfoEntry(n).ID = EntryID Then
            GetRowIDFromEntry = mta_InfoEntry(n).RowID
            Exit For
        End If
    Next
End Function
Public Function GetColumnIDFromEntry(EntryID As Long) As Long
Dim n As Long
    For n = 1 To mvar_NumEntrys
        If mta_InfoEntry(n).ID = EntryID Then
            GetColumnIDFromEntry = mta_InfoEntry(n).ColID
            Exit For
        End If
    Next
End Function

Public Function GetInfoFromEntry(EntryID As Long, pRetEntryText As String, pRetColID As Long, pRetRowID As Long)
Dim n As Long
    For n = 1 To mvar_NumEntrys
        If mta_InfoEntry(n).ID = EntryID Then
            pRetEntryText = mta_InfoEntry(n).sCaption
            pRetRowID = mta_InfoEntry(n).RowID
            pRetColID = mta_InfoEntry(n).ColID
            Exit For
        End If
    Next
End Function

Public Property Get HTMLOutString() As String
    HTMLOutString = mvar_sHTMLOUT
End Property
Public Property Get RecordsetOut() As ADODB.Recordset
    Set RecordsetOut = mvar_RSTOut
End Property


Public Property Let OffsetBetweenRows(ByVal newData As Single)
    mvar_OffsetRows = newData
End Property
Public Property Get OffsetBetweenRows() As Single
    OffsetBetweenRows = mvar_OffsetRows
End Property
Public Property Let OffsetBetweenColumns(ByVal newData As Single)
    mvar_OffsetColumns = newData
End Property
Public Property Get OffsetBetweenColumns() As Single
    OffsetBetweenColumns = mvar_OffsetColumns
End Property

'//Preset Column
Public Function AddColumn(ColumnID As Long, ColCaption As String) As Long
Dim n As Long, bFOund As Boolean, nFOund As Long, nx As Single
    For n = 1 To mvar_NumColumns
        If mta_InfoColumn(n).ID = ColumnID Then
            bFOund = True
            nFOund = n
            Exit For
        End If
    Next
    If bFOund = False Then
        mvar_NumColumns = mvar_NumColumns + 1
        ReDim Preserve mta_InfoColumn(mvar_NumColumns)
        
        ReDim Preserve ma_lColEntryPosX(mvar_NumColumns)
        ReDim Preserve ma_lColEntryNumEntrys(mvar_NumColumns)
        'ReDim Preserve ma_lColumnWidth(mvar_NumColumns)
        Load lblColumn(mvar_NumColumns)
        lblColumn(mvar_NumColumns).BackColor = m_ColumnHeadColor
        mta_InfoColumn(mvar_NumColumns).ID = ColumnID '//Save Database-ID
        ma_lColEntryPosX(mvar_NumColumns) = 200 '//Just an InitValue
        lblColumn(mvar_NumColumns).Tag = ColumnID
        
               
        lblColumn(mvar_NumColumns).Caption = ColCaption ' & " C:" & ColEntryID
        nx = UserControl.TextWidth(ColCaption & "  ")
        If lblColumn(mvar_NumColumns).Width < nx Then
            lblColumn(mvar_NumColumns).Width = nx
        End If
        mta_InfoColumn(mvar_NumColumns).TextWidth = nx
        AddColumn = mvar_NumColumns    '//Returnvalue = Index of Array
    Else
        lblColumn(mvar_NumColumns).Caption = ColCaption
        '//If found ....
        AddColumn = nFOund  '//Returnvalue = Index of Array
    End If
    
End Function

'//Preset Row
Public Function AddRow(RowID As Long, RowCaption As String) As Long
Dim n As Long, bFOund As Boolean, nFOund As Long, nx As Single
    For n = 1 To mvar_NumRows
        If mta_InfoRow(n).ID = RowID Then
            bFOund = True
            nFOund = n
            Exit For
        End If
    Next
    If bFOund = False Then
        mvar_NumRows = mvar_NumRows + 1
        ReDim Preserve mta_InfoRow(mvar_NumRows)

        
        ReDim Preserve ma_lRowEntryNumEntrys(mvar_NumRows)
        
        Load lblRowEntry(mvar_NumRows)
        lblRowEntry(mvar_NumRows).BackColor = m_RowHeadColor
        
        mta_InfoRow(mvar_NumRows).ID = RowID
        lblRowEntry(mvar_NumRows).Tag = RowID
        
        
        ma_lRowEntryNumEntrys(mvar_NumRows) = 1

        
        lblRowEntry(mvar_NumRows).Caption = RowCaption
        nx = UserControl.TextWidth(lblRowEntry(mvar_NumRows) & "    ")
        If lblRowEntry(mvar_NumRows).Width < nx Then lblRowEntry(mvar_NumRows).Width = nx
        mta_InfoRow(mvar_NumRows).TextWidth = nx
        
        AddRow = mvar_NumRows
    Else
        lblRowEntry(nFOund).Caption = RowCaption
        '//If found ....
        AddRow = nFOund
    End If

End Function

'//Blind entry of Data :
Public Function AddEntry(EntryID As Long, ColEntryID As Long, RowEntryID As Long, Optional EntryName As String, Optional ColEntryName As String, Optional RowEntryname As String) As Long
Dim bFOund As Boolean, n As Long, nFOund As Long, nx As Single
    bFOund = False
    '//check if already exists
    '//Entry
    For n = 1 To mvar_NumEntrys
        If mta_InfoEntry(n).ID = EntryID Then
            bFOund = True
            nFOund = n
            Exit For
        End If
    Next
    If bFOund = False Then
        mvar_NumEntrys = mvar_NumEntrys + 1   '//Redim arrays
        ReDim Preserve mta_InfoEntry(mvar_NumEntrys)
        
        
'        ReDim Preserve ma_lEntrysListWidth(mvar_NumEntrys)
        ReDim Preserve ma_bEntryHitInPaint(mvar_NumEntrys)
        Load lblEntry(mvar_NumEntrys)
        lblEntry(mvar_NumEntrys).BackColor = m_EntryColorMain
        '//Save Id of Entry and where it is located
        mta_InfoEntry(mvar_NumEntrys).ID = EntryID
        mta_InfoEntry(mvar_NumEntrys).RowID = RowEntryID
        mta_InfoEntry(mvar_NumEntrys).ColID = ColEntryID
        mta_InfoEntry(mvar_NumEntrys).sCaption = EntryName
        

        lblEntry(mvar_NumEntrys).Caption = EntryName
        lblEntry(mvar_NumEntrys).Tag = EntryID
        '//How wide is the Entry
        mta_InfoEntry(mvar_NumEntrys).ListWidth = UserControl.TextWidth(lblEntry(mvar_NumEntrys).Caption & "    ")
        mta_InfoEntry(mvar_NumEntrys).ListHeight = lblEntry(mvar_NumEntrys).Height
        lblEntry(mvar_NumEntrys).Width = mta_InfoEntry(mvar_NumEntrys).ListWidth
        mta_InfoEntry(mvar_NumEntrys).Height = lblEntry(mvar_NumEntrys).Height
        
    Else
        '//assign possibly new ID's
        mta_InfoEntry(nFOund).RowID = RowEntryID
        mta_InfoEntry(nFOund).ColID = ColEntryID

        If Len(EntryName) Then
            lblEntry(nFOund).Caption = EntryName
            mta_InfoEntry(nFOund).ListWidth = UserControl.TextWidth(lblEntry(nFOund).Caption & "    ")
            lblEntry(nFOund).Width = mta_InfoEntry(nFOund).ListWidth
        End If
        mta_InfoEntry(nFOund).Height = lblEntry(nFOund).Height
        mta_InfoEntry(mvar_NumEntrys).ListHeight = lblEntry(mvar_NumEntrys).Height
    End If
    

    '//ColEntry
    '//check if already exists
    bFOund = False
    For n = 1 To mvar_NumColumns
        If mta_InfoColumn(n).ID = ColEntryID Then
            bFOund = True
            nFOund = n
            Exit For
        End If
    Next
    
    If bFOund = False Then
        mvar_NumColumns = mvar_NumColumns + 1
        ReDim Preserve mta_InfoColumn(mvar_NumColumns)
        

        ReDim Preserve ma_lColEntryPosX(mvar_NumColumns)
        ReDim Preserve ma_lColEntryNumEntrys(mvar_NumColumns)
'        ReDim Preserve ma_lColumnWidth(mvar_NumColumns)
        
        Load lblColumn(mvar_NumColumns)
        lblColumn(mvar_NumColumns).BackColor = m_ColumnHeadColor
        
        mta_InfoColumn(mvar_NumColumns).ID = ColEntryID
        lblColumn(mvar_NumColumns).Tag = ColEntryID
        
        ma_lColEntryPosX(mvar_NumColumns) = 200

        ma_lColEntryNumEntrys(mvar_NumColumns) = 1
        
        If Len(ColEntryName) Then
            lblColumn(mvar_NumColumns).Caption = ColEntryName '& " C:" & ColEntryID
        End If
        nx = UserControl.TextWidth(lblColumn(mvar_NumColumns) & "    ")
        'If lblColumn(mvar_NumColumns).Width < nx Then lblColumn(mvar_NumColumns).Width = nx
        mta_InfoColumn(mvar_NumColumns).TextWidth = nx
    Else
        ma_lColEntryNumEntrys(nFOund) = ma_lColEntryNumEntrys(nFOund) + 1
        If Len(ColEntryName) Then
            lblColumn(nFOund).Caption = ColEntryName
            mta_InfoColumn(nFOund).TextWidth = UserControl.TextWidth(lblColumn(nFOund) & "    ")
        End If
        'mta_InfoColumn(nFOund).ActWidth = UserControl.TextWidth(lblColumn(nFOund) & "    ")
    End If
    
    '//RowEntry
    '//check if already exists
    bFOund = False
    For n = 1 To mvar_NumRows
        If mta_InfoRow(n).ID = RowEntryID Then
            bFOund = True
            nFOund = n
            Exit For
        End If
    Next
    If bFOund = False Then
        mvar_NumRows = mvar_NumRows + 1
        ReDim Preserve mta_InfoRow(mvar_NumRows)
        

        
        ReDim Preserve ma_lRowEntryNumEntrys(mvar_NumRows)
        
        Load lblRowEntry(mvar_NumRows)
        lblRowEntry(mvar_NumRows).BackColor = m_RowHeadColor
        
        mta_InfoRow(mvar_NumRows).ID = RowEntryID
        
        ma_lRowEntryNumEntrys(mvar_NumRows) = 1
        
        If Len(RowEntryname) Then
            lblRowEntry(mvar_NumRows).Caption = RowEntryname '& "R:" & RowEntryID
        End If
        lblRowEntry(mvar_NumRows).Tag = RowEntryID
        nx = UserControl.TextWidth(lblRowEntry(mvar_NumRows) & "    ")
        If lblRowEntry(mvar_NumRows).Width < nx Then lblRowEntry(mvar_NumRows).Width = nx
        mta_InfoRow(mvar_NumRows).TextWidth = nx
        
    Else
        '//Total number of Entries in Row
        ma_lRowEntryNumEntrys(nFOund) = ma_lRowEntryNumEntrys(nFOund) + 1
        If Len(RowEntryname) Then
            lblRowEntry(nFOund).Caption = RowEntryname ''& "R:" & RowEntryID
            nx = UserControl.TextWidth(lblRowEntry(nFOund) & "    ")
            If lblRowEntry(nFOund).Width < nx Then lblRowEntry(nFOund).Width = nx
            mta_InfoRow(nFOund).TextWidth = nx
        End If
        
    End If

End Function


Public Function PaintGrid()
Dim nCol As Long, nRow As Long, nH As Long, FoundCol As Long, nx As Long, bFoundInRow As Boolean
'Dim mtxEntryPosCnt() As Long
'Dim tpLine() As tpLineInfo
'Dim numLines As Long
'Dim actLine As Long
Dim LastX As Single, LastY As Single
Dim showWidth As Single, ShowHeight As Single
Dim xWidth As Single, xyz As Long
    LockWindowUpdate UserControl.hwnd
'//reset Values for Columns
    For nCol = 1 To mvar_NumColumns
        mta_InfoColumn(nCol).EntriesMaxWidth = 0
        mta_InfoColumn(nCol).NumEntries = 0
        ReDim mta_InfoColumn(nCol).EntriesID(0) '(mta_InfoColumn(nCol).NumEntries)
        ReDim mta_InfoColumn(nCol).EntriesNIdx(0)
        mta_InfoColumn(nCol).bHitInPaint = False
        lblColumn(nCol).Visible = False
    Next
    
    mvar_RowLabelMaxWidth = 0
    For nRow = 1 To mvar_NumRows
        mta_InfoRow(nRow).bHitInPaint = False
        mta_InfoRow(nRow).NumEntries = 0
        ReDim mta_InfoRow(nRow).EntriesID(0)
        ReDim mta_InfoRow(nRow).EntriesNIdx(0)
        '//Max width of Row "Container"
        If mvar_RowLabelMaxWidth < mta_InfoRow(nRow).TextWidth Then mvar_RowLabelMaxWidth = mta_InfoRow(nRow).TextWidth
        ReDim mta_InfoRow(nRow).LineInfo(0)
        mta_InfoRow(nRow).LineCount = 0
        lblRowEntry(nRow).Visible = False
    Next
'    numLines = 0
'    actLine = 0
    For nH = 1 To mvar_NumEntrys
        lblEntry(nH).Visible = False
        FoundCol = 0
        '//Determine column
        For nCol = 1 To mvar_NumColumns
            If mta_InfoEntry(nH).ColID = mta_InfoColumn(nCol).ID Then
                mta_InfoColumn(nCol).bHitInPaint = True
                '//Add Entry to Column
                mta_InfoColumn(nCol).NumEntries = mta_InfoColumn(nCol).NumEntries + 1
                ReDim Preserve mta_InfoColumn(nCol).EntriesID(mta_InfoColumn(nCol).NumEntries)
                ReDim Preserve mta_InfoColumn(nCol).EntriesNIdx(mta_InfoColumn(nCol).NumEntries)
                mta_InfoColumn(nCol).EntriesID(mta_InfoColumn(nCol).NumEntries) = mta_InfoEntry(nH).ID
                mta_InfoColumn(nCol).EntriesNIdx(mta_InfoColumn(nCol).NumEntries) = nH
                '//column is smaller than items
                If mta_InfoColumn(nCol).EntriesMaxWidth < mta_InfoEntry(nH).ListWidth Then
                    mta_InfoColumn(nCol).EntriesMaxWidth = mta_InfoEntry(nH).ListWidth
                End If
                If mta_InfoColumn(nCol).EntriesMaxWidth < mta_InfoColumn(nCol).TextWidth Then
                    mta_InfoColumn(nCol).EntriesMaxWidth = mta_InfoColumn(nCol).TextWidth
                End If
                FoundCol = nCol
                Exit For
            End If
        Next 'nCol

        '//determine Row
        For nRow = 1 To mvar_NumRows
            If mta_InfoEntry(nH).RowID = mta_InfoRow(nRow).ID Then
                mta_InfoRow(nRow).NumEntries = mta_InfoRow(nRow).NumEntries + 1
                ReDim Preserve mta_InfoRow(nRow).EntriesID(mta_InfoRow(nRow).NumEntries)
                ReDim Preserve mta_InfoRow(nRow).EntriesNIdx(mta_InfoRow(nRow).NumEntries)
                mta_InfoRow(nRow).EntriesID(mta_InfoRow(nRow).NumEntries) = mta_InfoEntry(nH).ID
                mta_InfoRow(nRow).EntriesNIdx(mta_InfoRow(nRow).NumEntries) = nH
                '//We hav a valid Columnentry
                If FoundCol Then
                    If mta_InfoRow(nRow).LineCount Then
                        '//Search for a free place
                        bFoundInRow = False
                        For nx = 1 To mta_InfoRow(nRow).LineCount
                            If mta_InfoRow(nRow).LineInfo(nx).ColEntry(FoundCol).colDataNIDx = 0 Then
                               mta_InfoRow(nRow).LineInfo(nx).ColEntry(FoundCol).colDataNIDx = nH
                               bFoundInRow = True
                               Exit For
                            End If
                        Next
                        '//No room found
                        If bFoundInRow = False Then
                            'Open next line
                            mta_InfoRow(nRow).LineCount = mta_InfoRow(nRow).LineCount + 1
                            ReDim Preserve mta_InfoRow(nRow).LineInfo(mta_InfoRow(nRow).LineCount)
                            ReDim mta_InfoRow(nRow).LineInfo(mta_InfoRow(nRow).LineCount).ColEntry(mvar_NumColumns)
                            mta_InfoRow(nRow).LineInfo(mta_InfoRow(nRow).LineCount).ColEntry(FoundCol).colDataNIDx = nH
                            mta_InfoRow(nRow).LineInfo(mta_InfoRow(nRow).LineCount).Height = mta_InfoEntry(nH).Height
                        End If
                    Else
                        mta_InfoRow(nRow).LineCount = 1
                        ReDim mta_InfoRow(nRow).LineInfo(1)
                        ReDim mta_InfoRow(nRow).LineInfo(1).ColEntry(mvar_NumColumns)
                        mta_InfoRow(nRow).LineInfo(1).ColEntry(FoundCol).colDataNIDx = nH
                        mta_InfoRow(nRow).LineInfo(1).Height = mta_InfoEntry(nH).Height
                    End If
                End If 'If FoundCol
                Exit For
            End If
        Next 'nRow
    Next 'nH
'//SO now we have the information where every entry is located
    

    '//Get biggest column
    xWidth = 0
    If m_AllColsHaveSameSize Then
        For nCol = 1 To mvar_NumColumns
            If xWidth < mta_InfoColumn(nCol).EntriesMaxWidth Then xWidth = mta_InfoColumn(nCol).EntriesMaxWidth
            If xWidth < mta_InfoColumn(nCol).TextWidth Then xWidth = mta_InfoColumn(nCol).TextWidth
        Next
    End If
    '//Get Positioning on X-axis
    LastX = mvar_OffsetColLeft
    For nCol = 1 To mvar_NumColumns
        If m_AllColsHaveSameSize Then
            mta_InfoColumn(nCol).EntriesMaxWidth = xWidth
        End If
        lblColumn(nCol).Left = LastX
        lblColumn(nCol).Width = mta_InfoColumn(nCol).EntriesMaxWidth
        
        If (m_ShowEmptyCols = True And mta_InfoColumn(nCol).NumEntries = 0) Or mta_InfoColumn(nCol).NumEntries > 0 Then
            lblColumn(nCol).Visible = True
            For nH = 1 To mta_InfoColumn(nCol).NumEntries
                mta_InfoEntry(mta_InfoColumn(nCol).EntriesNIdx(nH)).Xpos = LastX
                If mvarFillEntryToColumn = True Then
                    mta_InfoEntry(mta_InfoColumn(nCol).EntriesNIdx(nH)).Width = mta_InfoColumn(nCol).EntriesMaxWidth
                Else
                    If mvarCenterEntry Then
                        mta_InfoEntry(mta_InfoColumn(nCol).EntriesNIdx(nH)).Xpos = mta_InfoEntry(mta_InfoColumn(nCol).EntriesNIdx(nH)).Xpos + ((mta_InfoColumn(nCol).EntriesMaxWidth - mta_InfoEntry(mta_InfoColumn(nCol).EntriesNIdx(nH)).Width) / 2)
                    Else
                        mta_InfoEntry(mta_InfoColumn(nCol).EntriesNIdx(nH)).Width = mta_InfoEntry(mta_InfoColumn(nCol).EntriesNIdx(nH)).ListWidth
                    End If
                End If
            Next
        
            If m_AllColsHaveSameSize Then
                LastX = LastX + xWidth + mvar_OffsetColumns
            Else
                LastX = LastX + mta_InfoColumn(nCol).EntriesMaxWidth + mvar_OffsetColumns
            End If
        End If
    Next
    PicDataCol.Width = LastX
    PicDataEntry.Width = LastX
    
    LastY = mvar_OffsetRowFirst
    '//Get Positioning on Y-axis
    For nRow = 1 To mvar_NumRows
        lblRowEntry(nRow).Top = LastY
        
        Debug.Print "Row ", nRow, , "Count =", mta_InfoRow(nRow).LineCount
        'lblRowEntry(nRow).Top = mta_InfoColumn(nRow).Xpos
        If mta_InfoRow(nRow).LineCount > 0 Then
        
            For nx = 1 To mta_InfoRow(nRow).LineCount
                Debug.Print , "Line No ", nx
                For nCol = 1 To mvar_NumColumns
                    'Debug.Print mta_InfoRow(nRow).LineInfo(nx).ColEntry(nCol).colDataNIDx,
                    mta_InfoEntry(mta_InfoRow(nRow).LineInfo(nx).ColEntry(nCol).colDataNIDx).YPos = LastY
                Next
                Debug.Print
                LastY = LastY + mta_InfoRow(nRow).LineInfo(nx).Height + mvar_OffsetRows
            Next
            lblRowEntry(nRow).Height = LastY - lblRowEntry(nRow).Top
            lblRowEntry(nRow).Visible = True
        ElseIf (mta_InfoRow(nRow).LineCount = 0 And m_ShowEmptyRows = True) Then
            LastY = LastY + lblRowEntry(0).Height
            lblRowEntry(nRow).Height = LastY - lblRowEntry(nRow).Top
            lblRowEntry(nRow).Visible = True

        End If
    Next
    PicDataRow.Height = LastY
    PicDataEntry.Height = LastY
    
    For nH = 1 To mvar_NumEntrys
        If mvarCenterEntry Then
            lblEntry(nH).Alignment = vbCenter
        Else
            lblEntry(nH).Alignment = 0 ' vbAlignLeft
        End If
        lblEntry(nH).Left = mta_InfoEntry(nH).Xpos
        lblEntry(nH).Top = mta_InfoEntry(nH).YPos
        lblEntry(nH).Visible = True
    Next
    '//SizeGrid to content of Data and RowColInformation
    If mvarAutoSizeGrid Then
        UserControl.Height = LastY + lblHead.Height + PicOutCols.Top + PicOutCols.Height
        UserControl.Width = LastX + PicOutData.Left + PicOutData.Width
    End If
    UserControl_Resize
    LockWindowUpdate 0
End Function

'//Moved out of paint, because HTM-Out is not always needed
Public Function GetHTMLout() As String
Dim sLine As String, nCol As Long, nRow As Long, nH As Long, nx As Long
Dim CntCols As Long, ndx As Long
Dim sHTMOut As String
Dim ColorHead As String
Dim ColorColumn As String
Dim ColorRow As String
Dim ColorEntry As String
Dim sCenter
Dim ColMaxW As Single
    If mvarCenterEntry Then
        sCenter = " align ='center' "
    End If
    If m_HTMLuseColors Then
        If m_ColorHtmIsSameAsControl Then
            '//Convert colors from OLE(the VB-Controls color to HTML-Values
            ColorHead = " bgcolor='" & OleCOlToHTM(m_HeaderColor) & "' "
            ColorColumn = " bgcolor='" & OleCOlToHTM(m_ColumnHeadColor) & "' "
            ColorRow = " bgcolor='" & OleCOlToHTM(m_RowHeadColor) & "' "
            ColorEntry = " bgcolor='" & OleCOlToHTM(m_EntryColorMain) & "' "
        Else
            ColorHead = " bgcolor='" & m_ColorHTMHeader & "' "
            ColorColumn = " bgcolor='" & m_ColorHTMColumn & "' "
            ColorRow = " bgcolor='" & m_ColorHTMRow & "' "
            ColorEntry = " bgcolor='" & m_ColorHTMEntry & "' "
        End If
    End If

'//Preparation
    If mvar_IncludeBodyHTM Then
        sLine = "<HTML><HEAD><TITLE></TITLE></HEAD><BODY>"
    Else
        sLine = ""
    End If
    sLine = sLine & "<TABLE border=1>" & vbCrLf
    sHTMOut = sLine
'//Header
    CntCols = 0
    ColMaxW = 0
    For nCol = 1 To mvar_NumColumns
        If ColMaxW < mta_InfoColumn(nCol).EntriesMaxWidth Then ColMaxW = mta_InfoColumn(nCol).EntriesMaxWidth
        If mta_InfoColumn(nCol).NumEntries > 0 Or (mta_InfoColumn(nCol).NumEntries = 0 And m_ShowEmptyCols) Then
            CntCols = CntCols + 1
        End If
    Next
    '//Header
    sLine = "<TR><TD colspan=" & CntCols + 1 & " align='center'" & ColorHead & ">" & lblHead.Caption & "</TD></TR>" & vbCrLf
    '//Cell left top
    sLine = sLine & "<TR><TD>&nbsp;</TD>"
    '//The Columnheaders
    For nCol = 1 To mvar_NumColumns
        If mta_InfoColumn(nCol).NumEntries > 0 Or (mta_InfoColumn(nCol).NumEntries = 0 And m_ShowEmptyCols) Then
            sLine = sLine & "<TD align='center'" & ColorColumn & "width='" & ColMaxW & "'>" & lblColumn(nCol).Caption & "</TD>"
        End If
    Next
    sHTMOut = sHTMOut & sLine & vbCrLf
    
    For nRow = 1 To mvar_NumRows
        sLine = ""
        If mta_InfoRow(nRow).LineCount = 0 Then
            If m_ShowEmptyRows Then
                '//Rowheader
                sLine = sLine & "<TR><TD" & ColorRow & " width='" & lblRowEntry(nRow).Width / 2 & "'>" & lblRowEntry(nRow).Caption & "</TD>"
                '//And rest of line
                For nCol = 1 To mvar_NumColumns
                    If mta_InfoColumn(nCol).NumEntries Then
                        If m_HTMLshowEmptyCells Then
                            sLine = sLine & "<td>&nbsp;</td>"
                        Else
                             sLine = sLine & "<td visible='0'></td>"
                        End If
                    Else
                        If m_ShowEmptyCols Then
                            If m_HTMLshowEmptyCells Then
                                sLine = sLine & "<td>&nbsp;</td>"
                            Else
                                sLine = sLine & "<td visible='0'></td>"
                            End If
                        Else
                             'sLine = sLine & "<td visible='0'></td>"
                        End If
                    End If
                Next
'                If m_HTMLshowEmptyCells Then
'                    sLine = sLine & "<TD colspan=" & CntCols & " border=0 >&nbsp;</TD></TR>"
'                Else
'                    'sLine = sLine & "<TD colspan=" & CntCols & " visible='0' >&nbsp;</TD></TR>"
'                    For nCol = 1 To mvar_NumColumns
'                        If mta_InfoColumn(nCol).NumEntries Then
'                            sLine = sLine & "<td>&nbsp;</td>"
'                        Else
'                            If m_ShowEmptyCols Then
'                                sLine = sLine & "<td>&nbsp;</td>"
'                            Else
'                                sLine = sLine & "<td visible='0'></td>"
'                            End If
'                        End If
'                    Next
'
'                End If
            End If
        Else 'If mta_InfoRow(nRow).LineCount = 0 Then
            
            For nx = 1 To mta_InfoRow(nRow).LineCount
                If mta_InfoRow(nRow).LineCount = 1 Then
                    sLine = sLine & "<TR><TD" & ColorRow & " width='" & lblRowEntry(nRow).Width / 2 & "'>" & lblRowEntry(nRow).Caption & "</TD>"
                Else
                    If nx = 1 Then   '//First line
                        sLine = sLine & "<TR><TD rowspan=" & mta_InfoRow(nRow).LineCount & ColorRow & " width='" & lblRowEntry(nRow).Width / 2 & "' valign='top' >" & lblRowEntry(nRow).Caption & "</TD>"
                    Else    '//Others
                        '//No action needed on Rowspan and rowheader
                        If m_HTMLshowEmptyCells Then
                            'sLine = sLine & "<tr><td>&nbsp;</td>" ' "<tr><td>&nbsp;</td>"
                        Else
                            'sLine = sLine & "<tr><TD>&nbsp;</TD>" ' "<tr><TD visible=0></TD>"
                        End If
                    End If
                End If
                
                For nCol = 1 To mvar_NumColumns
                    '//There are some informations
                    If mta_InfoColumn(nCol).NumEntries Then
                            ndx = mta_InfoRow(nRow).LineInfo(nx).ColEntry(nCol).colDataNIDx
                            If ndx > 0 Then
                                sLine = sLine & "<TD" & ColorEntry & sCenter & ">" & lblEntry(ndx).Caption & "</TD>"
                            Else
                                If m_HTMLshowEmptyCells Then
                                    sLine = sLine & "<td>&nbsp;</td>" '"<tr><td>&nbsp;</td>"
                                Else
                                    sLine = sLine & "<TD visible=0></TD>"
                                End If
                            End If
                    Else
                        If m_ShowEmptyCols Then
                            ndx = mta_InfoRow(nRow).LineInfo(nx).ColEntry(nCol).colDataNIDx
                            If ndx > 0 Then
                                sLine = sLine & "<TD" & ColorEntry & sCenter & ">" & lblEntry(ndx).Caption & "</TD>"
                            Else
                                If m_HTMLshowEmptyCells Then
                                    sLine = sLine & "<td>&nbsp;</td>" '"<tr><td>&nbsp;</td>"
                                Else
                                    sLine = sLine & "<TD visible=0></TD>"
                                End If
                            End If
                        End If
                    End If

                Next 'For nCol = 1 To mvar_NumColumns
                sLine = sLine & "</tr>" & vbCrLf
                
            Next 'For nx = 1 To mta_InfoRow(nRow).LineCount
        End If 'If mta_InfoRow(nRow).LineCount = 0 Then
        sHTMOut = sHTMOut & sLine
    Next 'For nRow = 1 To mvar_NumRows
    
'//End
    sLine = "</TABLE>"
    If mvar_IncludeBodyHTM Then
        sLine = sLine & "</Body></HTML>"
    End If
    sHTMOut = sHTMOut & sLine
    mvar_sHTMLOUT = sHTMOut
    GetHTMLout = sHTMOut
End Function



'//HTML uses : Red, Green, blue as HEXvalues
'//OLECOlor : Blue Green Red as values
Public Function OleCOlToHTM(OleValue As Long) As String
Dim sOut As String, sHex As String, sHout As String
Dim hRed As String, hGreen As String, hBlue As String
Dim lRed As Long, lGreen As Long, lBlue As Long
Dim lTest As Long
    'sOut = "#" & MakeColnorWebSafe(OleValue) & "#"
    sHex = Hex(OleValue)
    '//Fill up to 6
    lRed = OleValue And 255 ' &HFF
    lGreen = OleValue And 65280 ' &HFF00&
    lBlue = OleValue And 16711680 '&HFF0000
    'If Len(sHex) <> 6 Then sHex = sHex & String(6 - Len(sHex), "0")
    hRed = Hex(lRed)
    hGreen = Hex(lGreen)
    hBlue = Hex(lBlue)
    sHout = Left(hRed, 2) & Left(hGreen, 2) & Left(hBlue, 2)
    sHout = LCase(sHout)
    'Debug.Print OleValue, sHex, sHout
    sOut = sHout ' MakeColorWebSafe(OleValue)
    
    OleCOlToHTM = "#" & sOut
'//Does not really work like desired ...
'    lTest = HexToLong(sOut)
'    Debug.Print sHex, OleValue, lTest, Hex(lTest), sHout
End Function

'//might be helpfull
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26213&lngWId=1

Private Function HexToLong(sHex As String) As Long
Dim bt() As Byte, lRet As Long, st() As String, n As Integer
Dim SW As String
    For n = 1 To 6 - Len(sHex)
        SW = SW & "0"
    Next
    SW = SW & sHex
    ReDim bt(2)
    ReDim st(2)
    st(0) = Left(SW, 2)
    st(1) = Mid(SW, 3, 2)
    st(2) = Right(SW, 2)
    For n = 0 To 2
        bt(n) = CByte("&H" & st(n))
    Next
    CopyMemory lRet, bt(0), 3
    HexToLong = lRet
End Function
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=60047&lngWId=1
Function MakeColorWebSafe(ByVal ColorValue As Long) As Long
    Dim bytData(2) As Byte, lonRet As Long
    If (ColorValue And &H80000000) Then
        lonRet = GetSysColor(ColorValue And &HFFFFFF)
    Else
        lonRet = ColorValue
    End If
    CopyMemory bytData(0), lonRet, 3
    MakeColorWebSafe = RGB(Round(bytData(0) / &H33) * &H33, Round(bytData(1) / &H33) * &H33, Round(bytData(2) / &H33) * &H33)
End Function

'//On Remove functions : Might be better to resize the arrays after find
Public Sub RemoveEntry(EntryID As Long)
Dim bFOund As Boolean, nFOund As Long, nH As Long
    For nH = 1 To mvar_NumEntrys
        If mta_InfoEntry(nH).ID = EntryID Then
            mta_InfoEntry(nH).ID = 0
            mta_InfoEntry(nH).ColID = 0
            mta_InfoEntry(nH).RowID = 0
        End If
    Next
End Sub
Public Sub RemoveRow(RowID As Long)
Dim bFOund As Boolean, nFOund As Long, nH As Long
    For nH = 1 To mvar_NumRows
        If mta_InfoRow(nH).ID = RowID Then
            mta_InfoRow(nH).ID = 0
        End If
    Next
End Sub
Public Sub RemoveColumn(ColumnID As Long)
Dim bFOund As Boolean, nFOund As Long, nH As Long
    For nH = 1 To mvar_NumColumns
        If mta_InfoColumn(nH).ID = ColumnID Then
            mta_InfoColumn(nH).ID = 0
        End If
    Next
End Sub
Public Sub ClearGrid()
Dim n As Long
    For n = 1 To mvar_NumEntrys
        Unload lblEntry(n)
    Next
    For n = 1 To mvar_NumColumns
        Unload lblColumn(n)
    Next
    For n = 1 To mvar_NumRows
        Unload lblRowEntry(n)
    Next

    mvar_NumRows = 0
    mvar_NumColumns = 0
    mvar_NumEntrys = 0
'//Might be slow but better is ... killing the array
    ReDim mta_InfoColumn(0)
    ReDim mta_InfoRow(0)
    ReDim mta_InfoEntry(0)
    PaintGrid
End Sub
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'
'
''MemberInfo=8,0,0,0
'Public Property Get ForeColor() As Long
'    ForeColor = m_ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As Long)
'    m_ForeColor = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property


'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property


'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property
'
'
''MemberInfo=7,0,0,0
'Public Property Get BackStyle() As Integer
'    BackStyle = m_BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    m_BackStyle = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property


'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Gibt den Rahmenstil für ein Objekt zurück oder legt diesen fest."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property


'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Erzwingt ein vollständiges Neuzeichnen eines Objekts."
    UserControl.Refresh
End Sub


'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Gibt eine Zugriffsnummer (von Microsoft Windows) auf ein Objektfenster zurück."
    hwnd = UserControl.hwnd
End Property


'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Gibt eine Zugriffsnummer (von Microsoft Windows) für den Gerätekontext des Objekts zurück."
    hdc = UserControl.hdc
End Property


'MappingInfo=lblHead,lblHead,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Gibt den Text zurück, der in der Titelleiste eines Objekts oder unter dem Symbol eines Objekts angezeigt wird, oder legt diesen fest."
    Caption = lblHead.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblHead.Caption() = New_Caption
    PropertyChanged "Caption"
End Property




Private Sub HScroll_Change()
    PicDataEntry.Move -HScroll.Value
    PicDataCol.Move -HScroll.Value
End Sub

Private Sub HScroll_Scroll()
    PicDataEntry.Move -HScroll.Value
    PicDataCol.Move -HScroll.Value
End Sub

Private Sub lblColumn_Click(Index As Integer)
    RaiseEvent ColumnSelected(lblColumn(Index).Tag)
End Sub

Private Sub lblColumn_DblClick(Index As Integer)
    RaiseEvent ColumnDblClick(lblColumn(Index).Tag)
End Sub

Private Sub lblColumn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent ColumnRightClick(lblColumn(Index).Tag)
End Sub

Private Sub lblEntry_Click(Index As Integer)
    GetInfoFromEntry lblEntry(Index).Tag, mvarSelectedEntry_Text, mvarSelectedEntry_RowID, mvarSelectedEntry_ColID
    RaiseEvent EntrySelected(lblEntry(Index).Tag)
End Sub

Private Sub lblEntry_DblClick(Index As Integer)
    GetInfoFromEntry lblEntry(Index).Tag, mvarSelectedEntry_Text, mvarSelectedEntry_RowID, mvarSelectedEntry_ColID
    RaiseEvent EntryDblClick(lblEntry(Index).Tag)
End Sub

Private Sub lblEntry_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent EntryRightClick(lblEntry(Index).Tag)
End Sub

Private Sub lblRowEntry_Click(Index As Integer)
    RaiseEvent RowSelected(lblRowEntry(Index).Tag)
End Sub

Private Sub lblRowEntry_DblClick(Index As Integer)
    RaiseEvent RowDblClick(lblRowEntry(Index).Tag)
End Sub

Private Sub lblRowEntry_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent RowRightClick(lblRowEntry(Index).Tag)
End Sub

Private Sub UserControl_InitProperties()
'    m_ForeColor = m_def_ForeColor
    UserControl.BackColor = m_def_BackColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
'    m_BackStyle = m_def_BackStyle
    
    mvarAutoSizeGrid = True
    m_ForeColor = m_def_ForeColor
    m_RowHeadColor = m_def_RowHeadColor
    m_ColumnHeadColor = m_def_ColumnHeadColor
    m_HeaderColor = m_def_HeaderColor
    m_EntryColorMain = m_def_EntryColorMain
    mvarCenterEntry = m_def_CenterEntry
    mvarFillEntryToColumn = m_def_FillEntryToColumn
    m_AllColsHaveSameSize = m_def_AllColsHaveSameSize
    m_ShowEmptyCols = m_def_ShowEmptyCols
    m_ShowEmptyRows = m_def_ShowEmptyRows
    lblColumn(0).Left = lblColumn(0).Width
    lblRowEntry(0).Top = lblRowEntry(0).Height
    m_ShowEmptyCols = m_def_ShowEmptyCols
    m_MinWidthRow = m_def_MinWidthRow
    m_MinWidthCol = m_def_MinWidthCol
    mvar_OffsetColLeft = m_def_OffsetColLeft
    mvar_OffsetRowFirst = m_def_OffsetRowFirst
    mvar_IncludeBodyHTM = m_def_IncludeBodyHTM
    m_ColorHTMHeader = m_def_ColorHTMHeader
    m_ColorHTMColumn = m_def_ColorHTMColumn
    m_ColorHTMRow = m_def_ColorHTMRow
    m_ColorHTMEntry = m_def_ColorHTMEntry
    m_ColorHtmIsSameAsControl = m_def_ColorHtmIsSameAsControl
    m_HTMLuseColors = m_def_HTMLuseColors
    m_HTMLshowEmptyCells = m_def_HTMLshowEmptyCells
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    lblHead.Caption = PropBag.ReadProperty("Caption", "Sectionname")
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    m_RowHeadColor = PropBag.ReadProperty("RowHeadColor", m_def_RowHeadColor)
    lblRowEntry(0).BackColor = m_RowHeadColor
    m_ColumnHeadColor = PropBag.ReadProperty("ColumnHeadColor", m_def_ColumnHeadColor)
    lblColumn(0).BackColor = m_ColumnHeadColor
    m_HeaderColor = PropBag.ReadProperty("HeaderColor", m_def_HeaderColor)
    lblHead.BackColor = m_HeaderColor
    m_EntryColorMain = PropBag.ReadProperty("EntryColorMain", m_def_EntryColorMain)
    lblEntry(0).BackColor = m_EntryColorMain
    mvarCenterEntry = PropBag.ReadProperty("CenterEntry", m_def_CenterEntry)
    mvarFillEntryToColumn = PropBag.ReadProperty("FillEntryToColumn", m_def_FillEntryToColumn)
    m_AllColsHaveSameSize = PropBag.ReadProperty("AllColsHaveSameSize", m_def_AllColsHaveSameSize)
    m_ShowEmptyCols = PropBag.ReadProperty("ShowEmptyCols", m_def_ShowEmptyCols)
    m_ShowEmptyRows = PropBag.ReadProperty("ShowEmptyRows", m_def_ShowEmptyRows)
    Set m_BackStyle = PropBag.ReadProperty("BackStyle", Nothing)
    m_ShowEmptyCols = PropBag.ReadProperty("ShowEmptyCols", m_def_ShowEmptyCols)
    m_MinWidthRow = PropBag.ReadProperty("MinWidthRow", m_def_MinWidthRow)
    m_MinWidthCol = PropBag.ReadProperty("MinWidthCol", m_def_MinWidthCol)
    mvar_OffsetColLeft = PropBag.ReadProperty("OffsetColLeft", m_def_OffsetColLeft)
    mvar_OffsetRowFirst = PropBag.ReadProperty("OffsetRowFirst", m_def_OffsetRowFirst)
    mvar_IncludeBodyHTM = PropBag.ReadProperty("IncludeBodyHTM", m_def_IncludeBodyHTM)
    m_ColorHTMHeader = PropBag.ReadProperty("ColorHTMHeader", m_def_ColorHTMHeader)
    m_ColorHTMColumn = PropBag.ReadProperty("ColorHTMColumn", m_def_ColorHTMColumn)
    m_ColorHTMRow = PropBag.ReadProperty("ColorHTMRow", m_def_ColorHTMRow)
    m_ColorHTMEntry = PropBag.ReadProperty("ColorHTMEntry", m_def_ColorHTMEntry)
    m_ColorHtmIsSameAsControl = PropBag.ReadProperty("ColorHtmIsSameAsControl", m_def_ColorHtmIsSameAsControl)
    m_HTMLuseColors = PropBag.ReadProperty("HTMLuseColors", m_def_HTMLuseColors)
    m_HTMLshowEmptyCells = PropBag.ReadProperty("HTMLshowEmptyCells", m_def_HTMLshowEmptyCells)
    
    PicDataCol.BorderStyle = 0
    PicDataRow.BorderStyle = 0
    PicDataEntry.BorderStyle = 0
    PicOutCols.BorderStyle = 0
    PicOutRow.BorderStyle = 0
    PicOutData.BorderStyle = 0
End Sub

Private Sub UserControl_Resize()
Dim x As Single, x2 As Single
Dim xLeft As Single, xTop As Single
Dim xWidth As Single, xHeight As Single

    lblHead.Move 0, 0, UserControl.ScaleWidth, lblHead.Height ' 255
    
    'PicOutRow.Width = x
    VScroll.Left = UserControl.ScaleWidth - VScroll.Width
    HScroll.Top = UserControl.ScaleHeight - HScroll.Height
    HScroll.Left = PicOutRow.Width
    VScroll.Top = PicOutCols.Top + PicOutCols.ScaleHeight
    
    HScroll.Width = UserControl.ScaleWidth - VScroll.Width - HScroll.Left
    VScroll.Height = UserControl.ScaleHeight - HScroll.Height - VScroll.Top
    VScroll.Visible = False
    HScroll.Visible = False
    
    xTop = PicOutCols.Top + PicOutCols.ScaleHeight
    
    '//Position datawindow in Control
    x = UserControl.ScaleWidth - PicOutRow.Width - PicOutRow.Left - VScroll.Width
    
    If x > 0 Then
        PicOutData.Width = x
        PicOutCols.Width = x
        HScroll.Width = x
    End If
    x = UserControl.ScaleHeight - PicOutCols.Top - PicOutCols.Height - HScroll.Height
    If x > 0 Then
        PicOutData.Height = x
        PicOutRow.Height = x
        VScroll.Height = x
    End If
    
    If PicDataEntry.Width >= PicOutData.Width Then
        HScroll.Visible = True
        x2 = PicDataEntry.ScaleWidth - PicOutData.ScaleWidth
        If x2 < 0 Then x2 = 1
        HScroll.Max = x2
        If x2 > 0 Then
            HScroll.LargeChange = x2 '/ 3
            HScroll.SmallChange = x2 / 20
        End If
    End If
    If PicDataEntry.Height >= PicOutData.Height Then
        VScroll.Visible = True
        x2 = PicDataEntry.ScaleHeight - PicOutData.ScaleHeight
        If x2 < 0 Then x2 = 1
        VScroll.Max = x2
        VScroll.LargeChange = x2 '* 0.75
        VScroll.SmallChange = x2 / 20
    End If
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
'    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", lblHead.Caption, "Sectionname")
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("RowHeadColor", m_RowHeadColor, m_def_RowHeadColor)
    Call PropBag.WriteProperty("ColumnHeadColor", m_ColumnHeadColor, m_def_ColumnHeadColor)
    Call PropBag.WriteProperty("HeaderColor", m_HeaderColor, m_def_HeaderColor)
    Call PropBag.WriteProperty("EntryColorMain", m_EntryColorMain, m_def_EntryColorMain)
    Call PropBag.WriteProperty("CenterEntry", mvarCenterEntry, m_def_CenterEntry)
    Call PropBag.WriteProperty("FillEntryToColumn", mvarFillEntryToColumn, m_def_FillEntryToColumn)
    Call PropBag.WriteProperty("AllColsHaveSameSize", m_AllColsHaveSameSize, m_def_AllColsHaveSameSize)
'    Call PropBag.WriteProperty("ShowEmptyCols", m_ShowEmptyCols, m_def_ShowEmptyCols)
    Call PropBag.WriteProperty("ShowEmptyRows", m_ShowEmptyRows, m_def_ShowEmptyRows)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, Nothing)
    Call PropBag.WriteProperty("ShowEmptyCols", m_ShowEmptyCols, m_def_ShowEmptyCols)
    Call PropBag.WriteProperty("MinWidthRow", m_MinWidthRow, m_def_MinWidthRow)
    Call PropBag.WriteProperty("MinWidthCol", m_MinWidthCol, m_def_MinWidthCol)
    Call PropBag.WriteProperty("OffsetColLeft", mvar_OffsetColLeft, m_def_OffsetColLeft)
    Call PropBag.WriteProperty("OffsetRowFirst", mvar_OffsetRowFirst, m_def_OffsetRowFirst)
    Call PropBag.WriteProperty("IncludeBodyHTM", mvar_IncludeBodyHTM, m_def_IncludeBodyHTM)
    Call PropBag.WriteProperty("ColorHTMHeader", m_ColorHTMHeader, m_def_ColorHTMHeader)
    Call PropBag.WriteProperty("ColorHTMColumn", m_ColorHTMColumn, m_def_ColorHTMColumn)
    Call PropBag.WriteProperty("ColorHTMRow", m_ColorHTMRow, m_def_ColorHTMRow)
    Call PropBag.WriteProperty("ColorHTMEntry", m_ColorHTMEntry, m_def_ColorHTMEntry)
    Call PropBag.WriteProperty("ColorHtmIsSameAsControl", m_ColorHtmIsSameAsControl, m_def_ColorHtmIsSameAsControl)
    Call PropBag.WriteProperty("HTMLuseColors", m_HTMLuseColors, m_def_HTMLuseColors)
    Call PropBag.WriteProperty("HTMLshowEmptyCells", m_HTMLshowEmptyCells, m_def_HTMLshowEmptyCells)
End Sub



'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property


'MemberInfo=10,0,0,0
Public Property Get RowHeadColor() As OLE_COLOR
    RowHeadColor = m_RowHeadColor
End Property

Public Property Let RowHeadColor(ByVal New_RowHeadColor As OLE_COLOR)
Dim n As Integer
    m_RowHeadColor = New_RowHeadColor
    PropertyChanged "RowHeadColor"
    For n = 0 To lblRowEntry.Count - 1
        lblRowEntry(n).BackColor = New_RowHeadColor
    Next
End Property


'MemberInfo=10,0,0,0
Public Property Get ColumnHeadColor() As OLE_COLOR
    ColumnHeadColor = m_ColumnHeadColor
End Property

Public Property Let ColumnHeadColor(ByVal New_ColumnHeadColor As OLE_COLOR)
Dim n As Integer
    m_ColumnHeadColor = New_ColumnHeadColor
    PropertyChanged "ColumnHeadColor"
    For n = 0 To lblColumn.Count - 1
        lblColumn(n).BackColor = New_ColumnHeadColor
    Next
End Property


'MemberInfo=10,0,0,0
Public Property Get HeaderColor() As OLE_COLOR
    HeaderColor = m_HeaderColor
End Property

Public Property Let HeaderColor(ByVal New_HeaderColor As OLE_COLOR)
    m_HeaderColor = New_HeaderColor
    lblHead.BackColor = New_HeaderColor
    PropertyChanged "HeaderColor"
End Property


'MemberInfo=10,0,0,0
Public Property Get EntryColorMain() As OLE_COLOR
    EntryColorMain = m_EntryColorMain
End Property

Public Property Let EntryColorMain(ByVal New_EntryColorMain As OLE_COLOR)
Dim n As Integer
    m_EntryColorMain = New_EntryColorMain
    PropertyChanged "EntryColorMain"
    For n = 0 To lblEntry.Count - 1
        lblEntry(n).BackColor = New_EntryColorMain
    Next
End Property


'MemberInfo=0,0,0,true0
Public Property Get CenterEntry() As Boolean
    CenterEntry = mvarCenterEntry
End Property

Public Property Let CenterEntry(ByVal New_CenterEntry As Boolean)
    mvarCenterEntry = New_CenterEntry
    PropertyChanged "CenterEntry"
End Property


'MemberInfo=0,0,0,true
Public Property Get FillEntryToColumn() As Boolean
    FillEntryToColumn = mvarFillEntryToColumn
End Property

Public Property Let FillEntryToColumn(ByVal New_FillEntryToColumn As Boolean)
    mvarFillEntryToColumn = mvarFillEntryToColumn
    PropertyChanged "FillEntryToColumn"
End Property


'MemberInfo=0,0,0,false
Public Property Get AllColsHaveSameSize() As Boolean
    AllColsHaveSameSize = m_AllColsHaveSameSize
End Property

Public Property Let AllColsHaveSameSize(ByVal New_AllColsHaveSameSize As Boolean)
    m_AllColsHaveSameSize = New_AllColsHaveSameSize
    PropertyChanged "AllColsHaveSameSize"
End Property
'
'
''MemberInfo=0,0,0,false
'Public Property Get ShowEmptyCols() As Boolean
'    ShowEmptyCols = m_ShowEmptyCols
'End Property
'
'Public Property Let ShowEmptyCols(ByVal New_ShowEmptyCols As Boolean)
'    m_ShowEmptyCols = New_ShowEmptyCols
'    PropertyChanged "ShowEmptyCols"
'End Property


'MemberInfo=0,0,0,false
Public Property Get ShowEmptyRows() As Boolean
    ShowEmptyRows = m_ShowEmptyRows
End Property

Public Property Let ShowEmptyRows(ByVal New_ShowEmptyRows As Boolean)
    m_ShowEmptyRows = New_ShowEmptyRows
    PropertyChanged "ShowEmptyRows"
End Property


'MemberInfo=17,0,0,0
Public Property Get BackStyle() As AmbientProperties
Attribute BackStyle.VB_Description = "Zeigt an, ob ein Bezeichnungsfeld oder der Hintergrund eines Figur-Steuerelements transparent oder undurchsichtig ist."
    Set BackStyle = m_BackStyle
End Property

Public Property Set BackStyle(ByVal New_BackStyle As AmbientProperties)
    Set m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Private Sub VScroll_Change()
    PicDataEntry.Move PicDataEntry.Left, -VScroll.Value
    PicDataRow.Move PicDataRow.Left, -VScroll.Value
End Sub

Private Sub VScroll_Scroll()
    PicDataEntry.Move PicDataEntry.Left, -VScroll.Value
    PicDataRow.Move PicDataRow.Left, -VScroll.Value
End Sub

'MemberInfo=0,0,0,false
Public Property Get ShowEmptyCols() As Boolean
    ShowEmptyCols = m_ShowEmptyCols
End Property

Public Property Let ShowEmptyCols(ByVal New_ShowEmptyCols As Boolean)
    m_ShowEmptyCols = New_ShowEmptyCols
    PropertyChanged "ShowEmptyCols"
End Property


'MemberInfo=12,0,0,1000
Public Property Get MinWidthRow() As Single
    MinWidthRow = m_MinWidthRow
End Property

Public Property Let MinWidthRow(ByVal New_MinWidthRow As Single)
    m_MinWidthRow = New_MinWidthRow
    PropertyChanged "MinWidthRow"
End Property


'MemberInfo=12,0,0,1000
Public Property Get MinWidthCol() As Single
    MinWidthCol = m_MinWidthCol
End Property

Public Property Let MinWidthCol(ByVal New_MinWidthCol As Single)
    m_MinWidthCol = New_MinWidthCol
    PropertyChanged "MinWidthCol"
End Property


'MemberInfo=12,0,0,0
Public Property Get OffsetColLeft() As Single
    OffsetColLeft = mvar_OffsetColLeft
End Property

Public Property Let OffsetColLeft(ByVal New_OffsetColLeft As Single)
    mvar_OffsetColLeft = New_OffsetColLeft
    PropertyChanged "OffsetColLeft"
End Property


'MemberInfo=12,0,0,0
Public Property Get OffsetRowFirst() As Single
    OffsetRowFirst = mvar_OffsetRowFirst
End Property

Public Property Let OffsetRowFirst(ByVal New_OffsetRowFirst As Single)
    mvar_OffsetRowFirst = New_OffsetRowFirst
    PropertyChanged "OffsetRowFirst"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,true
Public Property Get IncludeBodyHTM() As Boolean
    IncludeBodyHTM = mvar_IncludeBodyHTM
End Property

Public Property Let IncludeBodyHTM(ByVal New_IncludeBodyHTM As Boolean)
    mvar_IncludeBodyHTM = New_IncludeBodyHTM
    PropertyChanged "IncludeBodyHTM"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,#FFFFFF
Public Property Get ColorHTMHeader() As String
    ColorHTMHeader = m_ColorHTMHeader
End Property

Public Property Let ColorHTMHeader(ByVal New_ColorHTMHeader As String)
    m_ColorHTMHeader = New_ColorHTMHeader
    PropertyChanged "ColorHTMHeader"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,#FFFFFF
Public Property Get ColorHTMColumn() As String
    ColorHTMColumn = m_ColorHTMColumn
End Property

Public Property Let ColorHTMColumn(ByVal New_ColorHTMColumn As String)
    m_ColorHTMColumn = New_ColorHTMColumn
    PropertyChanged "ColorHTMColumn"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,#FFFFFF
Public Property Get ColorHTMRow() As String
    ColorHTMRow = m_ColorHTMRow
End Property

Public Property Let ColorHTMRow(ByVal New_ColorHTMRow As String)
    m_ColorHTMRow = New_ColorHTMRow
    PropertyChanged "ColorHTMRow"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,#FFFFFF
Public Property Get ColorHTMEntry() As String
    ColorHTMEntry = m_ColorHTMEntry
End Property

Public Property Let ColorHTMEntry(ByVal New_ColorHTMEntry As String)
    m_ColorHTMEntry = New_ColorHTMEntry
    PropertyChanged "ColorHTMEntry"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,false
Public Property Get ColorHtmIsSameAsControl() As Boolean
    ColorHtmIsSameAsControl = m_ColorHtmIsSameAsControl
End Property

Public Property Let ColorHtmIsSameAsControl(ByVal New_ColorHtmIsSameAsControl As Boolean)
    m_ColorHtmIsSameAsControl = New_ColorHtmIsSameAsControl
    PropertyChanged "ColorHtmIsSameAsControl"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,false
Public Property Get HTMLuseColors() As Boolean
    HTMLuseColors = m_HTMLuseColors
End Property

Public Property Let HTMLuseColors(ByVal New_HTMLuseColors As Boolean)
    m_HTMLuseColors = New_HTMLuseColors
    PropertyChanged "HTMLuseColors"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,false
Public Property Get HTMLshowEmptyCells() As Boolean
    HTMLshowEmptyCells = m_HTMLshowEmptyCells
End Property

Public Property Let HTMLshowEmptyCells(ByVal New_HTMLshowEmptyCells As Boolean)
    m_HTMLshowEmptyCells = New_HTMLshowEmptyCells
    PropertyChanged "HTMLshowEmptyCells"
End Property

