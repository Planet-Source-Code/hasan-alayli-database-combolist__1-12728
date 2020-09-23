VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   360
      TabIndex        =   10
      Top             =   5760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Database.frx":0000
            Key             =   "has"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Database.frx":7175
            Key             =   "hasan"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   5280
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   5280
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Database.frx":E96F
      Height          =   3735
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "EmployeeID"
         Caption         =   "EmployeeID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LastName"
         Caption         =   "LastName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FirstName"
         Caption         =   "FirstName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Title"
         Caption         =   "Title"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TitleOfCourtesy"
         Caption         =   "TitleOfCourtesy"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "BirthDate"
         Caption         =   "BirthDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "HireDate"
         Caption         =   "HireDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "City"
         Caption         =   "City"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Region"
         Caption         =   "Region"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "PostalCode"
         Caption         =   "PostalCode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Country"
         Caption         =   "Country"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "HomePhone"
         Caption         =   "HomePhone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Extension"
         Caption         =   "Extension"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Photo"
         Caption         =   "Photo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Notes"
         Caption         =   "Notes"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "ReportsTo"
         Caption         =   "ReportsTo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "PhotoPath"
         Caption         =   "PhotoPath"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add new"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "LastName"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "FirstName"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1680
      Top             =   4800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Northwind;Data Source=hasan"
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Northwind;Data Source=hasan"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employees"
      Caption         =   "Employees"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MoveCombo()
    ' In case of error, hide the ComboBox.
    On Error GoTo Error_Handler
    Dim gcol As MSDataGridLib.Column
    Set gcol = DataGrid1.Columns(DataGrid1.Col)
    
    If gcol.Caption = "FirstName" And DataGrid1.CurrentCellVisible Then
        ' Move the ComboBox inside the SupplierID column
        ' if it is the current column and it is visible.
        ImageCombo1.Move DataGrid1.Left + gcol.Left, _
            DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row), gcol.Width
        ImageCombo1.ZOrder
        ImageCombo1.SetFocus
        ImageCombo1.Text = gcol.Text
        Exit Sub
    End If
Error_Handler:
    ' In all other cases, hide the ComboBox.
   
    If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub imagecombo1_Click()
    ' Change the value of the underlying grid cell.
    DataGrid1.Columns("FirstName").Value = ImageCombo1.Text
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, _
    ByVal LastCol As Integer)
    MoveCombo
End Sub

Private Sub DataGrid1_RowResize(Cancel As Integer)
    MoveCombo
End Sub

Private Sub DataGrid1_ColResize(ByVal ColIndex As Integer, _
    Cancel As Integer)
    MoveCombo
End Sub

Private Sub DataGrid1_Scroll(Cancel As Integer)
    MoveCombo
End Sub

Private Sub DataGrid1_SplitChange()
    MoveCombo
End Sub



Private Sub Command1_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command2_Click()
DataGrid1.Columns(Combo1.Text).Text = Text3.Text
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete adAffectCurrent
End Sub



Private Sub Form_Load()
Set ImageCombo1.ImageList = ImageList1
ImageCombo1.ComboItems.Add , , "hasan", "has"
ImageCombo1.ComboItems.Add , , "said", "hasan"
ImageCombo1.ComboItems.Add , , "lina", "hasan"
ImageCombo1.ComboItems.Add , , "rana", "hasan"
DataGrid1.RowHeight = ImageCombo1.Height

x = Adodc1.Recordset.RecordCount
MsgBox x
Combo1.AddItem ("FirstName")
Combo1.AddItem ("LastName")







 
 
 
End Sub


