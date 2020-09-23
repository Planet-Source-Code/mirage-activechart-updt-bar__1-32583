VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "*\AYPChart.vbp"
Begin VB.Form frmTest 
   Caption         =   "ActiveChart Test"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin YP_ActiveChart.ActiveChart ActiveChart1 
      Height          =   6675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   11774
      uTopMargin      =   750
      uBottomMargin   =   1125
      uLeftMargin     =   825
      uRightMargin    =   825
      uContentBorder  =   -1  'True
      uSelectable     =   -1  'True
      uHotTracking    =   -1  'True
      uSelectedColumn =   -1
      uChartTitle     =   "ActiveChart"
      uChartSubTitle  =   "Created by Mirage-"
      uDisplayXAxis   =   -1  'True
      uDisplayYAxis   =   -1  'True
      uColorBars      =   0   'False
      uIntersectMajor =   10
      uIntersectMinor =   2
      uMaxYValue      =   100
      uDisplayDescript=   0   'False
      uXAxisLabel     =   "Test X Axis Label"
      uYAxislabel     =   "Test Y Axis Label"
      BackColor       =   -2147483643
      ForeColor       =   -2147483630
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   6690
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveChart1_ItemClick(cItem As YP_ActiveChart.ChartItem)
    grd.SelectionMode = flexSelectionByRow
    grd.Row = cItem.ItemID
    grd.ColSel = 2
End Sub

Private Sub Form_Load()
    Dim X As Integer, oChartItem As ChartItem
        
    Randomize
    grd.Rows = 1
    For X = 1 To 10
        oChartItem.ItemID = X
        oChartItem.SelectedDescription = "SelectedDescription " & X
        oChartItem.Value = CInt(Rnd * 100)
        oChartItem.XAxisDescription = "Item" & X
        ActiveChart1.AddItem oChartItem
            

        grd.AddItem X & vbTab & oChartItem.SelectedDescription & vbTab & oChartItem.Value
    Next X

    grd.FixedRows = 1
    grd.TextMatrix(0, 0) = "Item"
    grd.TextMatrix(0, 1) = "Description"
    grd.TextMatrix(0, 2) = "Value"

    grd.ColWidth(0) = 960
    grd.ColWidth(1) = Me.ScaleWidth - 960 - 2025
    grd.ColWidth(2) = 2025

End Sub

Private Sub Form_Resize()
    grd.Width = Me.ScaleWidth
    ActiveChart1.Width = Me.ScaleWidth

    grd.ColWidth(0) = 960
    grd.ColWidth(1) = Me.ScaleWidth - 960 - 2025
    grd.ColWidth(2) = 2025
End Sub

Private Sub grd_Click()
    DoEvents
    ActiveChart1.SelectedColumn = grd.Row - 1
End Sub

