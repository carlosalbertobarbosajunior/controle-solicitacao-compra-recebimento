VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim xPTable As PivotTable
    Dim xPFile As PivotField
    Dim xStr As String
    On Error GoTo ErrMsg
    If Intersect(Target, Range("B3:D4")) Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Set xPTable = Worksheets("Interface").PivotTables("Tabela din?mica1")
    Set xPFile = xPTable.PivotFields("Servi?o")
    xStr = Target.Text
    xPFile.ClearAllFilters
    xPFile.CurrentPage = xStr
    Application.ScreenUpdating = True
    Range("B3").Select
Exit Sub

ErrMsg:
    MsgBox ("Servi?o n?o existente. O filtro ser? resetado.")
    Range("B3").Select
    
End Sub
