VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

Dim armazena_horas(1 To 9) As Variant

Dim tempo As Date


Dim cell As Range

Sheets("Plan1").Activate

Range("D1").Select


For Each cell In ActiveCell.CurrentRegion

cell.Offset(0, 0).Select


tempo = cell.Value

cell.Offset(1, 0).Select
Application.OnTime Now + TimeValue(tempo), "msg_atencao"

Next cell


End Sub

