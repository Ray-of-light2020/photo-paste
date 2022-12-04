Attribute VB_Name = "photo_paste"
Option Explicit

Sub photo_paste()
  
  Dim shape_name As String
  Dim WDT, HGT, TOP, LEFT As Long
  
  Application.ScreenUpdating = False
  shape_name = Selection.Name
  WDT = ActiveSheet.Shapes(shape_name).Width
  HGT = ActiveSheet.Shapes(shape_name).Height
  TOP = ActiveSheet.Shapes(shape_name).TOP
  LEFT = ActiveSheet.Shapes(shape_name).LEFT

        On Error GoTo Fin
            Application.Dialogs(xlDialogInsertPicture).Show
             With Selection.ShapeRange
                .LockAspectRatio = msoTrue
                .Width = WDT
                .Height = HGT
                .TOP = TOP
                .LEFT = LEFT
             End With
        
Fin:         On Error GoTo 0

Application.ScreenUpdating = True

End Sub
