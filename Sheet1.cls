VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Sugoi_Macro()

  Dim idx, st As Integer, en As Integer, pre As String

  Const sh As String = "すごいマクロ" '名前の書式を入力するシート
  Const adr1 As String = "A1" '接頭辞を入力するセル
  Const adr2 As String = "B1" '連番の開始値を入力するセル
  Const adr3 As String = "C1" '連番の終了値を入力するセル

  pre = Sheets(sh).Range(adr1).Value
  
  If IsNumeric(Sheets(sh).Range(adr2).Value) Then
    st = Sheets(sh).Range(adr2).Value
  Else
    MsgBox ("開始番号が数値じゃないよ！半角数字に直してね♪")
    Exit Sub
  End If
  
  If IsNumeric(Sheets(sh).Range(adr3).Value) Then
    en = Sheets(sh).Range(adr3).Value
  Else
    MsgBox ("終了番号が数値じゃないよ！半角数字に直してね♪")
    Exit Sub
  End If
  
  For idx = st To en
    If ActiveSheet.Index = Worksheets.Count Then
      Worksheets().Add After:=ActiveSheet
    Else
      ActiveSheet.Next.Activate
    End If
    ActiveSheet.Name = pre + CStr(idx)
  Next idx

End Sub
