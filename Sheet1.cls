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

  Const sh As String = "�������}�N��" '���O�̏�������͂���V�[�g
  Const adr1 As String = "A1" '�ړ�������͂���Z��
  Const adr2 As String = "B1" '�A�Ԃ̊J�n�l����͂���Z��
  Const adr3 As String = "C1" '�A�Ԃ̏I���l����͂���Z��

  pre = Sheets(sh).Range(adr1).Value
  
  If IsNumeric(Sheets(sh).Range(adr2).Value) Then
    st = Sheets(sh).Range(adr2).Value
  Else
    MsgBox ("�J�n�ԍ������l����Ȃ���I���p�����ɒ����Ăˁ�")
    Exit Sub
  End If
  
  If IsNumeric(Sheets(sh).Range(adr3).Value) Then
    en = Sheets(sh).Range(adr3).Value
  Else
    MsgBox ("�I���ԍ������l����Ȃ���I���p�����ɒ����Ăˁ�")
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