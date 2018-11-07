VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_message_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2018-09-02 Create by Tony Lin
Option Explicit

Public Sub err_message_log(ByVal msg As String)
    Dim rg_msg As Range
    Dim log_record_number As Long
    Dim i As Long
    
    log_record_number = err_msg.Range("B1").Value
    Set rg_msg = err_msg.Range("A4")
    
    i = 0
    While Trim(rg_msg.Offset(i).Value) <> ""
        i = i + 1
    Wend
    If i > log_record_number Then
        Call clear_msg
        i = 0
    End If
    rg_msg.Offset(i).Value = "[" & Now & "]" & msg
            
End Sub

Private Sub clear_msg()
    err_msg.Range("A4:A10000").Clear
End Sub

Public Function set_start_time() As String
    err_msg.Range("E2").Value = Now
End Function

Public Function check_infinity_loop() As String
    If Now > DateAdd("M", err_msg.Range("E1").Value, err_msg.Range("E2").Value) Then
        check_infinity_loop = 1
    Else
        check_infinity_loop = 0
    End If
End Function
