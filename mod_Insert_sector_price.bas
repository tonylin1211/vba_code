Attribute VB_Name = "mod_Insert_sector_price"
Dim strSQL As String
Dim ma As String
Dim code_id As String
Dim sub_security_code As String
Dim cname As String

Dim rg_Code As Range
Dim rg_score As Range
Dim rg_da As Range

Dim DBObject As cls_DBobject
Dim log_Object As cls_message_log

Dim i As Long
Dim j As Long


Sub insert_gice_sector_score()
On Error GoTo EH
    ma = "tw"
    
    Set DBObject = New cls_DBobject
    Set log_Object = New cls_message_log
    
    Call market_parameter_setup(ma)
    Call DBObject.Open_Conn(ma)
    Call log_Object.set_start_time
    Call insert_main_flow
    
    Call DBObject.Close_Conn

    Call MsgBox("done !")
    Exit Sub
EH:
    Call log_Object.err_message_log("insert_mkt_cap:" & Err.Description)
End Sub

Private Function insert_main_flow() As String
    Dim iRet As Integer
On Error GoTo EH
    i = 0
    
    ' X 軸
    While Trim(rg_Code.Offset(i).Value) <> ""
        code_id = rg_Code.Offset(i).Value

        Call DayBYDay_sector_score
        
        i = i + 1
        If log_Object.check_infinity_loop = 1 Then
            iRet = MsgBox("執行時間過長，是否繼續 ?", vbYesNo, "Warning")
            If iRet = vbYes Then GoTo END_func
        End If
    Wend
    
END_func:
    insert_main_flow = 0
    Exit Function
EH:
    Call log_Object.err_message_log("insert_main_flow" & ":" & Err.Description)
    insert_main_flow = 1
End Function


Private Function DayBYDay_sector_score() As String
    Dim j As Long
    Dim da As String
    Dim score As String
On Error GoTo EH
    j = 0
        
    ' Y 軸
    While Trim(rg_da.Offset(0, j).Value) <> ""
        da = rg_da.Offset(0, j).Value
        score = rg_score.Offset(i, j).Value
        Call update_mkt_cap_into_DB(da, score)
        j = j + 1
        If log_Object.check_infinity_loop = 1 Then
            iRet = MsgBox("執行時間過長，是否繼續 ?", vbYesNo, "Warning")
            If iRet = vbYes Then GoTo END_func
        End If
    Wend
END_func:
    DayBYDay_price = 0
    Exit Function
EH:
    Call log_Object.err_message_log("DayBYDay_price" & ":" & Err.Description)
End Function

Public Function update_mkt_cap_into_DB(ByVal da As String, ByVal score As String) As String
On Error GoTo EH
        
    strSQL = "insert into daily.gics_score(id, da, score) select id, '" + da + "', " + score + " from daily.gics_main where ename='" + code_id + "';"
    result = DBObject.exec_sql(strSQL)
    
    Exit Function
EH:
    Call log_Object.err_message_log("Insert_price_into_DB" & ":" & Err.Description)
End Function



Sub market_parameter_setup(ByVal market As String)
On Error GoTo EH
    If market = "tw" Then
        sub_security_code = " TT Equity"
    ElseIf market = "jp" Then
        sub_security_code = " JP Equity"
    ElseIf market = "sp500" Then
        sub_security_code = " TT Equity"
    ElseIf market = "cn" Then
        sub_security_code = " CH Equity"
    ElseIf market = "hk" Then
        sub_security_code = " HK Equity"
    Else
        Call cls_message_log.err_message_log("Error Market !")
    End If
    
    Set rg_Code = gics.Range("A2")
    Set rg_score = gics.Range("B2")
    Set rg_da = gics.Range("B1")
    
    Exit Sub
EH:
    Call log_Object.err_message_log("market_parameter_setup" & ":" & Err.Description)
End Sub

