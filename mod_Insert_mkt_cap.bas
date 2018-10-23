Attribute VB_Name = "mod_Insert_mkt_cap"
Dim strSQL As String
Dim ma As String
Dim code_id As String
Dim sub_security_code As String
Dim cname As String

'today data
Dim rg_Code As Range
Dim rg_mkt_cap As Range
Dim rg_da As Range

'history data
Dim rg_Code_hist As Range
Dim rg_mkt_cap_hist As Range
Dim rg_da_hist As Range

Dim DBObject As cls_DBobject
Dim log_Object As cls_message_log

Dim i As Long
Dim j As Long

Sub Insert_Today_Mkt_Cap()
On Error GoTo EH
    ma = "tw"
    
    Set DBObject = New cls_DBobject
    Set log_Object = New cls_message_log
    
    'setup paramter
    Call market_parameter_setup(ma)
    
    'open db connection
    Call DBObject.Open_Conn(ma)
    
    'start check unlimit loop timer
    Call log_Object.set_start_time
    
    'insert today data main flow
    Call insert_today_data_main_flow
    
    'close DB connection
    Call DBObject.Close_Conn

    Call MsgBox("done !")
    Exit Sub
EH:
    Call log_Object.err_message_log("Insert_Today_Mkt_Cap:" & Err.Description)
End Sub

Private Function insert_today_data_main_flow() As String
    Dim iRet As Integer
On Error GoTo EH
    Dim da As String
    i = 0
    
    ' X 軸
    da = rg_da.Value
    While Trim(rg_Code.Offset(0, i).Value) <> ""
        code_id = rg_Code.Offset(0, i).Value
        'check stock in correct market
        If InStr(1, code_id, sub_security_code) < 0 Then GoTo NEXT_Code
        Call check_main_code
        
        mkt_cap = rg_mkt_cap.Offset(0, i).Value
        Call update_mkt_cap_into_DB(da, mkt_cap)
        
NEXT_Code:
        i = i + 1
        If log_Object.check_infinity_loop = 1 Then
            iRet = MsgBox("執行時間過長，是否繼續 ?", vbYesNo, "Warning")
            If iRet = vbYes Then
                'reset check unlimit loop timer
                Call log_Object.set_start_time
                GoTo END_func
            End If
        End If
    Wend
    
END_func:
    insert_today_data_main_flow = 0
    Exit Function
EH:
    Call log_Object.err_message_log("insert_today_data_main_flow" & ":" & Err.Description)
    insert_today_data_main_flow = 1
End Function


Sub Insert_Hist_Mkt_Cap()
On Error GoTo EH
    ma = "tw"
    
    Set DBObject = New cls_DBobject
    Set log_Object = New cls_message_log
    
    'setup paramter
    Call market_parameter_setup(ma)
    
    'open db connection
    Call DBObject.Open_Conn(ma)
    
    'start check unlimit loop timer
    Call log_Object.set_start_time
    
    'insert today data main flow
    Call insert_today_data_main_flow
    
    'close DB connection
    Call DBObject.Close_Conn

    Call MsgBox("done !")
    Exit Sub
EH:
    Call log_Object.err_message_log("Insert_Hist_Mkt_Cap:" & Err.Description)
End Sub


Private Function insert_Hist_data_main_flow() As String
    Dim iRet As Integer
On Error GoTo EH
    Dim da As String
    i = 0
    
    ' X 軸
    While Trim(rg_Code.Offset(0, i).Value) <> ""
        code_id = rg_Code.Offset(0, i).Value
        'check stock in correct market
        If InStr(1, code_id, sub_security_code) < 0 Then GoTo NEXT_Code
        Call check_main_code
        Call DayBYDay_mkt_cap
        
NEXT_Code:
        i = i + 1
        If log_Object.check_infinity_loop = 1 Then
            iRet = MsgBox("執行時間過長，是否繼續 ?", vbYesNo, "Warning")
            If iRet = vbNo Then
                GoTo END_func
            Else
                'reset check unlimit loop timer
                Call log_Object.set_start_time
            End If
        End If
    Wend
    
END_func:
    insert_today_data_main_flow = 0
    Exit Function
EH:
    Call log_Object.err_message_log("insert_Hist_data_main_flow" & ":" & Err.Description)
    insert_today_data_main_flow = 1
End Function


Private Function DayBYDay_mkt_cap() As String
    Dim iRet As Integer
    Dim j As Long
    Dim da As String
    Dim mkt_cap As String
On Error GoTo EH
    j = 0
        
    ' Y 軸
    While Trim(rg_mkt_cap_hist.Offset(j).Value) <> ""
        da = rg_da.Offset(j, 0).Value
        mkt_cap = rg_mkt_cap_hist.Offset(j, i).Value
        Call update_mkt_cap_into_DB(da, mkt_cap)
        j = j + 1
        If log_Object.check_infinity_loop = 1 Then
            iRet = MsgBox("執行時間過長，是否繼續 ?", vbYesNo, "Warning")
            If iRet = vbNo Then
                GoTo END_func
            Else
                'reset check unlimit loop timer
                Call log_Object.set_start_time
            End If
        End If
    Wend
END_func:
    DayBYDay_price = 0
    Exit Function
EH:
    Call log_Object.err_message_log("DayBYDay_price" & ":" & Err.Description)
End Function


Public Function update_mkt_cap_into_DB(ByVal da As String, ByVal cap As String) As String
On Error GoTo EH
        
    strSQL = "update daily.price set market_cap = " + cap + " where da='" + da + "' and code='" + code_id + "';"
    result = DBObject.exec_sql(strSQL)
    
    Exit Function
EH:
    Call log_Object.err_message_log("Insert_price_into_DB" & ":" & Err.Description)
End Function

'check maincode table, if not EXIST then insert into code
Function check_main_code() As String
On Error GoTo EH
    Dim result As ADODB.Recordset
    
    strSQL = "select 1 from daily.main_code where code='" + code_id + "'"
    Set result = DBObject.select_sql(strSQL)
    If Not result.EOF Then
        result.Close
    Else
        result.Close
        strSQL = "Insert into daily.main_code (code, cname) values ('" + code_id + "', '" + cname + "');"
        Call DBObject.exec_sql(strSQL)
    End If
    Exit Function
EH:
    Call log_Object.err_message_log("check_main_code" & ":" & Err.Description)
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
        Call err_message_log("Error Market !")
    End If
    
    Set rg_Code = mkt_cap.Range("B4")
    Set rg_mkt_cap = mkt_cap.Range("B7")
    Set rg_da = mkt_cap.Range("A7")
    
    Set rg_Code_hist = mkt_cap.Range("C10")
    Set rg_mkt_cap_hist = mkt_cap.Range("B11")
    Set rg_da_hist = mkt_cap.Range("A11")
    
    Exit Sub
EH:
    Call log_Object.err_message_log("market_parameter_setup" & ":" & Err.Description)
End Sub
