Attribute VB_Name = "mod_insert_price"

Dim strSQL As String
Dim ma As String
Dim code_id As String
Dim sub_security_code As String
Dim cname As String

'today data
Dim rg_Code As Range
Dim rg_price As Range
Dim rg_da As Range

'history data
Dim rg_Code_hist As Range
Dim rg_price_hist As Range
Dim rg_da_hist As Range


Dim DBObject As cls_DBobject
Dim log_Object As cls_message_log

Dim i As Long
Dim j As Long

Sub Insert_Today_Price()
On Error GoTo EH
    ma = "tw"
    
    Set DBObject = New cls_DBobject
    Set log_Object = New cls_message_log
    
    Call market_parameter_setup(ma)
    Call DBObject.Open_Conn(ma)
    Call log_Object.set_start_time
    Call insert_Today_main_flow
    
    Call DBObject.Close_Conn

    Call MsgBox("done !")
    Exit Sub
EH:
    Call log_Object.err_message_log("insert_Today_price" & ":" & Err.Description)
End Sub



Private Function insert_Today_main_flow() As String
    Dim iRet As Integer
    Dim da As String
    Dim price As String
On Error GoTo EH
    i = 0
    
    ' X 軸
    da = rg_da.Value
    While Trim(rg_Code.Offset(0, i).Value) <> ""
        code_id = rg_Code.Offset(0, i).Value
        'check stock in correct market
        If InStr(1, code_id, sub_security_code) < 0 Then GoTo NEXT_Code
        Call check_main_code
        
        price = rg_price.Offset(0, i).Value
        Call Insert_price_into_DB(da, price)
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
    insert_Today_main_flow = 0
    Exit Function
EH:
    Call log_Object.err_message_log("insert_Today_main_flow" & ":" & Err.Description)
    insert_Today_main_flow = 1
End Function

Sub insert_historial_price()
On Error GoTo EH
    ma = "tw"
    
    Set DBObject = New cls_DBobject
    Set log_Object = New cls_message_log
    
    Call market_parameter_setup(ma)
    Call DBObject.Open_Conn(ma)
    Call log_Object.set_start_time
    Call insert_Hist_main_flow
    
    Call DBObject.Close_Conn

    Call MsgBox("done !")
    Exit Sub
EH:
    Call log_Object.err_message_log("insert_historial_price" & ":" & Err.Description)
End Sub

Private Function insert_Hist_main_flow() As String
    Dim iRet As Integer
On Error GoTo EH
    i = 0
    
    ' X 軸
    While Trim(rg_Code.Offset(0, i).Value) <> ""
        code_id = rg_Code.Offset(0, i).Value
        If InStr(1, code_id, sub_security_code) < 0 Then GoTo NEXT_Code
        Call check_main_code
        Call DayBYDay_price
NEXT_Code:
        i = i + 1
        If log_Object.check_infinity_loop = 1 Then
            iRet = MsgBox("執行時間過長，是否繼續 ?", vbYesNo, "Warning")
            If iRet = vbYes Then GoTo END_func
        End If
    Wend
    
END_func:
    insert_Hist_main_flow = 0
    Exit Function
EH:
    Call log_Object.err_message_log("insert_Hist_main_flow" & ":" & Err.Description)
    insert_Hist_main_flow = 1
End Function

Private Function DayBYDay_price() As String
    
    Dim j As Long
    Dim da As String
    Dim pr As String
On Error GoTo EH
    j = 0
        
    ' Y 軸
    While Trim(rg_price_hist.Offset(j, i).Value) <> ""
        da = rg_da_hist.Offset(j, 0).Value
        pr = rg_price_hist.Offset(j, i).Value
        Call Insert_price_into_DB(da, pr)
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


Public Function Insert_price_into_DB(ByVal da As String, ByVal pr As String) As String
On Error GoTo EH
        
    strSQL = "Insert into daily.price(da, code, cl) values ('" + da + "', '" + code_id + "', " + Str(pr) + ");"
    result = DBObject.exec_sql(strSQL)
    
    Exit Function
EH:
    Call log_Object.err_message_log("Insert_price_into_DB" & ":" & Err.Description)
End Function


'檢查maincode 中是否已有，沒有的話要新增代碼
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
        Call log_Object.err_message_log("Error Market !")
    End If
    
    Set rg_Code = sh_price.Range("B3")
    Set rg_price = sh_price.Range("B4")
    Set rg_da = sh_price.Range("B2")
    
    Set rg_Code_hist = sh_price.Range("B10")
    Set rg_price_hist = sh_price.Range("B11")
    Set rg_da_hist = sh_price.Range("A11")
    
    Exit Sub
EH:
    Call log_Object.err_message_log("market_parameter_setup" & ":" & Err.Description)
End Sub
