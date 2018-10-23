Attribute VB_Name = "insert_price"

Dim strSQL As String
Dim ma As String
Dim code_id As String
Dim sub_security_code As String
Dim cname As String

Dim rg_Code As Range
Dim rg_price As Range
Dim rg_da As Range

Dim DBObject As cls_DBobject
Dim log_Object As cls_message_log

Dim i As Long
Dim j As Long

Sub insert_historial_price()
On Error GoTo EH
    ma = "tw"
    
    Set DBObject = New cls_DBobject
    Set log_Object = New cls_message_log
    
    Call market_parameter_setup(ma)
    Call DBObject.Open_Conn(ma)
    Call err_module.set_start_time
    Call insert_main_flow
    
    Call DBObject.Close_Conn

    Call MsgBox("done !")
    Exit Sub
EH:
    Call log_Object.err_message_log("insert_historial_price" & ":" & Err.Description)
End Sub

Private Function insert_main_flow() As String
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
        If err_module.check_infinity_loop = 1 Then
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

Private Function DayBYDay_price() As String
    
    Dim j As Long
    Dim da As String
    Dim pr As String
On Error GoTo EH
    j = 0
        
    ' Y 軸
    While Trim(rg_price.Offset(j).Value) <> ""
        da = rg_da.Offset(j, 0).Value
        pr = rg_price.Offset(j, i).Value
        Call Insert_price_into_DB(da, pr)
        j = j + 1
        If err_module.check_infinity_loop = 1 Then
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
        Call err_message_log("Error Market !")
    End If
    
    Set rg_Code = tw_price.Range("B1")
    Set rg_price = tw_price.Range("B4")
    Set rg_da = tw_price.Range("A4")
    
    Exit Sub
EH:
    Call log_Object.err_message_log("market_parameter_setup" & ":" & Err.Description)
End Sub

'
'''''''''''''''''
''  DB Action
'''''''''''''''''
'Public Function exec_sql(ByVal SQL As String) As String
'On Error GoTo EH
'    Call DB_Conn.Execute(SQL)
'    exec_sql = 0
'    Exit Function
'EH:
'    Call err_message_log("exec_sql" & ":" & Err.Description)
'    exec_sql = 1
'End Function
'
'Public Function select_sql(ByVal SQL As String) As ADODB.Recordset
'On Error GoTo EH
'    Set Recordset = DB_Conn.Execute(SQL)
'    Set select_sql = Recordset
'    Exit Function
'EH:
'    Call err_message_log("select_sql" & ":" & Err.Description)
'End Function
'
'Public Function Open_Conn(ByVal ma As String)
'On Error GoTo EH
'    DB_Conn.Open (ma)
'    Exit Function
'EH:
'    Call err_message_log("Open_Conn" & ":" & Err.Description)
'End Function
'
'
'Public Function Close_Conn()
'On Error GoTo EH
'    DB_Conn.Close
'    Exit Function
'EH:
'    Call err_message_log("Close_Conn" & ":" & Err.Description)
'End Function


