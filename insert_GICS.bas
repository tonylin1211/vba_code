Attribute VB_Name = "insert_GICS"
Dim strSQL As String
Dim ma As String
Dim code_id As String
Dim sub_security_code As String
Dim cname As String

Dim rg_Code As Range
Dim rg_CName As Range
Dim rg_weight As Range
Dim rg_sector1 As Range
Dim rg_sector2 As Range
Dim rg_da As Range

Dim DBObject As cls_DBobject
Dim log_Object As cls_message_log

Dim i As Long
Dim j As Long

Sub insert_GICS_Sector_main()
On Error GoTo EH
    ma = "tw"
    
    Set DBObject = New cls_DBobject
    Set log_Object = New cls_message_log
    
    Call market_parameter_setup(ma)
    Call DBObject.Open_Conn(ma)
    Call err_module.set_start_time
    Call insert_GICS_Sector_Weight
    
    Call DBObject.Close_Conn

    Call MsgBox("done !")
    Exit Sub
EH:
    Call log_Object.err_message_log("insert_GICS_Sector_main" & ":" & Err.Description)
End Sub

Private Function insert_GICS_Sector_Weight() As String
    Dim cname As String
    Dim sector1 As String
    Dim sector2 As String
    Dim weight As String
    Dim da As String
    Dim strSQL As String
On Error GoTo EH
    i = 0
    j = 0
    da = gics_weight.Range("A1").Value
    
    strSQL = "SET search_path=daily"
    Call DBObject.exec_sql(strSQL)
    
    While rg_Code.Offset(i).Value <> ""
        code_id = rg_Code.Offset(i).Value
        sector1 = rg_sector1.Offset(i).Value
        sector2 = rg_sector2.Offset(i).Value
        cname = rg_CName.Offset(i).Value
        weight = rg_weight.Offset(i).Value
        strSQL = "update main_code set cname='" & cname & "'  where code='" & code_id & "';"
        Call DBObject.exec_sql(strSQL)
        
        strSQL = "insert into com_gics(code, da, weight, gics_sector1, gics_sector2) values('" & code_id & "','" & da & "', " & weight & ", '" & sector1 & "','" & sector2 & "');"
        Call DBObject.exec_sql(strSQL)
        i = i + 1
    Wend
    
    Exit Function
EH:
    Call log_Object.err_message_log("insert_GICS_Sector_Weight" & ":" & Err.Description)
End Function

Private Function insert_sql(ByVal strSQL As String) As String
On Error GoTo EH
EH:
    Call log_Object.err_message_log("insert_sql" & ":" & Err.Description)
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
    
    Set rg_Code = gics_weight.Range("C3")
    Set rg_weight = gics_weight.Range("B3")
    Set rg_sector1 = gics_weight.Range("E3")
    Set rg_sector2 = gics_weight.Range("F3")
    Set rg_CName = gics_weight.Range("D3")
    Exit Sub
EH:
    Call log_Object.err_message_log("market_parameter_setup" & ":" & Err.Description)
End Sub
