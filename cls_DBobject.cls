VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_DBobject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim DB_Conn As New ADODB.Connection
Dim Recordset As New ADODB.Recordset


Dim strSQL As String
Dim ma As String
Dim code_id As String
Dim sub_security_code As String
Dim cname As String

Dim log_Object As cls_message_log
''''''''''''''''
'  DB Action
''''''''''''''''
Public Function exec_sql(ByVal SQL As String) As String
On Error GoTo EH
    Call DB_Conn.Execute(SQL)
    exec_sql = 0
    Exit Function
EH:
    Call log_Object.err_message_log("exec_sql" & ":" & Err.Description)
    exec_sql = 1
End Function

Public Function select_sql(ByVal SQL As String) As ADODB.Recordset
On Error GoTo EH
    Set Recordset = DB_Conn.Execute(SQL)
    Set select_sql = Recordset
    Exit Function
EH:
    Call log_Object.err_message_log("select_sql" & ":" & Err.Description)
    'Call MsgBox("資料庫維護中，請稍候再試或聯絡負責人員")
End Function

Public Function Open_Conn(ByVal ma As String)
On Error GoTo EH
    Set log_Object = New cls_message_log
    DB_Conn.Open (ma)
    Exit Function
EH:
    Call log_Object.err_message_log("Open_Conn" & ":" & Err.Description)
End Function


Public Function Close_Conn()
On Error GoTo EH
    DB_Conn.Close
    Exit Function
EH:
    Call log_Object.err_message_log("Close_Conn" & ":" & Err.Description)
End Function
