Attribute VB_Name = "Module1"
Public current_user As Integer
Public login As String
Public maturitydate As Date
Public StrSql As String
Public RS_Name As ADODB.Recordset
    Public strCnxn As String
   Public strSQLTitles As String
    
Public cnn As ADODB.Connection
Public RS_account As ADODB.Recordset
Public RS_atm As ADODB.Recordset
Public RS_deposit  As ADODB.Recordset
Public RS_Payment As ADODB.Recordset
Public RS_loan As ADODB.Recordset
Public RS_employee As ADODB.Recordset
Public RS_fixdep As ADODB.Recordset
Public RS_login As ADODB.Recordset
Public RS_transaction As ADODB.Recordset

Public welcometime As Integer
Public uname As String
Public Welcome As Boolean


Public Sub connect()
Set cnn = New ADODB.Connection
Set RS_account = New ADODB.Recordset
Set RS_atm = New ADODB.Recordset
Set RS_deposit = New ADODB.Recordset
Set RS_Payment = New ADODB.Recordset
Set RS_loan = New ADODB.Recordset
Set RS_employee = New ADODB.Recordset
Set RS_fixdep = New ADODB.Recordset
Set RS_login = New ADODB.Recordset
Set RS_transaction = New ADODB.Recordset
Set RS_Name = New ADODB.Recordset

cnn.CursorLocation = adUseClient
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\bank.mdb;Persist Security Info=False"

RS_account.Open "SELECT * FROM account", cnn, adOpenDynamic, adLockPessimistic
RS_atm.Open "SELECT * FROM atm_info", cnn, adOpenDynamic, adLockPessimistic
RS_deposit.Open "SELECT * FROM transact", cnn, adOpenDynamic, adLockPessimistic
RS_Payment.Open "SELECT * FROM payment", cnn, adOpenDynamic, adLockPessimistic
RS_loan.Open "SELECT * FROM loan", cnn, adOpenDynamic, adLockPessimistic
RS_employee.Open "SELECT * FROM employee", cnn, adOpenDynamic, adLockPessimistic
RS_fixdep.Open "SELECT * FROM fd", cnn, adOpenDynamic, adLockPessimistic
RS_login.Open "select * from login", cnn, adOpenDynamic, adLockPessimistic
RS_transaction.Open "select * FROM trans_log", cnn, adOpenDynamic, adLockPessimistic
RS_Name.Open "select * from account", cnn, adOpenDynamic, adLockPessimistic
 'strSQLTitles = "SELECT Name FROM account"
  '  rstName.Open strSQLTitles, cnn, adOpenDynamic, adLockPessimistic
End Sub

Public Sub trans()

End Sub
Public Sub close_connect()
RS_account.Close
RS_customer.Close
RS_deposit.Close
RS_Payment.Close
RS_loan.Close
RS_employee.Close
RS_fixdep.Close
RS_login.Close
RS_transaction.Close
RS_Name.Close
End Sub
