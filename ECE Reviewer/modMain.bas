Attribute VB_Name = "modMain"
'variable declarations
Public dbCN As ADODB.Connection
Public ExamType As Integer ' EXAM TYPES COM, MAT,ELEX
Public ExamQues As Integer ' QUESTIONS
Public Max As Integer ' MAXIMUM NUMBER OF QUESTIONS
Public QuesID(49) As Integer 'Variable used for storing Question IDs
Public Username As String

'Connect to database
Public Function DBConnect() As Boolean

On Error GoTo OpenErr

Dim MSDatabase

Set dbCN = New ADODB.Connection
MSDatabase = App.Path & "\" & "Questions.mdb"

    dbCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MSDatabase & ";Persist Security Info=False;Jet OLEDB:Database Password = pass"
    DBConnect = True
Exit Function

OpenErr:

    MsgBox "Error Opening " & MSDatabase & vbNewLine & Err.Description, vbCritical, "Open Database Error"
    DBConnect = False

End Function
