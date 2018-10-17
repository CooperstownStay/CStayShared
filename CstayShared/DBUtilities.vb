Public Class DBUtilities
  Public Connection As System.Data.SqlClient.SqlConnection = Nothing
  Public Transaction As System.Data.SqlClient.SqlTransaction = Nothing
  Public ConnectionString As String = ""
  Public SelectedData As Object = Nothing
  Public CurrentRow As Object = Nothing
  Public CurrentRecordNumber As Integer = -1
  Public DatabaseToUse As String = ""
  Public Function FixParam(ByVal oIm As Object, ByVal bIsNullable As Boolean) As String
    If CNullS(oIm) = "" Then
      FixParam = ""
    ElseIf TypeOf oIm Is String Then
      FixParam = "'" & Trim(Replace(oIm, "'", "''")) & "'"
    ElseIf TypeOf oIm Is Boolean Then
      If CBool(oIm) Then
        FixParam = "True"
      Else
        FixParam = "False"
      End If
    Else
      FixParam = CNullS(oIm)
    End If
    If FixParam.Trim.ToString = "" Or FixParam.Trim.ToString.ToLower = "null" Then
      If bIsNullable Then
        FixParam = "Null"
      End If
    End If

  End Function
  Public Function SelectData( _
  ByVal sSelectStatement As String, _
  Optional ByVal sConnString As String = "", _
  Optional ByRef oInConn As System.Data.SqlClient.SqlConnection = Nothing, _
  Optional ByVal oTrans As System.Data.SqlClient.SqlTransaction = Nothing, _
  Optional ByRef bUseDataView As Boolean = True) As Object
    Dim oCmd As New System.Data.SqlClient.SqlCommand
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable, bConnectionOpened As Boolean = False
    SelectData = 0
    If sSelectStatement.ToString = "" Then Exit Function
    If Not SelectedData Is Nothing Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.Close()
      End If
      SelectedData = Nothing
    End If
    If oInConn Is Nothing Then
      OpenConnection(oInConn, Transaction, ConnectionString)
      bConnectionOpened = True
    End If
    oCmd = New System.Data.SqlClient.SqlCommand(sSelectStatement, oInConn)
    If Not (oTrans Is Nothing) Then
      oCmd.Transaction = oTrans
    End If
    If bUseDataView Then
      SelectedData = Nothing
      CurrentRow = Nothing
      oSQLAdapter.SelectCommand = oCmd
      oSQLAdapter.Fill(oDataTable)
      SelectedData = New System.Data.DataView(oDataTable)
    Else
      SelectedData = oCmd.ExecuteReader
    End If
    SelectData = SelectedData
    If (oTrans Is Nothing) And bConnectionOpened Then CloseConnection(oInConn, oTrans)
    If bUseDataView Then CurrentRecordNumber = -1 ' For DataView, need to start at -1 so initial Move puts us on first record
  End Function
  Public Function UpdateData( _
  ByVal sSQLStatement As String, _
  Optional ByVal sConnString As String = "", _
  Optional ByRef oInConn As System.Data.SqlClient.SqlConnection = Nothing, _
  Optional ByVal oTrans As System.Data.SqlClient.SqlTransaction = Nothing) As Object
    Dim oCmd As New System.Data.SqlClient.SqlCommand
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable
    UpdateData = 0
    If sSQLStatement.ToString = "" Then Exit Function
    If Not SelectedData Is Nothing Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.Close()
      End If
      SelectedData = Nothing
    End If
    If oInConn Is Nothing Then
      OpenConnection(Connection, Transaction, ConnectionString)
    Else
      Connection = oInConn
    End If
    oCmd = New System.Data.SqlClient.SqlCommand(sSQLStatement, Connection)
    If Not (oTrans Is Nothing) Then
      Transaction = oTrans
      oCmd.Transaction = Transaction
    End If
    UpdateData = oCmd.ExecuteNonQuery
    If Transaction Is Nothing Then CloseConnection(Connection, Transaction)
  End Function
  Public Function InsertData( _
  ByVal sSQLStatement As String, _
  Optional ByVal sConnString As String = "", _
  Optional ByRef oInConn As System.Data.SqlClient.SqlConnection = Nothing, _
  Optional ByVal oTrans As System.Data.SqlClient.SqlTransaction = Nothing) As Object
    Dim oCmd As New System.Data.SqlClient.SqlCommand
    Dim oSQLAdapter As New System.Data.SqlClient.SqlDataAdapter
    Dim oDataTable As New System.Data.DataTable
    InsertData = 0
    If sSQLStatement.ToString = "" Then Exit Function
    If Not SelectedData Is Nothing Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.Close()
      End If
      SelectedData = Nothing
    End If
    If oInConn Is Nothing Then
      OpenConnection(Connection, Transaction, ConnectionString)
    Else
      Connection = oInConn
    End If
    oCmd = New System.Data.SqlClient.SqlCommand(sSQLStatement, Connection)
    If Not (oTrans Is Nothing) Then
      Transaction = oTrans
      oCmd.Transaction = Transaction
    End If
    InsertData = oCmd.ExecuteNonQuery
    If Transaction Is Nothing Then CloseConnection(Connection, Transaction)
  End Function
  Public Function Move(Optional ByVal bCloseDataSourceAfterRead As Boolean = False, Optional ByVal sFilterForDataView As String = "", Optional ByVal iAmount As Integer = 1, Optional ByVal bMoveFirst As Boolean = False, Optional ByVal bMoveLast As Boolean = False) As Boolean
    Move = False
    If Not (SelectedData Is Nothing) Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        If iAmount > 0 Then
          Dim iCount As Integer = 0
          For iCount = 1 To iAmount
            If Not SelectedData.Read Then
              SelectedData.Close()
              SelectedData = Nothing
              Exit Function
            End If
          Next
        End If
        CurrentRow = SelectedData
      Else
        If sFilterForDataView.Trim = "" Then
          If bMoveFirst Then
            CurrentRecordNumber = 0
          ElseIf bMoveLast Then
            CurrentRecordNumber = SelectedData.Count - 1
          Else
            CurrentRecordNumber += iAmount
          End If
          If (SelectedData.Count <= CurrentRecordNumber) Or (CurrentRecordNumber < 0) Then
            Exit Function
          End If
        Else
          CurrentRecordNumber = 0
          SelectedData.RowFilter = sFilterForDataView.ToString
        End If
        CurrentRow = SelectedData.Item(CurrentRecordNumber)
      End If
      If bCloseDataSourceAfterRead Then
        If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
          SelectedData.Close()
        End If
        SelectedData = Nothing
      End If
      Move = True
    End If
  End Function
  Public Function MoveFirst(Optional ByVal bCloseDataSourceAfterRead As Boolean = False) As Boolean
    ' Move First should not be used with SQLDataReader
    MoveFirst = Move(bCloseDataSourceAfterRead, "", 0, True)
  End Function
  Public Function MovePrev(Optional ByVal bCloseDataSourceAfterRead As Boolean = False) As Boolean
    ' Move Prev should not be used with SQLDataReader
    MovePrev = Move(bCloseDataSourceAfterRead, "", -1)
  End Function
  Public Function MoveNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False) As Boolean
    MoveNext = Move(bCloseDataSourceAfterRead, "", 1)
  End Function
  Public Function MoveLast(Optional ByVal bCloseDataSourceAfterRead As Boolean = False) As Boolean
    ' Move Last should not be used with SQLDataReader
    MoveLast = Move(bCloseDataSourceAfterRead, "", 0, , True)
  End Function
  Public Function GetNext(Optional ByVal bCloseDataSourceAfterRead As Boolean = False) As Boolean
    ' This here for backward compatibility
    GetNext = Move(bCloseDataSourceAfterRead, "", 1)
  End Function
  Public Function IsActiveConnectionString(ByVal sConnString As String, Optional ByVal iTimeOut As Int16 = 5) As Boolean
    Dim oConn As New System.Data.SqlClient.SqlConnection
    IsActiveConnectionString = False
    On Error Resume Next
    oConn.ConnectionString = sConnString & ";Connect Timeout=" & iTimeOut & ";"
    oConn.Open()
    If Err.Number = 0 Then
      IsActiveConnectionString = True
    End If
    Err.Clear()
    oConn = Nothing
  End Function
  Public Function GetConnection(ByVal sConnnectionString As String, ByRef oConnection As System.Data.SqlClient.SqlConnection) As Boolean
    GetConnection = False
    If sConnnectionString.ToString = "" Then
      sConnnectionString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString").ToString
    End If
    If sConnnectionString.ToString = "" Then Exit Function
    If Not IsActiveConnectionString(sConnnectionString) Then Exit Function
    oConnection = New System.Data.SqlClient.SqlConnection
    oConnection.ConnectionString = sConnnectionString
    oConnection.Open()
    GetConnection = True
  End Function

  Public Function HasRows() As Boolean
    If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
      HasRows = SelectedData.HasRows
    Else
      HasRows = (SelectedData.Count > 0)
    End If
  End Function
  Public Sub OpenConnection()
    If Connection Is Nothing Then Connection = New System.Data.SqlClient.SqlConnection
    If Transaction Is Nothing Then
      If Connection.State = System.Data.ConnectionState.Closed Then
        If ConnectionString.ToString = "" Then ConnectionString = CreateConnectionStringFromConfig()
        Connection.ConnectionString = ConnectionString
        Connection.Open()
      End If
    Else
      Connection = Transaction.Connection
    End If
  End Sub
  Public Sub OpenConnection(ByRef MyConnection As System.Data.SqlClient.SqlConnection, ByRef MyTransaction As System.Data.SqlClient.SqlTransaction, ByRef MyConnectionString As String)
    If MyConnection Is Nothing Then MyConnection = New System.Data.SqlClient.SqlConnection
    If MyTransaction Is Nothing Then
      If MyConnection.State = System.Data.ConnectionState.Closed Then
        If MyConnectionString.ToString = "" Then MyConnectionString = CreateConnectionStringFromConfig()
        MyConnection.ConnectionString = MyConnectionString
        MyConnection.Open()
      End If
    Else
      MyConnection = MyTransaction.Connection
    End If
  End Sub
  Public Sub CloseConnection(ByRef MyConnection As System.Data.SqlClient.SqlConnection, ByRef MyTransaction As System.Data.SqlClient.SqlTransaction, Optional ByVal bKillTransaction As Boolean = False)
    If MyTransaction Is Nothing Or bKillTransaction Then
      If Not MyConnection Is Nothing Then
        If MyConnection.State <> System.Data.ConnectionState.Closed Then
          MyConnection.Close()
          System.Data.SqlClient.SqlConnection.ClearPool(MyConnection)
        End If
      End If
    End If
  End Sub
  Public Sub ProcessTransaction(ByRef MyConnection As System.Data.SqlClient.SqlConnection, ByRef MyTransaction As System.Data.SqlClient.SqlTransaction, Optional ByVal bCommit As Boolean = True)
    If Not (MyTransaction Is Nothing) Then
      If bCommit Then
        MyTransaction.Commit()
      Else
        MyTransaction.Rollback()
      End If
      CloseConnection(MyConnection, MyTransaction, True)
      System.Data.SqlClient.SqlConnection.ClearAllPools()
    End If
  End Sub

  Public Function CreateConnectionStringFromConfig(Optional ByVal sDatabase As String = "") As String

    Dim sDBEnv As String = "", sDBType As String = "", sConnStr As String = ""
    Dim sServer As String = "", sUsername As String = "", sPassword As String = ""

    sDBEnv = CNullS(System.Configuration.ConfigurationSettings.AppSettings("AppEnvironment"))
    sDBType = CNullS(System.Configuration.ConfigurationSettings.AppSettings("AppType"))

    If sDatabase = "" Then sDatabase = DatabaseToUse
    If sDatabase = "" Then sDatabase = "HouseRental"
    CreateConnectionStringFromConfig = CNullS(System.Configuration.ConfigurationSettings.AppSettings(sDBEnv & sDatabase & sDBType & "ConnectionString"))
    On Error Resume Next
    If CreateConnectionStringFromConfig.ToString = "" Then
      sServer = CNullS(System.Configuration.ConfigurationSettings.AppSettings(sDBEnv & "DBServer"))
      sUsername = CNullS(System.Configuration.ConfigurationSettings.AppSettings(sDBEnv & "DBUsername"))
      sPassword = CNullS(System.Configuration.ConfigurationSettings.AppSettings(sDBEnv & "DBPassword"))

      CreateConnectionStringFromConfig = CreateConnectionString(sServer, sDatabase & sDBType, sUsername, sPassword)
    End If
  End Function

  Public Function CreateConnectionString(ByVal sServer As String, ByVal sDatabase As String, ByVal sUsername As String, ByVal sPassword As String, Optional ByVal sAdditionalDBAttributes As String = "") As String

    CreateConnectionString = ""
    If sServer.ToString <> "" And sDatabase.ToString <> "" And sUsername.ToString <> "" Then
      CreateConnectionString = "data source=" & sServer & ";initial catalog=" & sDatabase & ";User ID=" & sUsername & ";Password=" & sPassword & ";persist security info=False;packet size=4096;" & sAdditionalDBAttributes
    End If

  End Function

  Public Sub New()

  End Sub

  Public Sub New(ByVal sDatabase As String)
    DatabaseToUse = sDatabase.ToString
  End Sub

  Protected Overrides Sub Finalize()
    If Not SelectedData Is Nothing Then
      If TypeOf SelectedData Is System.Data.SqlClient.SqlDataReader Then
        SelectedData.Close()
      End If
      SelectedData = Nothing
    End If
    MyBase.Finalize()
  End Sub
  Public Function CNullS(ByVal oInp As Object) As String
    CNullS = ""
    If Not IsDBNull(oInp) Then
      If Not (oInp Is Nothing) Then CNullS = CStr(oInp)
    End If
  End Function
  Public Function CNullI(ByVal oInp As Object) As Integer
    CNullI = 0
    If Not IsDBNull(oInp) Then
      If Not (oInp Is Nothing) Then
        If IsNumeric(oInp) Then CNullI = CInt(oInp)
      End If
    End If
  End Function
  Public Function CNullD(ByVal oInp As Object) As Double
    CNullD = 0
    If Not IsDBNull(oInp) Then
      If Not (oInp Is Nothing) Then
        If IsNumeric(oInp) Then CNullD = CDbl(oInp)
      End If
    End If
  End Function

End Class
