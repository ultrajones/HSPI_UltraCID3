Imports System.Data
Imports System.Data.Common
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Data.SQLite

Module hspi_database

  Public DBConnectionMain As SQLite.SQLiteConnection  ' Our main database connection
  Public DBConnectionTemp As SQLite.SQLiteConnection  ' Our temp database connection

  Public gDBInsertSuccess As ULong = 0            ' Tracks DB insert success
  Public gDBInsertFailure As ULong = 0            ' Tracks DB insert success

  Public bDBInitialized As Boolean = False        ' Indicates if database successfully initialized

  Public SyncLockMain As New Object
  Public SyncLockTemp As New Object

#Region "Database Initilization"

  ''' <summary>
  ''' Initializes the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitializeMainDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered InitializeMainDatabase() function.", MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionMain Is Nothing Then
        If CloseDBConn(DBConnectionMain) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Create the database directory if it does not exist
      '
      Dim databaseDir As String = FixPath(String.Format("{0}\Data\{1}\", hs.GetAppPath, IFACE_NAME.ToLower))
      If Directory.Exists(databaseDir) = False Then
        Directory.CreateDirectory(databaseDir)
      End If

      '
      ' Determine the database filename
      '
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strDbProviderFactory As String = "System.Data.SQLite"
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionMain, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeDatabase()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Initializes the temporary database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitializeTempDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    strMessage = "Entered InitializeChannelDatabase() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionTemp Is Nothing Then
        If CloseDBConn(DBConnectionTemp) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Determine the database filename
      '
      Dim dtNow As DateTime = DateTime.Now
      Dim iHour As Integer = dtNow.Hour
      Dim strDBDate As String = iHour.ToString.PadLeft(2, "0")
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}_{2}.db3", hs.GetAppPath(), IFACE_NAME.ToLower, strDBDate))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3; Journal Mode=Off;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionTemp, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeTempDatabase()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Opens a connection to the database
  ''' </summary>
  ''' <param name="objConn"></param>
  ''' <param name="strConnectionString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function OpenDBConn(ByRef objConn As SQLite.SQLiteConnection,
                              ByVal strConnectionString As String) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered OpenDBConn() function.", MessageType.Debug)

    Try
      '
      ' Open database connection
      '
      objConn = New SQLite.SQLiteConnection()
      objConn.ConnectionString = strConnectionString
      objConn.Open()

      '
      ' Run database vacuum
      '
      WriteMessage("Running SQLite database vacuum.", MessageType.Debug)
      Using MyDbCommand As DbCommand = objConn.CreateCommand()

        MyDbCommand.Connection = objConn
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = "VACUUM"
        MyDbCommand.ExecuteNonQuery()

        MyDbCommand.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "OpenDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = objConn.State = ConnectionState.Open

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database initialization complete."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Database initialization failed using [" & strConnectionString & "]."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Closes database connection
  ''' </summary>
  ''' <param name="objConn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CloseDBConn(ByRef objConn As SQLite.SQLiteConnection) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered CloseDBConn() function.", MessageType.Debug)

    Try
      '
      ' Attempt to the database
      '
      If objConn.State <> ConnectionState.Closed Then
        objConn.Close()
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CloseDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = objConn.State = ConnectionState.Closed

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database connection closed successfuly."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Unable to close database; Try restarting HomeSeer."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Checks to ensure a table exists
  ''' </summary>
  ''' <param name="strTableName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckDatabaseTable(ByVal strTableName As String) As Boolean

    Dim strMessage As String = ""
    Dim bSuccess As Boolean = False

    Try
      '
      ' Build SQL delete statement
      '
      If Regex.IsMatch(strTableName, "tblCallerLog") = True Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue

            sqlQueue.Enqueue("CREATE TABLE tblCallerLog(" _
                            & "id INTEGER PRIMARY KEY," _
                            & "ts integer," _
                            & "nmbr varchar(15) NOT NULL," _
                            & "name varchar(15) NOT NULL" _
                          & ")")

            sqlQueue.Enqueue("CREATE INDEX idxTS   ON tblCallerLog (ts)")
            sqlQueue.Enqueue("CREATE INDEX idxNAME ON tblCallerLog (name)")
            sqlQueue.Enqueue("CREATE INDEX idxNMBR ON tblCallerLog (nmbr)")

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      ElseIf Regex.IsMatch(strTableName, "tblCallerDetails") = True Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue
            sqlQueue.Enqueue("CREATE TABLE tblCallerDetails(" _
                            & "id INTEGER PRIMARY KEY," _
                            & "nmbr varchar(15) NOT NULL," _
                            & "name varchar(25) NOT NULL," _
                            & "attr INTEGER," _
                            & "notes varchar(255), " _
                            & "last_ts integer," _
                            & "call_count INTEGER" _
                          & ")")

            sqlQueue.Enqueue("CREATE UNIQUE INDEX idxNUMR2 ON tblCallerDetails (nmbr)")
            sqlQueue.Enqueue("CREATE INDEX idxNAME2 ON tblCallerDetails (name)")
            sqlQueue.Enqueue("CREATE INDEX idxATTR2 ON tblCallerDetails (attr)")

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      Else
        Throw New Exception(strTableName & " not currently supported.")
      End If

      bSuccess = True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDatabaseTables()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Returns the size of the selected database
  ''' </summary>
  ''' <param name="databaseName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDatabaseSize(ByVal databaseName As String)

    Try

      Select Case databaseName
        Case "DBConnectionMain"
          '
          ' Determine the database filename
          '
          Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))
          Dim file As New FileInfo(strDataSource)
          Return FormatFileSize(file.Length)

      End Select

    Catch pEx As Exception

    End Try

    Return "Unknown"

  End Function

  ''' <summary>
  ''' Converts filesize to String
  ''' </summary>
  ''' <param name="FileSizeBytes"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function FormatFileSize(ByVal FileSizeBytes As Long) As String

    Try

      Dim sizeTypes() As String = {"B", "KB", "MB", "GB"}
      Dim Len As Decimal = FileSizeBytes
      Dim sizeType As Integer = 0

      Do While Len > 1024
        Len = Decimal.Round(Len / 1024, 2)
        sizeType += 1
        If sizeType >= sizeTypes.Length - 1 Then Exit Do
      Loop

      Dim fileSize As String = String.Format("{0} {1}", Len.ToString("N0"), sizeTypes(sizeType))
      Return fileSize

    Catch pEx As Exception

    End Try

    Return FileSizeBytes.ToString

  End Function

#End Region

#Region "Database Queries"

  ''' <summary>
  ''' Return values from database
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function QueryDatabase(ByVal strSQL As String) As DataSet

    Dim strMessage As String = ""

    WriteMessage("Entered QueryDatabase() function.", MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      ' Initialize the command object
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        '
        ' Initialize the dataset, then populate it
        '
        Dim MyDS As DataSet = New DataSet

        Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
        MyDA.SelectCommand = MyDbCommand

        SyncLock SyncLockMain
          MyDA.Fill(MyDS)
        End SyncLock

        MyDbCommand.Dispose()

        Return MyDS

      End Using

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "QueryDatabase()")
      Return New DataSet
    End Try

  End Function

  ''' <summary>
  ''' Insert data into database
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <remarks></remarks>
  Public Sub InsertData(ByVal strSQL As String)

    Dim strMessage As String = ""
    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    Select Case DBConnectionTemp.State
      Case ConnectionState.Broken, ConnectionState.Closed
        Exit Sub
    End Select

    Try

      Using dbcmd As DbCommand = DBConnectionTemp.CreateCommand()

        dbcmd.Connection = DBConnectionTemp
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockTemp
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "InsertData() Reports Error: [" & pEx.ToString & "], " _
                  & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strSQL, MessageType.Debug)
    Finally
      '
      ' Update counter
      '
      If iRecordsAffected = 1 Then
        gDBInsertSuccess += 1
      Else
        gDBInsertFailure += 1
      End If
    End Try

  End Sub

#End Region

#Region "Caller Logging"

  ''' <summary>
  ''' Process incoming call
  ''' </summary>
  ''' <param name="ts"></param>
  ''' <param name="strNmbr"></param>
  ''' <param name="strName"></param>
  ''' <param name="iCallerAttr"></param>
  ''' <remarks></remarks>
  Public Sub ProcessIncomingCall(ByVal ts As Integer, _
                                 ByVal strNmbr As String, _
                                 ByVal strName As String, _
                                 ByVal iCallerAttr As Integer)

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Exit Sub

    Try
      '
      ' Ensure our data doesn't contain single quotes
      '
      strNmbr = strNmbr.Trim.Replace("'", "")
      strName = strName.Trim.Replace("'", "")

      '
      ' Insert new incoming caller log data
      '
      hspi_database.InsertCallerLogData(ts, strNmbr, strName)

      If iCallerAttr = -1 Then
        '
        ' Insert new caller detail data
        '
        hspi_database.InsertCallerDetailData(ts, strNmbr, strName)
      Else
        '
        ' Update existing caller detail data
        '
        hspi_database.UpdateCallerDetailData(ts, strNmbr, strName)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ProcessIncomingCall()")
    End Try

  End Sub

  ''' <summary>
  ''' Insert caller log data into database
  ''' </summary>
  ''' <param name="ts"></param>
  ''' <param name="strNmbr"></param>
  ''' <param name="strName"></param>
  ''' <remarks></remarks>
  Public Sub InsertCallerLogData(ByVal ts As Integer, _
                                 ByVal strNmbr As String, _
                                 ByVal strName As String)

    Dim strMessage As String = ""
    Dim strSQL As String = ""
    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Exit Sub

    Try

      strNmbr = strNmbr.Trim.Replace("'", "").Trim
      strName = strName.Trim.Replace("'", "").Trim

      strSQL = String.Format("INSERT INTO tblCallerLog (ts, nmbr, name) VALUES ({0}, '{1}', '{2}')", ts.ToString, strNmbr, strName)

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "InsertCallerLogData() Reports Error: [" & pEx.ToString & "], " _
                  & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strMessage, MessageType.Error)
    End Try

  End Sub

  ''' <summary>
  ''' Insert caller detail data into database
  ''' </summary>
  ''' <param name="ts"></param>
  ''' <param name="strNmbr"></param>
  ''' <param name="strName"></param>
  ''' <remarks></remarks>
  Public Sub InsertCallerDetailData(ByVal ts As Integer, _
                                    ByVal strNmbr As String, _
                                    ByVal strName As String)

    Dim strMessage As String = ""
    Dim strSQL As String = ""
    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Exit Sub

    Try
      '
      ' Check for default caller attributes
      '
      Dim iAnnounceAttr As Integer = IIf(CStr(hs.GetINISetting("CallerAttr", "Announce", "True", gINIFile)) = "True", 1, 0)

      Dim iCallerAttributes As Integer = 0
      Try
        Dim bUseNmbrBlockMask As Boolean = CBool(hs.GetINISetting("CallerAttr", "UseNmbrBlockMask", "False", gINIFile))
        Dim strNmbrBlockMask As String = hs.GetINISetting("CallerAttr", "NmbrBlockMask", "", gINIFile)

        If bUseNmbrBlockMask = True And Regex.IsMatch(strNmbr, strNmbrBlockMask) = True Then
          iCallerAttributes = iCallerAttributes Or CallerAttrs.Block
        End If
      Catch pEx As Exception
        strMessage = "InsertCallerDetailData() Reports Error: [" & pEx.ToString & "]"
        Call WriteMessage(strMessage, MessageType.Error)
      End Try

      Try
        Dim bUseNameBlockMask As Boolean = CBool(hs.GetINISetting("CallerAttr", "UseNameBlockMask", "False", gINIFile))
        Dim strNameBlockMask As String = hs.GetINISetting("CallerAttr", "NameBlockMask", "", gINIFile)

        If bUseNameBlockMask = True And Regex.IsMatch(strName, strNameBlockMask) = True Then
          iCallerAttributes = iCallerAttributes Or CallerAttrs.Block
        End If
      Catch pEx As Exception
        strMessage = "InsertCallerDetailData() Reports Error: [" & pEx.ToString & "]"
        Call WriteMessage(strMessage, MessageType.Error)
      End Try

      '
      ' Determine if caller should be announced
      '
      If iAnnounceAttr = 1 And iCallerAttributes = 0 Then
        iCallerAttributes = iCallerAttributes Or CallerAttrs.Announce
      End If

      '
      ' Insert caller into caller detail table
      '
      strSQL = String.Format("INSERT INTO tblCallerDetails (last_ts, nmbr, name, attr, call_count) VALUES ({0}, '{1}', '{2}', {3}, {4})", ts.ToString, strNmbr, strName, iCallerAttributes.ToString, 1)

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()
        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "InsertCallerDetailData() Reports Error: [" & pEx.ToString & "], " _
                  & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strMessage, MessageType.Error)
    End Try

  End Sub

  ''' <summary>
  ''' Update caller detail data in database
  ''' </summary>
  ''' <param name="ts"></param>
  ''' <param name="strNmbr"></param>
  ''' <param name="strName"></param>
  ''' <remarks></remarks>
  Public Sub UpdateCallerDetailData(ByVal ts As Integer, _
                                    ByVal strNmbr As String, _
                                    ByVal strName As String)

    Dim strMessage As String = ""
    Dim strSQL As String = ""
    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Exit Sub

    Try
      '
      ' Remove single quotes
      '
      strNmbr = strNmbr.Trim.Replace("'", "")
      strName = strName.Trim.Replace("'", "")

      '
      ' Update last timestamp and caller count
      '
      strSQL = String.Format("UPDATE tblCallerDetails SET " & _
                              " last_ts = {0}, " &
                              " call_count = call_count + 1 " &
                              "WHERE nmbr = '{1}'", ts.ToString, strNmbr)

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()
        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "UpdateCallerDetailData() Reports Error: [" & pEx.ToString & "], " _
                  & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strMessage, MessageType.Error)
    End Try

  End Sub

  ''' <summary>
  ''' Gets the Caller Number By Id
  ''' </summary>
  ''' <param name="Id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCallerNumberById(ByVal Id As Integer) As String

    Dim nmbr As String = String.Empty

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Return nmbr

    Try
      '
      ' Determine which SQL should be executed
      '
      Dim strSQL As String = String.Format("SELECT nmbr from tblCallerDetails where id = {0}", Id)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            nmbr = dtrResults("nmbr")
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetCallerNumberById()")
    End Try

    Return nmbr

  End Function

  ''' <summary>
  ''' Gets the caller alias from the database
  ''' </summary>
  ''' <param name="nmbr"></param>
  ''' <param name="name"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCallerAlias(ByVal nmbr As String, ByVal name As String) As String

    Dim strCallerAlias As String = name

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Return name

    Try
      '
      ' Determine which SQL should be executed
      '
      Dim strSQL As String = String.Format("SELECT name from tblCallerDetails where nmbr = '{0}'", nmbr)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            strCallerAlias = dtrResults("name")
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetCallerAlias()")
    End Try

    Return strCallerAlias

  End Function

  ''' <summary>
  ''' Gets the caller attributes from the database
  ''' </summary>
  ''' <param name="nmbr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCallerAttributes(ByVal nmbr As String) As Integer

    Dim iCallerAttr As Integer = -1

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Return 0

    Try
      '
      ' Determine which SQL should be executed
      '
      Dim strSQL As String = String.Format("SELECT attr FROM tblCallerDetails WHERE nmbr = '{0}'", nmbr)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            iCallerAttr = CInt(dtrResults("attr"))
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetCallerAttributes()")
    End Try

    Return iCallerAttr

  End Function

  ''' <summary>
  ''' Gets the caller attributes from the database
  ''' </summary>
  ''' <param name="nmbr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SetCallerAttributes(ByVal nmbr As String, ByRef attr As Integer) As Boolean

    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Return 0

    Try
      '
      ' Determine which SQL should be executed
      '
      Dim strSQL As String = String.Format("UPDATE tblCallerDetails set attr={0} WHERE nmbr = '{1}'", attr, nmbr)

      '
      ' ExecuteNonQuery
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = MyDbCommand.ExecuteNonQuery()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SetCallerAttributes()")
      Return False
    End Try

  End Function

#End Region

#Region "Database Date Formatting"

  ''' <summary>
  ''' dateTime as DateTime
  ''' </summary>
  ''' <param name="dateTime"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertDateTimeToEpoch(ByVal dateTime As DateTime) As Long

    Dim baseTicks As Long = 621355968000000000
    Dim tickResolution As Long = 10000000

    Return (dateTime.ToUniversalTime.Ticks - baseTicks) / tickResolution

  End Function

  ''' <summary>
  ''' Converts Epoch to datetime
  ''' </summary>
  ''' <param name="epochTicks"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertEpochToDateTime(ByVal epochTicks As Long) As DateTime

    '
    ' Create a new DateTime value based on the Unix Epoch
    '
    Dim converted As New DateTime(1970, 1, 1, 0, 0, 0, 0)

    '
    ' Return the value in string format
    '
    Return converted.AddSeconds(epochTicks).ToLocalTime

  End Function

#End Region

End Module

