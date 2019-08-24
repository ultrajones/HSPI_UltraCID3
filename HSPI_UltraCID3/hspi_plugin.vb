Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Data.Common
Imports HomeSeerAPI
Imports Scheduler
Imports System.ComponentModel
Imports System.Data.SQLite

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  'Dim action As New action
  Dim triggers As New hsCollection
  'Dim trigger As New trigger
  Dim conditions As New Hashtable
  Const Pagename = "Events"

  Public ModemInterfaces As New List(Of hspi_modem)
  Public gModemInterfaces As Integer = 1

  Public Const IFACE_NAME As String = "UltraCID3"

  Public Const LINK_TARGET As String = "hspi_ultracid3/hspi_ultracid3.aspx"
  Public Const LINK_URL As String = "hspi_ultracid3.html"
  Public Const LINK_TEXT As String = "UltraCID3"
  Public Const LINK_PAGE_TITLE As String = "UltraCID3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultracid3/UltraCID3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = String.Empty
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultracid3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public Const MODEM_INIT As String = "ATQ0V1E0~AT+GMM~AT+VCID=1~AT+FCLASS=8~AT-STE=7"
  Public Const DROP_CALLER As String = "AT+VLS=5~ATH~AT+FCLASS=8"
  Public Const CID_QUERY_URL As String = "http://www.google.com/search?hl=en&q=$nmbr"

  Public Const NMBR_FORMAT_09 As String = "(0#) ###-####"
  Public Const NMBR_FORMAT_10 As String = "(0##) ###-####"
  Public Const NMBR_FORMAT_11 As String = "0## #### ####"

  Public Const EMAIL_SUBJECT As String = "UltraCID3 E-Mail Notification"
  Public Const EMAIL_BODY_TEMPLATE As String = "$name called from $nmbr at $datetime."

  Public gNumberQueryURL As String = String.Empty

  Public DEV_HARDWARE_MODEM As Byte = 1
  Public DEV_DATABASE_INTERFACE As Byte = 2
  Public DEV_EXTENSION_STATE As Byte = 3
  Public DEV_LAST_CALLER_NAME As Byte = 4
  Public DEV_LAST_CALLER_NMBR As Byte = 5
  Public DEV_LAST_CALLER_ATTR As Byte = 6
  Public DEV_LAST_CALLER_RING As Byte = 7

  Public MAX_ATTEMPTS As Byte = 2

#Region "UltraCID3 Public Functions"

  ''' <summary>
  ''' Returns a list of valid serial ports
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSerialPortNames() As Specialized.StringCollection

    Dim PortNames As New Specialized.StringCollection

    Try
      '
      ' Find all available serial port names on this computer
      '
      For Each strPortName As String In My.Computer.Ports.SerialPortNames
        PortNames.Add(strPortName)
      Next
    Catch pEx As Exception

    End Try

    Return PortNames

  End Function

  ''' <summary>
  ''' Gets the modem status
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetModemStatus(ByVal lineNumber As Integer) As String

    Dim ModemInterface As hspi_modem = ModemInterfaces.Find(Function(ModemLine) ModemLine.lineNumber = lineNumber)
    If Not IsNothing(ModemInterface) = True Then
      Return ModemInterface.GetModemStatus
    Else
      Return "Unknown"
    End If

  End Function

  ''' <summary>
  ''' Returns various statistics
  ''' </summary>
  ''' <param name="StatisticsType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStatistics(ByVal lineNumber As Integer, ByVal StatisticsType As String) As Integer

    Dim ModemInterface As hspi_modem = ModemInterfaces.Find(Function(ModemLine) ModemLine.lineNumber = lineNumber)
    If Not IsNothing(ModemInterface) = True Then
      Return ModemInterface.GetStatistics(StatisticsType)
    Else
      Return 0
    End If

  End Function

  ''' <summary>
  ''' Formats caller ID for display
  ''' </summary>
  ''' <param name="nmbr"></param>
  ''' <param name="name"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FormatCID(ByVal nmbr As String, ByVal name As String) As String

    Try

      Return String.Format("{0} [{1}]", FormatNmbr(nmbr), name)

    Catch pEx As Exception
      ' Ignore error
      Return String.Format("{0} [{1}]", nmbr, name)
    End Try

  End Function

  ''' <summary>
  ''' Formats telephone number for display
  ''' </summary>
  ''' <param name="nmbr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FormatNmbr(ByVal nmbr As String, Optional ByVal bDisplayAsURL As Boolean = False) As String

    Dim strNmbr As String = nmbr

    Try

      If IsNumeric(nmbr) Then

        If nmbr.StartsWith("1") AndAlso nmbr.Length > 9 Then
          nmbr = nmbr.TrimStart("1")
        End If
        Select Case nmbr.Length
          Case 9
            Dim strNmbrFormat As String = hs.GetINISetting("NmbrFormat", "09", NMBR_FORMAT_09, gINIFile)
            strNmbr = CDbl(nmbr).ToString(strNmbrFormat)
          Case 10
            Dim strNmbrFormat As String = hs.GetINISetting("NmbrFormat", "10", NMBR_FORMAT_10, gINIFile)
            strNmbr = CDbl(nmbr).ToString(strNmbrFormat)
          Case 11
            Dim strNmbrFormat As String = hs.GetINISetting("NmbrFormat", "11", NMBR_FORMAT_11, gINIFile)
            strNmbr = CDbl(nmbr).ToString(strNmbrFormat)
        End Select

      End If

      If bDisplayAsURL = True Then
        Dim strURL As String = gNumberQueryURL.Replace("$nmbr", nmbr)
        strNmbr = String.Format("<a target=""{0}"" href=""{1}"">{2}</a>", "_blank", strURL, strNmbr)
      End If

    Catch ex As Exception
      ' Ignore error
    End Try

    Return strNmbr

  End Function

  ''' <summary>
  ''' Formats caller attribures for display
  ''' </summary>
  ''' <param name="attr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function FormatCallerAttrs(ByVal attr As String) As String

    Dim strCallerAttr As String = ""

    Try

      If IsNumeric(attr) Then

        Dim iCallerAttr As Integer = CInt(attr)

        '
        ' Loop through the enumeration members
        '
        Dim names As String() = System.Enum.GetNames(GetType(CallerAttrs))
        For i As Integer = 0 To names.Length - 1
          Dim iAttr As CallerAttrs = System.Enum.Parse(GetType(CallerAttrs), names(i))
          If iCallerAttr And iAttr Then
            strCallerAttr &= iAttr & ","
          End If
        Next

      End If

    Catch pEx As Exception
      ' Ignore error
    End Try

    Return strCallerAttr.TrimEnd(",")

  End Function

  ''' <summary>
  ''' Formats caller attribures for display
  ''' </summary>
  ''' <param name="attrs"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function FormatCallerAttrNames(ByVal attrs As Integer) As String

    Try
      '
      ' Update Caller Attributes
      '
      Dim CallerAttributes As New List(Of String)
      Dim names As String() = System.Enum.GetNames(GetType(CallerAttrs))

      For i As Integer = 0 To names.Length - 1
        Dim iAttr As CallerAttrs = System.Enum.Parse(GetType(CallerAttrs), names(i))
        If attrs And iAttr Then
          CallerAttributes.Add(names(i))
        End If
      Next

      If CallerAttributes.Count = 0 Then CallerAttributes.Add("None")
      Return String.Join(",", CallerAttributes.ToArray)

    Catch pEx As Exception
      ' Ignore error
    End Try

    Return "None"

  End Function

  ''' <summary>
  ''' Validates the provided number format
  ''' </summary>
  ''' <param name="strNmbrFormatType"></param>
  ''' <param name="strNmbrFormatValue"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidateNmbrFormat(ByVal strNmbrFormatType As String, ByVal strNmbrFormatValue As String)

    Try

      Dim objMatches As MatchCollection = Regex.Matches(strNmbrFormatValue, "#|0", RegexOptions.None)
      Dim iMatchCount As Integer = objMatches.Count

      Select Case strNmbrFormatType
        Case "tbNumberFormat09"
          If iMatchCount <> 9 Then Return False
        Case "tbNumberFormat10"
          If iMatchCount <> 10 Then Return False
        Case "tbNumberFormat11"
          If iMatchCount <> 11 Then Return False
        Case Else
          Return False
      End Select

      Return True

    Catch ex As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Builds the caller log SQL query
  ''' </summary>
  ''' <param name="dbTable"></param>
  ''' <param name="dbFields"></param>
  ''' <param name="startDate"></param>
  ''' <param name="endDate"></param>
  ''' <param name="filterType"></param>
  ''' <param name="sortOrder"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildCallerLogSQL(ByVal dbTable As String, _
                                    ByVal dbFields As String, _
                                    ByVal startDate As String, _
                                    ByVal endDate As String, _
                                    ByVal filterType As String, _
                                    ByVal sortOrder As String) As String

    Try

      Dim dStartDateTime As Date
      Date.TryParse(startDate, dStartDateTime)

      Dim dEndDateTime As Date
      Date.TryParse(endDate, dEndDateTime)

      Dim ts_start = ConvertDateTimeToEpoch(dStartDateTime)
      Dim ts_end = ConvertDateTimeToEpoch(dEndDateTime)

      Dim dateField As String = "ts"
      Dim filterField As String = "nmbr"

      Dim comparison As String = "="
      If filterType = "%" Then
        comparison = "LIKE"
      End If

      '
      ' Build SQL
      '
      Dim strSQL As String = String.Format("SELECT {0} " _
                                         & "FROM {1} " _
                                         & "WHERE {2} >= {3} " _
                                         & "AND {2} <= {4} " _
                                         & "AND {5} {6} '{7}' " _
                                         & "ORDER BY {2} {8}", dbFields, dbTable, dateField, ts_start.ToString, ts_end.ToString, filterField, comparison, filterType, sortOrder)

      WriteMessage(strSQL, MessageType.Debug)

      Return strSQL

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "BuildCallerLogSQL()")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Gets the caller summary from the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCallerSummary() As SortedList

    Dim CallerSummary As New SortedList
    Dim strSQL As String = ""
    Dim iIndex As Long = 0

    Try
      '
      ' Determine which SQL should be executed
      '
      strSQL = "SELECT nmbr, name from tblCallerDetails"

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
            If CallerSummary.ContainsKey(dtrResults("nmbr")) = False Then
              CallerSummary.Add(dtrResults("nmbr"), dtrResults("name"))
            End If
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetCallerSummary()")
    End Try

    Return CallerSummary

  End Function

  ''' <summary>
  ''' Execute Raw SQL
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <param name="iRecordCount"></param>
  ''' <param name="iPageSize"></param>
  ''' <param name="iPageCount"></param>
  ''' <param name="iPageCur"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ExecuteSQL(ByVal strSQL As String, _
                             ByRef iRecordCount As Integer, _
                             ByVal iPageSize As Integer, _
                             ByRef iPageCount As Integer, _
                             ByRef iPageCur As Integer) As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered ExecuteSQL() function."
    Call WriteMessage(strMessage, MessageType.Debug)

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
      ' Determine Requested database action
      '
      If strSQL.StartsWith("SELECT", StringComparison.CurrentCultureIgnoreCase) Then

        '
        ' Populate the DataSet
        '

        '
        ' Initialize the command object
        '
        Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

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

        '
        ' Get our DataTable
        '
        Dim MyDT As DataTable = MyDS.Tables(0)

        '
        ' Get record count
        '
        iRecordCount = MyDT.Rows.Count

        If iRecordCount > 0 Then
          '
          ' Determine the number of pages available
          '
          iPageSize = IIf(iPageSize <= 0, 1, iPageSize)
          iPageCount = iRecordCount \ iPageSize
          If iRecordCount Mod iPageSize > 0 Then
            iPageCount += 1
          End If

          '
          ' Find starting and ending record
          '
          Dim nStart As Integer = iPageSize * (iPageCur - 1)
          Dim nEnd As Integer = nStart + iPageSize - 1
          If nEnd > iRecordCount - 1 Then
            nEnd = iRecordCount - 1
          End If

          '
          ' Build field names
          '
          Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
          For iFieldNum As Integer = 0 To iFieldCount
            '
            ' Create the columns
            '
            Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
            Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

            '
            ' Add the columns to the DataTable's Columns collection
            '
            ResultsDT.Columns.Add(MyDataColumn)
          Next

          '
          ' Let's output our records	
          '
          Dim i As Integer = 0
          For i = nStart To nEnd
            'Add some rows
            Dim dr As DataRow
            dr = ResultsDT.NewRow()
            For iFieldNum As Integer = 0 To iFieldCount
              dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
            Next
            ResultsDT.Rows.Add(dr)
          Next

          '
          ' Make sure current page count is valid
          '
          If iPageCur > iPageCount Then iPageCur = iPageCount

        Else
          '
          ' Query succeeded, but returned 0 records
          '
          strMessage = "Your query executed and returned 0 record(s)."
          Call WriteMessage(strMessage, MessageType.Debug)

        End If

      Else
        '
        ' Execute query (does not return recordset)
        '
        Try
          '
          ' Build the insert/update/delete query
          '
          Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

            MyDbCommand.Connection = DBConnectionMain
            MyDbCommand.CommandType = CommandType.Text
            MyDbCommand.CommandText = strSQL

            Dim iRecordsAffected As Integer = 0
            SyncLock SyncLockMain
              iRecordsAffected = MyDbCommand.ExecuteNonQuery()
            End SyncLock

            strMessage = "Your query executed and affected " & iRecordsAffected & " row(s)."
            Call WriteMessage(strMessage, MessageType.Debug)

            MyDbCommand.Dispose()

          End Using

        Catch pEx As Common.DbException
          '
          ' Process Database Error
          '
          strMessage = "Your query failed for the following reason(s):  "
          strMessage &= "[Error Source: " & pEx.Source & "] " _
                      & "[Error Number: " & pEx.ErrorCode & "] " _
                      & "[Error Desciption:  " & pEx.Message & "] "
          Call WriteMessage(strMessage, MessageType.Error)
        End Try

      End If

      Call WriteMessage("SQL: " & strSQL, MessageType.Debug)
      Call WriteMessage("Record Count: " & iRecordCount, MessageType.Debug)
      Call WriteMessage("Page Count: " & iPageCount, MessageType.Debug)
      Call WriteMessage("Page Current: " & iPageCur, MessageType.Debug)

      '
      ' Populate the table name
      '
      ResultsDT.TableName = String.Format("Table1:{0}:{1}:{2}", iRecordCount, iPageCount, iPageCur)

    Catch pEx As Exception
      '
      ' Error:  Query error
      '
      strMessage = "Your query failed for the following reason:  " _
                  & Err.Source & " function/subroutine:  [" _
                  & Err.Number & " - " & Err.Description _
                  & "]"

      Call ProcessError(pEx, "ExecuteSQL()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered GetSetting() function.", MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  '''  Saves plug-in settings to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered SaveSetting() subroutine.", MessageType.Debug)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "UltraCID3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      triggers.Add(o, "Telephone Ring")   ' 1
      triggers.Add(o, "Incoming Call")    ' 2
      triggers.Add(o, "Local Handset")    ' 3
      triggers.Add(o, "Phone Extension")  ' 4
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case 1, 2, 3, 4
          Return False  ' Render trigger as IF / OR IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber
      Case CIDTriggers.TelephoneRing
        Dim triggerName As String = GetEnumName(CIDTriggers.TelephoneRing)
        Dim triggerId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "Line", UID, sUnique)
        Dim triggerSelected As String = trigger.Item("Line")

        '
        ' Configure Line Number
        '
        Dim jqLine As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqLine.autoPostBack = True
        jqLine.AddItem("--Please Select--", "", False)
        jqLine.AddItem("Any", "0", (triggerSelected = "0"))
        jqLine.AddItem("Line 1", "1", (triggerSelected = "1"))
        jqLine.AddItem("Line 2", "2", (triggerSelected = "2"))

        stb.Append("Select Line:")
        stb.Append(jqLine.Build)

        '
        ' Configure Ring Count
        '
        triggerId = String.Format("{0}{1}_{2}_{3}", triggerName, "Ring", UID, sUnique)
        triggerSelected = trigger.Item("Ring")
        Dim jqDropList As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqDropList.autoPostBack = True

        jqDropList.AddItem("--Please Select--", "", False)
        jqDropList.AddItem("Any", "0", (triggerSelected = "0"))
        jqDropList.AddItem("First", "1", (triggerSelected = "1"))
        jqDropList.AddItem("Second", "2", (triggerSelected = "2"))
        jqDropList.AddItem("Third", "3", (triggerSelected = "3"))
        jqDropList.AddItem("Forth", "4", (triggerSelected = "4"))
        jqDropList.AddItem("Fifth", "5", (triggerSelected = "5"))

        stb.Append("Select Ring:")
        stb.Append(jqDropList.Build)

      Case CIDTriggers.IncomingCall
        Dim triggerName As String = GetEnumName(CIDTriggers.IncomingCall)
        Dim triggerId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "Line", UID, sUnique)
        Dim triggerSelected As String = trigger.Item("Line")

        '
        ' Configure Line Number
        '
        Dim jqLine As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqLine.autoPostBack = True
        jqLine.AddItem("--Please Select--", "", False)
        jqLine.AddItem("Any", "0", (triggerSelected = "0"))
        jqLine.AddItem("Line 1", "1", (triggerSelected = "1"))
        jqLine.AddItem("Line 2", "2", (triggerSelected = "2"))

        stb.Append("Select Line:")
        stb.Append(jqLine.Build)

        '
        ' Configure Caller
        '
        triggerId = String.Format("{0}{1}_{2}_{3}", triggerName, "Caller", UID, sUnique)
        triggerSelected = trigger.Item("Caller")
        Dim jqDropList As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqDropList.autoPostBack = True
        jqDropList.AddItem("--Please Select--", "", False)
        jqDropList.AddItem("Any Caller", "Any", (triggerSelected = "Any"))
        jqDropList.AddItem("First Time Caller", "FirstTime", (triggerSelected = "FirstTime"))
        jqDropList.AddItem("Announce Caller", "Announce", (triggerSelected = "Announce"))
        jqDropList.AddItem("Blocked Caller", "Blocked", (triggerSelected = "Blocked"))
        jqDropList.AddItem("Business Caller", "Business", (triggerSelected = "Business"))
        jqDropList.AddItem("Family Caller", "Family", (triggerSelected = "Family"))
        jqDropList.AddItem("Friends Caller", "Friends", (triggerSelected = "Friends"))
        jqDropList.AddItem("Telemarketer Caller", "Telemarketer", (triggerSelected = "Telemarketer"))

        Dim CallerSummary As New SortedList
        CallerSummary = GetCallerSummary()

        For Each nmbr As String In CallerSummary.Keys
          Dim strOptionValue As String = nmbr
          Dim strOptionDesc As String = FormatCID(nmbr, CallerSummary(nmbr))
          jqDropList.AddItem(strOptionDesc, strOptionValue, (triggerSelected = strOptionValue))
        Next

        stb.Append("Select Caller:")
        stb.Append(jqDropList.Build)

      Case CIDTriggers.LocalHandset
        Dim triggerName As String = GetEnumName(CIDTriggers.LocalHandset)
        Dim triggerId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "Line", UID, sUnique)
        Dim triggerSelected As String = trigger.Item("Line")

        '
        ' Configure Line Number
        '
        Dim jqLine As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqLine.autoPostBack = True
        jqLine.AddItem("--Please Select--", "", False)
        jqLine.AddItem("Any", "0", (triggerSelected = "0"))
        jqLine.AddItem("Line 1", "1", (triggerSelected = "1"))
        jqLine.AddItem("Line 2", "2", (triggerSelected = "2"))

        stb.Append("Select Line:")
        stb.Append(jqLine.Build)

        '
        ' Configure Local Handset State
        '
        triggerId = String.Format("{0}{1}_{2}_{3}", triggerName, "State", UID, sUnique)
        triggerSelected = trigger.Item("State")
        Dim jqDropList As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqDropList.autoPostBack = True
        jqDropList.AddItem("--Please Select--", "", False)
        jqDropList.AddItem("Any", "Any", (triggerSelected = "Any"))
        jqDropList.AddItem("On Hook", "On Hook", (triggerSelected = "On Hook"))
        jqDropList.AddItem("Off Hook", "Off Hook", (triggerSelected = "Off Hook"))

        stb.Append("Select Action:")
        stb.Append(jqDropList.Build)

      Case CIDTriggers.PhoneExtension
        Dim triggerName As String = GetEnumName(CIDTriggers.PhoneExtension)
        Dim triggerId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "Line", UID, sUnique)
        Dim triggerSelected As String = trigger.Item("Line")

        '
        ' Configure Line Number
        '
        Dim jqLine As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqLine.autoPostBack = True
        jqLine.AddItem("--Please Select--", "", False)
        jqLine.AddItem("Any", "0", (triggerSelected = "0"))
        jqLine.AddItem("Line 1", "1", (triggerSelected = "1"))
        jqLine.AddItem("Line 2", "2", (triggerSelected = "2"))

        stb.Append("Select Line:")
        stb.Append(jqLine.Build)

        '
        ' Configure Phone Extension State
        '
        triggerId = String.Format("{0}{1}_{2}_{3}", triggerName, "State", UID, sUnique)
        triggerSelected = trigger.Item("State")
        Dim jqDropList As New clsJQuery.jqDropList(triggerId, Pagename, True)
        jqDropList.autoPostBack = True
        jqDropList.AddItem("--Please Select--", "", False)
        jqDropList.AddItem("Any", "Any", (triggerSelected = "Any"))
        jqDropList.AddItem("On Hook", "On Hook", (triggerSelected = "On Hook"))
        jqDropList.AddItem("Off Hook", "Off Hook", (triggerSelected = "Off Hook"))

        stb.Append("Select Action:")
        stb.Append(jqDropList.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case CIDTriggers.TelephoneRing
          Dim triggerName As String = GetEnumName(CIDTriggers.TelephoneRing)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "Line_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Line") = ActionValue

              Case InStr(sKey, triggerName & "Ring_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Ring") = ActionValue

            End Select
          Next

        Case CIDTriggers.IncomingCall
          Dim triggerName As String = GetEnumName(CIDTriggers.IncomingCall)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "Line_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Line") = ActionValue

              Case InStr(sKey, triggerName & "Caller_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Caller") = ActionValue

            End Select
          Next

        Case CIDTriggers.LocalHandset
          Dim triggerName As String = GetEnumName(CIDTriggers.LocalHandset)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "Line_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Line") = ActionValue

              Case InStr(sKey, triggerName & "State_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("State") = ActionValue

            End Select
          Next

        Case CIDTriggers.PhoneExtension
          Dim triggerName As String = GetEnumName(CIDTriggers.PhoneExtension)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "Line_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Line") = ActionValue

              Case InStr(sKey, triggerName & "State_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("State") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case CIDTriggers.TelephoneRing
          If trigger.Item("Line") = "" Then Configured = False
          If trigger.Item("Ring") = "" Then Configured = False

        Case CIDTriggers.IncomingCall
          If trigger.Item("Line") = "" Then Configured = False
          If trigger.Item("Caller") = "" Then Configured = False

        Case CIDTriggers.LocalHandset
          If trigger.Item("Line") = "" Then Configured = False
          If trigger.Item("State") = "" Then Configured = False

        Case CIDTriggers.PhoneExtension
          If trigger.Item("Line") = "" Then Configured = False
          If trigger.Item("State") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property


  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case CIDTriggers.TelephoneRing
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim lineNumber As String = trigger.Item("Line")
          Dim triggerValue As String = trigger.Item("Ring")
          Select Case triggerValue
            Case "0"
              triggerValue = "Any"
            Case "1"
              triggerValue = "First"
            Case "2"
              triggerValue = "Second"
            Case "3"
              triggerValue = "Third"
            Case "4"
              triggerValue = "Forth"
            Case "5"
              triggerValue = "Fifth"
          End Select

          stb.AppendFormat("UltraCID3 Telephone Ring: {0} {1}", lineNumber, triggerValue)
        End If

      Case CIDTriggers.IncomingCall
        Dim lineNumber As String = trigger.Item("Line")
        Dim triggerValue As String = trigger.Item("Caller")
        Dim callerInfo As String = triggerValue

        Dim CallerSummary As New SortedList
        CallerSummary = GetCallerSummary()

        If CallerSummary.ContainsKey(triggerValue) Then
          callerInfo = FormatCID(triggerValue, CallerSummary(triggerValue))
        End If

        stb.AppendFormat("UltraCID3 Incoming Call: {0} {1}", lineNumber, callerInfo)

      Case CIDTriggers.LocalHandset
        Dim lineNumber As String = trigger.Item("Line")
        Dim triggerValue As String = trigger.Item("State")

        stb.AppendFormat("UltraCID3 Local Handset: {0} {1}", lineNumber, triggerValue)

      Case CIDTriggers.PhoneExtension
        Dim lineNumber As String = trigger.Item("Line")
        Dim triggerValue As String = trigger.Item("State")

        stb.AppendFormat("UltraCID3 Phone Extension: {0} {1}", lineNumber, triggerValue)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Check if Trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Public Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID
              Case CIDTriggers.TelephoneRing
                Dim strLineNumber As String = trigger.Item("Line")
                Dim strEventType As String = "TelephoneRing"
                Dim strStatus As String = trigger.Item("Ring")

                Dim strTriggerCheck As String = String.Format("{0},{1},{2}", strLineNumber, strEventType, strStatus)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

              Case CIDTriggers.IncomingCall
                Dim strLineNumber As String = trigger.Item("Line")
                Dim strEventType As String = "IncomingCall"
                Dim strStatus As String = trigger.Item("Caller")

                Dim strTriggerCheck As String = String.Format("{0},{1},{2}", strLineNumber, strEventType, strStatus)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

              Case CIDTriggers.LocalHandset
                Dim strLineNumber As String = trigger.Item("Line")
                Dim strEventType As String = "LocalHandset"
                Dim strStatus As String = trigger.Item("State")

                Dim strTriggerCheck As String = String.Format("{0},{1},{2}", strLineNumber, strEventType, strStatus)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

              Case CIDTriggers.PhoneExtension
                Dim strLineNumber As String = trigger.Item("Line")
                Dim strEventType As String = "PhoneExtension"
                Dim strStatus As String = trigger.Item("State")

                Dim strTriggerCheck As String = String.Format("{0},{1},{2}", strLineNumber, strEventType, strStatus)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      actions.Add(o, "Drop Caller") ' 1
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = ActInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case CIDActions.DropCaller
        Dim actionName As String = GetEnumName(CIDActions.DropCaller)
        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "Line", UID, sUnique)
        Dim ActionSelected As String = action.Item("Line")

        '
        ' Configure Line Number
        '
        Dim jqLine As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqLine.autoPostBack = True
        jqLine.AddItem("--Please Select--", "", False)
        jqLine.AddItem("Any", "0", (ActionSelected = "0"))
        jqLine.AddItem("Line 1", "1", (ActionSelected = "2"))
        jqLine.AddItem("Line 2", "2", (ActionSelected = "2"))

        stb.Append("Select Line:")
        stb.Append(jqLine.Build)

        '
        ' Configure Drop Caller
        '
        actionId = String.Format("{0}_{1}_{2}", actionName, UID, sUnique)
        ActionSelected = action.Item("Action")
        Dim jqDropList As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqDropList.autoPostBack = True

        jqDropList.AddItem("--Please Select--", "", False)
        jqDropList.AddItem("Drop Caller", "Drop Caller", (ActionSelected = "Drop Caller"))

        stb.Append("Select Action:")
        stb.Append(jqDropList.Build)

    End Select

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = ActInfo.UID.ToString
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case CIDActions.DropCaller
          Dim actionName As String = GetEnumName(CIDActions.DropCaller)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "Line_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Line") = ActionValue

              Case InStr(sKey, "DropCaller_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Action") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case CIDActions.DropCaller
        If action.Item("Line") = "" Then Configured = False
        If action.Item("Action") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case CIDActions.DropCaller
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim lineNumber As String = action.Item("Line")
          Dim actionName As String = action.Item("Action")

          stb.AppendFormat("UltraCID3 Action: {0} {1}", lineNumber, actionName)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber
        Case CIDActions.DropCaller
          Dim lineNumber As String = action.Item("Line")
          Dim actionName As String = GetEnumName(CIDActions.DropCaller)

          '
          ' Drop caller
          '
          Dim ModemInterface As hspi_modem = ModemInterfaces.Find(Function(ModemLine) ModemLine.lineNumber = lineNumber)
          If Not IsNothing(ModemInterface) = True Then
            Call ModemInterface.DropCaller()
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum CIDTriggers
  <Description("Telephone Ring")> _
  TelephoneRing = 1
  <Description("Incoming Call")> _
  IncomingCall = 2
  <Description("Local Handset")> _
  LocalHandset = 3
  <Description("Phone Extension")> _
  PhoneExtension = 4
End Enum

Public Enum CIDActions
  <Description("Drop Caller")>
  DropCaller = 1
End Enum

<Flags()> Public Enum CallerAttrs
  None = 0
  Block = 1
  Telemarketer = 2
  Announce = 4
  Business = 8
  Family = 16
  Friends = 32
End Enum
