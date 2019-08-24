Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class hspi_modem
  Public WithEvents serialPort As New IO.Ports.SerialPort

  Private ModemInit As New Specialized.StringDictionary
  Private ModemCmds As New Specialized.StringDictionary

  Private CommandQueue As New Queue

  Private _lineNumber As Integer = 0
  Private _interface As String = String.Empty

  Private _connectionPortName As String = ""
  Private _connected As Boolean = False

  Private _cmdWait As String = String.Empty

  Private _callerName As String = String.Empty
  Private _callerNmbr As String = String.Empty
  Private _callerAttr As Integer = 0

  Private _lastRingTicks As Long = 0
  Private _ringCount As Long = 0
  Private _callCount As Long = 0
  Private _callRings As Integer = 0
  Private _dropCount As Long = 0

  Private SendCommandThread As Thread

  Public Sub New(ByVal lineNumber As Integer)

    _lineNumber = lineNumber
    _interface = String.Format("Interface{0}", _lineNumber)

    '
    ' Start the process command queue thread
    '
    SendCommandThread = New Thread(New ThreadStart(AddressOf ProcessCommandQueue))
    SendCommandThread.Name = "ProcessCommandQueue"
    SendCommandThread.Start()

    WriteMessage(String.Format("{0} Thread Started", SendCommandThread.Name), MessageType.Debug)

  End Sub

  ''' <summary>
  ''' Reutrn the Modem Line Number
  ''' </summary>
  ''' <returns></returns>
  Public Property lineNumber() As Integer
    Get
      Return _lineNumber
    End Get
    Set(value As Integer)
      _lineNumber = value
    End Set
  End Property

  Public Property CallRings As Integer
    Get
      Return _callRings
    End Get
    Set(value As Integer)
      _callRings = value
    End Set
  End Property

  Public Property DropCount As Integer
    Get
      Return _dropCount
    End Get
    Set(value As Integer)
      _dropCount = value
    End Set
  End Property

  Public Property CallCount As Integer
    Get
      Return _callCount
    End Get
    Set(value As Integer)
      _callCount = value
    End Set
  End Property

  Public Property RingCount As Integer
    Get
      Return _ringCount
    End Get
    Set(value As Integer)
      _ringCount = value
    End Set
  End Property

  Public Property CallerNumber As String
    Get
      Return _callerNmbr
    End Get
    Set(value As String)
      _callerNmbr = value
    End Set
  End Property

#Region "UltraCID3 Modem Hardware"

  ''' <summary>
  ''' Initiialze the modem
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitHW() As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered InitHW() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Do hardware initialization here (optional)
      '
      _connectionPortName = hs.GetINISetting(_interface, "Serial", "", gINIFile)

      '
      ' Bail out when no port is defined (user has not set a port yet)
      '
      If _connectionPortName.Length = 0 Then
        strMessage = "Modem serial port not set."
        Call WriteMessage(strMessage, MessageType.Warning)

        gIOEnabled = False
      Else
        '
        ' Try connecting to the serial port
        '
        strMessage = "Initiating connection to " & _connectionPortName & "..."
        Call WriteMessage(strMessage, MessageType.Debug)

        Try
          '
          ' Close port if already open
          '
          If serialPort.IsOpen Then
            serialPort.Close()
          End If

          With serialPort
            .ReadTimeout = 500
            .PortName = _connectionPortName
            .BaudRate = 115200
            .Parity = IO.Ports.Parity.None
            .DataBits = 8
            .StopBits = IO.Ports.StopBits.One
          End With
          serialPort.Open()

          _connected = serialPort.IsOpen

        Catch ex As Exception
          strMessage = "Could not open " & _connectionPortName & " - " & ex.ToString
          Call WriteMessage(strMessage, MessageType.Debug)

          _connected = False
        End Try
      End If

      If _connected = True Then
        '
        ' Enable Caller ID
        '
        If EnableModemCID() = False Then

        End If
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "InitHW()")
    Finally
      UpdateModemConnectionStatus(_lineNumber, _connected)
    End Try

    Return gIOEnabled

  End Function

  ''' <summary>
  ''' Shuts down the com port connection to modem
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ShutdownHW() As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered ShutdownHW() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      '
      ' Close the serial port
      '
      If serialPort.IsOpen Then
        serialPort.Close()
      End If

      _connected = False
      Return True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ShutdownHW()")
      Return False
    Finally
      UpdateModemConnectionStatus(_lineNumber, _connected)
    End Try

  End Function

  ''' <summary>
  ''' Enables modem caller ID
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function EnableModemCID() As Boolean

    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan
    Dim iMillisecondsWaited As Integer
    Dim Attempts As Integer = 0

    Try

      If _connected = False Then
        Return False
      End If

      '
      ' Define modem initilization commands
      '
      Dim strModemInit As String = hs.GetINISetting("Interface", "ModemInit", "ATQ0V1E0~AT+GMM~AT+VCID=1~AT+FCLASS=8~AT-STE=7", gINIFile)
      Dim Commands As String() = strModemInit.Trim.Split("~")

      If Commands.Length = 0 Then
        Call WriteMessage("No modem initialization commands defined.", MessageType.Warning)
        Return False
      End If

      '
      ' Begin modem initilization
      '
      Call WriteMessage("Attemping to enable modem caller ID ...", MessageType.Debug)

      '
      ' Clear out the current modem initilization results
      '
      ModemInit.Clear()

      '
      ' Send initilization AT commands
      '
      For Each strCommand As String In Commands
        '
        ' Add the command to the initilization hashtable
        '
        ModemInit.Add(strCommand, "")

        '
        ' Add the command
        '
        AddCommand(strCommand, False)
      Next

      '
      ' Wait until initialization completes
      '
      dtStartTime = Now
      Do
        Thread.Sleep(50)
        etElapsedTime = Now.Subtract(dtStartTime)
        iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)
      Loop While CommandQueue.Count > 0 And gIOEnabled = True And iMillisecondsWaited < 10000

      '
      ' Check if we successfully enabled Caller ID
      '
      For Each strCommand As String In ModemInit.Keys
        Dim strCmdResult As String = ModemInit(strCommand)

        If strCmdResult.Contains("OK") = False Then
          If Regex.IsMatch(strCommand, "cid", RegexOptions.IgnoreCase) = True Then
            Dim strMessage As String = String.Format("Failed to initialize caller ID.  The command {0} failed.", strCommand)
            Call WriteMessage(strMessage, MessageType.Error)
            Return False
          End If
        End If
      Next

      '
      ' Modem caller ID was successfully enabled
      '
      Call WriteMessage("Modem caller ID was enabled.", MessageType.Informational)

      Return True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "EnableModemCID()")
      Return False
    End Try

  End Function

#End Region

#Region "UltraCID3 Processing"

  ''' <summary>
  ''' Adds command to command buffer for processing
  ''' </summary>
  ''' <param name="strCommand"></param>
  ''' <param name="bForce"></param>
  ''' <remarks></remarks>
  Public Sub AddCommand(ByVal strCommand As String, Optional ByVal bForce As Boolean = False)

    '
    ' bForce may be used to for the same command in the command queue
    '
    If CommandQueue.Contains(strCommand) = False Or bForce = True Then
      CommandQueue.Enqueue(strCommand)
    End If

  End Sub

  ''' <summary>
  ''' Processes commands and waits for the response
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessCommandQueue()

    Dim bAbortThread As Boolean = False
    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan
    Dim iMillisecondsWaited As Integer
    Dim iCmdAttempt As Integer = 0
    Dim strCommand As String = ""
    Dim iMaxWaitTime As Double = 0

    Try

      While bAbortThread = False
        '
        ' Process commands in command queue
        '
        While CommandQueue.Count > 0 And gIOEnabled = True
          '
          ' Set the command response we are waiting for
          '
          strCommand = CommandQueue.Peek
          iMaxWaitTime = 0

          '
          ' Determine if we need to modify the strCmdWait
          '
          If Regex.IsMatch(strCommand, "^AT(A|H)") = True Then
            '
            ' The response should be in upper case
            '
            _cmdWait = "OK|ERROR|CONNECT|NO "
            iMaxWaitTime = 10.0
          Else
            '
            ' The response should be in upper case
            '
            _cmdWait = "OK"
            iMaxWaitTime = 2.0
          End If

          '
          ' Increment the counter
          '
          iCmdAttempt += 1
          WriteMessage(String.Format("Sending command: '{0}' to modem, attempt # {1}", strCommand, iCmdAttempt), MessageType.Debug)
          SendToModem(strCommand)

          '
          ' Determine if we need to wait for a response
          '
          If iMaxWaitTime > 0 And _cmdWait.Length > 0 Then
            '
            ' A response to our command is expected, so lets wait for it
            '
            WriteMessage(String.Format("Waiting for the modem to respond with '{0}' for up to {1} seconds...", _cmdWait, iMaxWaitTime), MessageType.Debug)

            '
            ' Keep track of when we started waiting for the response
            '
            dtStartTime = Now

            '
            '  Wait for the proper response to come back, or the maximum wait time
            '
            Do
              '
              ' Sleep this thread for 50ms giving the receive function time to get the response
              '
              Thread.Sleep(50)

              '
              ' Find out how long we have been waiting in total
              '
              etElapsedTime = Now.Subtract(dtStartTime)
              iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)

              '
              ' Loop until the expected command was received (strCmdWait is cleared) or we ran past the maximum wait time
              '
            Loop Until _cmdWait.Length = 0 Or iMillisecondsWaited > iMaxWaitTime * 1000 ' Now abort if the command was recieved or we ran out of time

            WriteMessage(String.Format("Waited {0} milliseconds for the command response.", iMillisecondsWaited), MessageType.Debug)

            If _cmdWait.Length > 0 Or iCmdAttempt > MAX_ATTEMPTS Then
              '
              ' Command failed, so lets stop trying to send this commmand
              '
              If _cmdWait <> "" Then
                WriteMessage(String.Format("No response/improper response from modem command '{0}'", strCommand), MessageType.Warning)
              End If
              _cmdWait = String.Empty

              '
              ' Only Dequeue the command if we have tried more than MAX_ATTEMPTS times
              '
              If iCmdAttempt > MAX_ATTEMPTS Then
                CommandQueue.Dequeue()
                _cmdWait = String.Empty
                iCmdAttempt = 0
              End If
            Else
              CommandQueue.Dequeue()
              iCmdAttempt = 0
            End If
          Else
            '
            ' No response expected, so remove command from queue
            '
            WriteMessage(String.Format("Command {0} does not produce a result.", strCommand), MessageType.Debug)
            CommandQueue.Dequeue()
            _cmdWait = String.Empty
            iCmdAttempt = 0
          End If

        End While ' Done with all commands in queue

        '
        ' Give up some time to allow the main thread to populate the command queue with more commands
        '
        Thread.Sleep(50)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessCommandQueue thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessCommandQueue()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessCommandQueue terminated."), MessageType.Debug)

    End Try

  End Sub

  ''' <summary>
  ''' Sends command to modem
  ''' </summary>
  ''' <param name="strDataToSend"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SendToModem(ByVal strDataToSend As String) As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered SendToModem() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      If serialPort.IsOpen = True Then
        strMessage = "Sending " & strDataToSend & " to modem serial port ..."
        Call WriteMessage(strMessage, MessageType.Debug)

        serialPort.Write(strDataToSend & vbCrLf)
      End If
      Return serialPort.IsOpen

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SendToModem()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Event handler for Com Port data received
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub DataReceived(ByVal sender As Object,
                           ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) _
                           Handles serialPort.DataReceived

    Static Str As New StringBuilder
    'Static Dat As Char
    'Static By As Integer

    Dim strMessage As String = ""

    strMessage = "Entered Serial DataReceived() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      'While serialPort.BytesToRead > 0

      '  By = serialPort.ReadByte()
      '  Dat = Chr(By)
      '  'WriteMessage(Dat, MSG_DEBUG)

      '  If Dat = vbLf Then
      '    If Str.Length > 0 Then
      '      '
      '      ' Process data received
      '      '
      '      Dim strDataRec As String = Str.ToString
      '      If Asc(strDataRec) <> 13 Then
      '        ProcessReceived(strDataRec.Replace(Chr(16), "").Trim)
      '      End If
      '    End If
      '    Str.Length = 0
      '  Else
      '    Str.Append(Dat)
      '  End If

      '  Thread.Sleep(0)
      'End While

      'If Str.Length > 0 Then
      '  Dim strDataRec As String = Str.ToString
      '  If Asc(strDataRec) <> 13 Then
      '    ProcessReceived(strDataRec.Replace(Chr(16), "").Trim)
      '  End If
      'End If

      Do

        Str.Length = 0

        Dim strDataRec As String = ""
        Try
          strDataRec = serialPort.ReadLine()
        Catch ex As Exception
          strDataRec = serialPort.ReadExisting
        End Try

        '
        ' Process data received
        '
        If Asc(strDataRec) <> 13 Then
          ProcessReceived(strDataRec.Replace(Chr(16), "").Trim)
        End If

      Loop While serialPort.BytesToRead > 0

      strMessage = "Exited Serial DataReceived() function."
      Call WriteMessage(strMessage, MessageType.Debug)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "DataReceived()")
    End Try

  End Sub

  ''' <summary>
  ''' Processes a received data string
  ''' </summary>
  ''' <param name="strDataRec"></param>
  ''' <remarks></remarks>
  Public Sub ProcessReceived(ByVal strDataRec As String)

    Static LastCID As New Hashtable
    Static strAnnounceMsg As String = String.Empty
    Static iAnnounced As Integer = 0

    Dim strMessage As String = ""

    strMessage = "Entered ProcessReceived() function with a string [" & strDataRec & "]"
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      If (Regex.IsMatch(strDataRec, "^(OK|CONNECT|ERROR|BUSY|DELAYED|BLACKLIST|NO )") = True) Then
        '
        ' Command response
        '
        Dim strCommand As String = CommandQueue.Peek

        If ModemInit.ContainsKey(strCommand) Then
          '
          ' Indicate the response from the modem init command
          '
          ModemInit(strCommand) = strDataRec
          _cmdWait = String.Empty

          strMessage = String.Format("Modem init command {0} returned {1}.", strCommand, strDataRec)
          WriteMessage(strMessage, MessageType.Debug)
        Else
          _cmdWait = String.Empty

          strMessage = String.Format("Modem command {0} returned {1}.", strCommand, strDataRec)
          WriteMessage(strMessage, MessageType.Debug)
        End If

      ElseIf (Regex.IsMatch(strDataRec, "^(R|RING).?$") = True) Then
        '
        ' Ring action
        '
        If LastCID.Count And LastCID.ContainsKey("NMBR") Then
          ProcessCallerCID(LastCID, strAnnounceMsg)
        End If
        LastCID.Clear()

        _ringCount += 1

        Dim iTicks As Long = DateTime.Now.Ticks
        Dim iTickDiff As Long = (iTicks - _lastRingTicks) / 10000000.0
        _lastRingTicks = iTicks

        '
        ' Check if this ring is a new call
        '
        If (iTickDiff) > 15 Then
          strMessage = "New incoming call detected ..."
          _callCount += 1
          _callRings = 1
          iAnnounced = 0

          strAnnounceMsg = ""
        Else
          _callRings += 1
          strMessage = "Ring of existing call detected ..."
        End If
        WriteMessage(strMessage, MessageType.Informational)

        '
        ' Update the last caller ring HomeSeer device
        '
        Dim dv_addr As String = String.Format("{0}-{1}{2}", IFACE_NAME, "Ring", _lineNumber)
        SetDeviceValue(dv_addr, _callRings.ToString)

        Dim arrTriggers() As String = {_callRings.ToString, "0"}

        For Each strActions As String In arrTriggers
          '
          ' Check TelephoneRing trigger
          '
          Dim strTrigger As String = String.Format("{0},{1},{2}", _lineNumber, "TelephoneRing", strActions)
          hspi_plugin.CheckTrigger(IFACE_NAME, CIDTriggers.TelephoneRing, -1, strTrigger)
        Next

        '
        ' Announce caller on ring
        '
        'If _callRings > 1 And strAnnounceMsg.Length > 0 Then
        '  Dim strAnnounceRings As String = hs.GetINISetting("Announce", "Ring", "1", gINIFile)
        '  If strAnnounceRings = "*" Then
        '    hs.Speak(strAnnounceMsg, False)
        '  End If
        'End If

      ElseIf (Regex.IsMatch(strDataRec, "(DATE|TIME|NAME|NMBR|DDN|MESG|TYPE)\s?=\s?(.+)")) Then
        '
        ' CID data
        '
        Dim colMatches As MatchCollection = Regex.Matches(strDataRec, "(?<KEY>(DATE|TIME|NAME|NMBR|DDN|MESG|TYPE))\s?=\s?(?<VALUE>.+)")
        If colMatches.Count > 0 Then

          Dim strKeyName As String = colMatches.Item(0).Groups("KEY").Value
          Dim strKeyValue As String = colMatches.Item(0).Groups("VALUE").Value

          '
          ' Add work-around for TFM-760U modem
          '
          If strKeyName = "DDN" Then strKeyName = "NMBR"

          '
          ' Add work-around for TFM-560U modem
          ' 
          If strKeyName = "NMBR" And strKeyValue.Contains("OUT OF AREA") Then
            If LastCID.ContainsKey("NAME") = False Then
              strKeyName = "NAME"
            End If
          End If

          '
          ' Add the caller ID data to the LastCID hash
          '
          If LastCID.ContainsKey(strKeyName) = False Then
            LastCID.Add(strKeyName, strKeyValue)
          End If

          '
          ' Check if CID data is complete
          '
          If LastCID.Count = 4 Then
            ProcessCallerCID(LastCID, strAnnounceMsg)
          End If

        End If

      ElseIf (Regex.IsMatch(strDataRec, "^(h|H).?$")) Then
        '
        ' h = On Hook, H = Off Hook
        ' 
        Dim strAction As String = IIf(String.Compare(strDataRec, "H", False), "On Hook", "Off Hook")
        Dim iDeviceValue As Integer = IIf(String.Compare(strDataRec, "H", False), 0, 1)

        Dim arrTriggers() As String = {strAction, "Any"}
        For Each strActions As String In arrTriggers
          '
          ' Check LocalHandset trigger
          '
          Dim strTrigger As String = String.Format("{0},{1},{2}", _lineNumber, "LocalHandset", strActions)
          hspi_plugin.CheckTrigger(IFACE_NAME, CIDTriggers.LocalHandset, -1, strTrigger)

        Next

      ElseIf (Regex.IsMatch(strDataRec, "^(p|P).?$")) Then
        '
        ' p = Line voltage increase, P = Line voltage decrease (extension pickup)
        ' 
        Dim strAction As String = IIf(String.Compare(strDataRec, "P", False), "On Hook", "Off Hook")
        Dim iDeviceValue As Integer = IIf(String.Compare(strDataRec, "P", False), 0, 1)

        Dim arrTriggers() As String = {strAction, "Any"}
        For Each strActions As String In arrTriggers
          '
          ' Check PhoneExtension trigger
          '
          Dim strTrigger As String = String.Format("{0},{1},{2}", _lineNumber, "PhoneExtension", strActions)
          hspi_plugin.CheckTrigger(IFACE_NAME, CIDTriggers.PhoneExtension, -1, strTrigger)

        Next

        '
        ' Update the HomeSeer device
        '
        Dim dv_addr As String = String.Format("{0}-{1}{2}", IFACE_NAME, "Extn", _lineNumber)
        SetDeviceValue(dv_addr, iDeviceValue)

      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ProcessReceived()")
    End Try

  End Sub

  ''' <summary>
  ''' Process Caller CID
  ''' </summary>
  ''' <param name="LastCID"></param>
  ''' <param name="strAnnounceMsg"></param>
  ''' <remarks></remarks>
  Private Sub ProcessCallerCID(ByVal LastCID As Hashtable, ByRef strAnnounceMsg As String)

    Dim strMessage As String = String.Empty
    Dim dv_addr As String = String.Empty

    Try
      '
      ' Provide default name if not available in CID
      '
      If LastCID.ContainsKey("NAME") = False Then
        LastCID.Add("NAME", "UNKNOWN")
      End If

      '
      ' Remove the leading 1
      '
      If LastCID.ContainsKey("NMBR") = True Then
        Dim nmbr As String = LastCID("NMBR")
        If nmbr.StartsWith("1") AndAlso nmbr.Length > 9 Then
          LastCID("NMBR") = nmbr.TrimStart("1")
        End If
      End If

      '
      ' Get/Format caller information
      '
      _callerAttr = hspi_database.GetCallerAttributes(LastCID("NMBR"))
      _callerName = hspi_database.GetCallerAlias(LastCID("NMBR"), LastCID("NAME"))
      _callerNmbr = FormatNmbr(LastCID("NMBR"))

      '
      ' Update the HomeSeer devices
      '
      dv_addr = String.Format("{0}-{1}{2}", IFACE_NAME, "Name", _lineNumber)
      SetDeviceString(dv_addr, _callerName)

      dv_addr = String.Format("{0}-{1}{2}", IFACE_NAME, "Nmbr", _lineNumber)
      SetDeviceString(dv_addr, _callerNmbr)

      '
      ' Write to HomeSeer log
      '
      strMessage = String.Format("Incoming call from {0} {1}", _callerName, _callerNmbr)
      WriteMessage(strMessage, MessageType.Informational)

      '
      ' Process incoming call
      '
      Dim ts As Integer = ConvertDateTimeToEpoch(DateTime.Now)
      hspi_database.ProcessIncomingCall(ts, LastCID("NMBR"), LastCID("NAME"), _callerAttr)

      '
      ' Get updated caller attributes
      '
      Dim firstTimeCaller As Boolean = IIf(_callerAttr = -1, True, False)
      If _callerAttr = -1 Then
        _callerAttr = hspi_database.GetCallerAttributes(LastCID("NMBR"))
      End If

      '
      ' Check for Caller Number HomeSeer Triggers
      '
      Dim arrTriggers() As String = {LastCID("NMBR"), "Any", "FirstTime"}
      For Each strTriggers As String In arrTriggers
        If strTriggers = "FirstTime" And firstTimeCaller = False Then Continue For
        '
        ' Check IncomingCall trigger
        '
        Dim strTrigger As String = String.Format("{0},{1},{2}", _lineNumber, "IncomingCall", strTriggers)
        hspi_plugin.CheckTrigger(IFACE_NAME, CIDTriggers.IncomingCall, -1, strTrigger)
      Next

      '
      ' Format the Caller Attributes
      '
      Dim strCallerAttributes As String = FormatCallerAttrNames(_callerAttr)

      dv_addr = String.Format("{0}-{1}{2}", IFACE_NAME, "Attr", _lineNumber)
      SetDeviceValue(dv_addr, -1)
      SetDeviceString(dv_addr, strCallerAttributes)

      '
      ' Check for Caller Attribute HomeSeer Triggers
      '
      Dim names As String() = System.Enum.GetNames(GetType(CallerAttrs))
      For i As Integer = 0 To names.Length - 1
        Dim iAttr As CallerAttrs = System.Enum.Parse(GetType(CallerAttrs), names(i))
        If _callerAttr And iAttr Then
          Dim strTriggers As String = names(i)

          '
          ' Check IncomingCall trigger
          '
          Dim strTrigger As String = String.Format("{0},{1},{2}", _lineNumber, "IncomingCall", strTriggers)
          hspi_plugin.CheckTrigger(IFACE_NAME, CIDTriggers.IncomingCall, -1, strTrigger)
        End If
      Next

      '
      ' Check if we should block the caller
      '
      If _callerAttr > 0 And (_callerAttr And CallerAttrs.Block) Then
        DropCaller()
        _dropCount += 1
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ProcessCallerCID()")
    Finally
      LastCID.Clear()
    End Try

  End Sub

  ''' <summary>
  ''' E-Mails last caller
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub EmailLastCaller()

    Try

      Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "smtp_from", "")
      Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "smtp_to", "")

      Dim strEmailTo As String = hs.GetINISetting("EmailNotification", "EmailRcptTo", strEmailRcptTo, gINIFile)
      Dim strEmailFrom As String = hs.GetINISetting("EmailNotification", "EmailFrom", strEmailFromDefault, gINIFile)
      Dim strEmailSubject As String = hs.GetINISetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT, gINIFile)
      Dim strEmailBody As String = hs.GetINISetting("EmailNotification", "EmailBody", EMAIL_BODY_TEMPLATE, gINIFile)

      If Regex.IsMatch(strEmailFrom, ".+@.+") = False Then
        Throw New Exception("Unable to send last caller e-mail notification because the sender is not a valid e-mail address.")
      ElseIf Regex.IsMatch(strEmailTo, ".+@.+") = False Then
        Throw New Exception("Unable to send last caller e-mail notification because the recipient is not a valid e-mail address.")
      End If

      Dim strCallerName As String = hs.DeviceString(gBaseCode & DEV_LAST_CALLER_NAME)
      Dim strCallerNmbr As String = hs.DeviceString(gBaseCode & DEV_LAST_CALLER_NMBR)
      Dim strCallerTime As String = hs.DeviceLastChange(gBaseCode & DEV_LAST_CALLER_NMBR)

      strEmailBody = strEmailBody.Replace("$name", strCallerName)
      strEmailBody = strEmailBody.Replace("$nmbr", strCallerNmbr)
      strEmailBody = strEmailBody.Replace("$datetime", strCallerTime)
      strEmailBody = strEmailBody.Replace("$rings", _callRings.ToString)

      strEmailSubject = strEmailSubject.Replace("$name", strCallerName)
      strEmailSubject = strEmailSubject.Replace("$nmbr", strCallerNmbr)
      strEmailSubject = strEmailSubject.Replace("$datetime", strCallerTime)
      strEmailSubject = strEmailSubject.Replace("$rings", _callRings.ToString)

      Dim List() As String = hs.GetPluginsList()
      If List.Contains("UltraSMTP3:") = True Then
        '
        ' Send e-mail using UltraSMTP3
        '
        hs.PluginFunction("UltraSMTP3", "", "SendMail", New Object() {strEmailTo, strEmailSubject, strEmailBody, Nothing})
      Else
        '
        ' Send e-mail using HomeSeer
        '
        hs.SendEmail(strEmailTo, strEmailFrom, "", "", strEmailSubject, strEmailBody, "")
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EmailLastCaller()")
    End Try

  End Sub

  ''' <summary>
  ''' Drops the current caller
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub DropCaller()

    Try
      '
      ' Write debug message
      '
      WriteMessage("Drop Caller action activated ...", MessageType.Warning)

      '
      ' Define drop caller commands
      '
      Dim strDropCaller As String = hs.GetINISetting("Interface", "DropCaller", "AT+VLS=5~ATH~AT+FCLASS=8", gINIFile)
      Dim Commands As String() = strDropCaller.Trim.Split("~")

      If Commands.Length = 0 Then
        Call WriteMessage("No drop caller commands defined.", MessageType.Warning)
        Exit Sub
      End If

      '
      ' Send initilization AT commands
      '
      For Each strCommand As String In Commands
        '
        ' Add the command
        '
        AddCommand(strCommand, False)
      Next

    Catch ex As Exception
      '
      ' Ignore error
      '
    End Try

  End Sub

#End Region

  ''' <summary>
  ''' Gets the modem status
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetModemStatus() As String

    If serialPort.IsOpen Then
      Return serialPort.PortName
    Else
      Return "Not connected"
    End If

  End Function

  ''' <summary>
  ''' Returns various statistics
  ''' </summary>
  ''' <param name="StatisticsType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStatistics(ByVal StatisticsType As String) As Integer

    Select Case StatisticsType
      Case "CallCount"
        Return _callCount
      Case "RingCount"
        Return _ringCount
      Case "DropCount"
        Return _dropCount
      Case "DBInsSuccess"
        Return gDBInsertSuccess
      Case "DBInsFailure"
        Return gDBInsertFailure
      Case Else
        Return 0
    End Select

  End Function

End Class
