Imports HomeSeerAPI

Module hspi_devices

  Public DEV_STATUS As Byte = 1

  Dim bCreateRootDevice = True
  ''' <summary>
  ''' Update the list of monitored devices
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub UpdateMiscDeviceSettings(ByVal bDeviceNoLog As Boolean)

    Dim dv As Scheduler.Classes.DeviceClass
    Dim DevEnum As Scheduler.Classes.clsDeviceEnumeration

    Dim strMessage As String = ""

    strMessage = "Entered UpdateMiscDeviceSettings() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Go through devices to see if we have one assigned to our plug-in
      '
      DevEnum = hs.GetDeviceEnumerator

      If Not DevEnum Is Nothing Then

        Do While Not DevEnum.Finished
          dv = DevEnum.GetNext
          If dv Is Nothing Then Continue Do
          If dv.Interface(Nothing) IsNot Nothing Then
            If dv.Interface(Nothing) = IFACE_NAME Then
              '
              ' We found our device, so process based on device type
              '
              Dim dv_type As String = dv.Device_Type_String(hs)

              '
              ' Set options based on root or child device
              '
              If dv.Relationship(hs) = Enums.eRelationship.Parent_Root Then
                dv.MISC_Set(hs, Enums.dvMISC.NO_STATUS_TRIGGER)               ' When set, the device status values will Not appear in the device change trigger
              ElseIf dv.Relationship(hs) = Enums.eRelationship.Child Then
                dv.MISC_Clear(hs, Enums.dvMISC.NO_STATUS_TRIGGER)             ' When set, the device status values will Not appear in the device change trigger

                Select Case dv_type
                  Case "Modem Connection", "Last Caller Attributes"
                    dv.MISC_Clear(hs, Enums.dvMISC.STATUS_ONLY)               ' When set, the device cannot be controlled
                    dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)        ' When set, this device is included in the voice recognition context for device commands
                  Case "Something Else"
                    dv.MISC_Clear(hs, Enums.dvMISC.STATUS_ONLY)               ' When set, the device cannot be controlled
                    dv.MISC_Set(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)          ' When set, this device is included in the voice recognition context for device commands
                  Case Else
                    dv.MISC_Set(hs, Enums.dvMISC.STATUS_ONLY)                 ' When set, the device cannot be controlled
                    dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)        ' When set, this device is included in the voice recognition context for device commands
                End Select
              End If

              '
              ' Apply Logging Options
              '
              If bDeviceNoLog = False Then
                dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)                          ' When set, no logging to the log for this device
              Else
                dv.MISC_Clear(hs, Enums.dvMISC.NO_LOG)                        ' When set, no logging to the log for this device
              End If

              '
              ' This property indicates (when True) that the device supports the retrieval of its status on-demand through the "Poll" feature on the device utility page.
              '
              dv.Status_Support(hs) = False

              '
              ' If an event or device was modified by a script, this function should be called to update HomeSeer with the changes.
              '
              hs.SaveEventsDevices()

            End If
          End If
        Loop
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdateMiscDeviceSettings()")
    End Try

    DevEnum = Nothing
    dv = Nothing

  End Sub

  ''' <summary>
  ''' Update the list of monitored devices
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub UpdateRootDevices()

    Dim dv As Scheduler.Classes.DeviceClass
    Dim DevEnum As Scheduler.Classes.clsDeviceEnumeration

    Dim strMessage As String = ""

    strMessage = "Entered UpdateMonitoredDevices() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Go through devices to see if we have one assigned to our plug-in
      '
      DevEnum = hs.GetDeviceEnumerator

      If Not DevEnum Is Nothing Then

        Do While Not DevEnum.Finished
          dv = DevEnum.GetNext
          If dv Is Nothing Then Continue Do
          If dv.Interface(Nothing) IsNot Nothing Then

            If dv.Interface(Nothing) = IFACE_NAME Then
              '
              ' We found our device, so process based on device type
              '
              Dim dv_ref As Integer = dv.Ref(hs)
              Dim dv_addr As String = dv.Address(hs)
              Dim dv_name As String = dv.Name(hs)
              Dim dv_type As String = dv.Device_Type_String(hs)

              If dv.Relationship(hs) = Enums.eRelationship.Not_Set Then

                Select Case dv_addr
                  Case "UltraCID3:1" : dv_addr = "UltraCID3-Conn"
                  Case "UltraCID3:3" : dv_addr = "UltraCID3-Extn"
                  Case "UltraCID3:4" : dv_addr = "UltraCID3-Name"
                  Case "UltraCID3:5" : dv_addr = "UltraCID3-Nmbr"
                  Case "UltraCID3:6" : dv_addr = "UltraCID3-Attr"
                  Case "UltraCID3:7" : dv_addr = "UltraCID3-Ring"
                  Case "UltraCID3-Conn" : dv_addr = "UltraCID3-Conn1"
                  Case "UltraCID3-Extn" : dv_addr = "UltraCID3-Extn1"
                  Case "UltraCID3-Name" : dv_addr = "UltraCID3-Name1"
                  Case "UltraCID3-Nmbr" : dv_addr = "UltraCID3-Nmbr1"
                  Case "UltraCID3-Attr" : dv_addr = "UltraCID3-Attr1"
                  Case "UltraCID3-Ring" : dv_addr = "UltraCID3-Ring1"
                End Select

                dv.Address(hs) = dv_addr
                hs.SaveEventsDevices()

                If bCreateRootDevice = True Then
                  '
                  ' Make this device a child of the root
                  '
                  dv.AssociatedDevice_ClearAll(hs)
                  Dim dvp_ref As Integer = CreateRootDevice("Plugin", IFACE_NAME, dv_ref)
                  If dvp_ref > 0 Then
                    dv.AssociatedDevice_Add(hs, dvp_ref)
                  End If
                  dv.Relationship(hs) = Enums.eRelationship.Child
                  hs.SaveEventsDevices()
                End If

              End If

            End If
          End If
        Loop

      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdateRootDevices()")
    End Try

    DevEnum = Nothing
    dv = Nothing

  End Sub

  ''' <summary>
  ''' Create the HomeSeer Root Device
  ''' </summary>
  ''' <param name="strRootId"></param>
  ''' <param name="strRootName"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateRootDevice(ByVal strRootId As String, _
                                   ByVal strRootName As String, _
                                   ByVal dv_ref_child As Integer) As Integer

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Set the local variables
      '
      If strRootId = "Plugin" Then
        dv_name = "UltraCID Plugin"
        dv_addr = String.Format("{0}-Root", strRootName.Replace(" ", "-"))
        dv_type = dv_name
      Else
        dv_name = strRootId
        dv_addr = String.Format("{0}-Root", strRootId, strRootName.Replace(" ", "-"))
        dv_type = String.Concat(IFACE_NAME, " Modem")
      End If

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} root device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} root device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = IIf(strRootId = "Plugin", "Plug-ins", dv_type)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this a parent root device
      '
      dv.Relationship(hs) = Enums.eRelationship.Parent_Root
      dv.AssociatedDevice_Add(hs, dv_ref_child)

      Dim image As String = "device_root.png"

      Dim VSPair As VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Root"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, image)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.STATUS_ONLY)           ' When set, the device cannot be controlled
      dv.MISC_Set(hs, Enums.dvMISC.NO_STATUS_TRIGGER)     ' When set, the device status values will Not appear in the device change trigger
      dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)  ' When set, this device is included in the voice recognition context for device commands

      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)         ' When set, device control options will be displayed
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)              ' When set, no logging to the log for this device
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception

    End Try

    Return dv_ref

  End Function


  ''' <summary>
  ''' Function to initilize our plug-ins devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitPluginDevices() As String

    Dim strMessage As String = ""

    WriteMessage("Entered InitPluginDevices() function.", MessageType.Debug)

    Try
      For lineNumber As Integer = 1 To gModemInterfaces
        Dim Devices As Byte() = {DEV_HARDWARE_MODEM, DEV_EXTENSION_STATE, DEV_LAST_CALLER_NAME, DEV_LAST_CALLER_NMBR, DEV_LAST_CALLER_ATTR, DEV_LAST_CALLER_RING}
        For Each dev_cod As Byte In Devices
          Dim strResult As String = CreatePluginDevice(IFACE_NAME, dev_cod, lineNumber)
          If strResult.Length > 0 Then Return strResult
        Next
      Next

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "InitPluginDevices()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Subroutine to create a HomeSeer device
  ''' </summary>
  ''' <param name="base_code"></param>
  ''' <param name="dev_code"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreatePluginDevice(ByVal base_code As String, ByVal dev_code As String, ByVal lineNumber As Integer) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim DevicePairs As New List(Of hspi_device_pairs)

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_root_name As String = String.Empty
    Dim dv_root_addr As String = String.Empty

    Dim dv_name As String = String.Empty
    Dim dv_type As String = String.Empty
    Dim dv_addr As String = String.Empty

    Dim DeviceShowValues As Boolean = False

    Try

      Select Case dev_code
        Case DEV_HARDWARE_MODEM.ToString
          '
          ' Create the Monitoring State device
          '
          dv_root_name = String.Format("{0} Line #{1}", "Modem", lineNumber)
          dv_root_addr = String.Format("{0}-Line{1}", base_code, lineNumber)

          dv_name = "Modem Connection"
          dv_type = "Modem Connection"
          dv_addr = String.Concat(base_code, "-Conn", lineNumber)

        Case DEV_EXTENSION_STATE.ToString
          '
          ' Create the Phone Extension device
          '
          dv_root_name = String.Format("{0} Line #{1}", "Modem", lineNumber)
          dv_root_addr = String.Format("{0}-Line{1}", base_code, lineNumber)

          dv_name = "Phone Extension"
          dv_type = "Phone Extension"
          dv_addr = String.Concat(base_code, "-Extn", lineNumber)

        Case DEV_LAST_CALLER_NAME.ToString
          '
          ' Create the Last Caller Name device
          '
          dv_root_name = String.Format("{0} Line #{1}", "Modem", lineNumber)
          dv_root_addr = String.Format("{0}-Line{1}", base_code, lineNumber)

          dv_name = "Last Caller Name"
          dv_type = "Last Caller Name"
          dv_addr = String.Concat(base_code, "-Name", lineNumber)

        Case DEV_LAST_CALLER_NMBR.ToString
          '
          ' Create the Last Caller Number device
          '
          dv_root_name = String.Format("{0} Line #{1}", "Modem", lineNumber)
          dv_root_addr = String.Format("{0}-Line{1}", base_code, lineNumber)

          dv_name = "Last Caller Number"
          dv_type = "Last Caller Number"
          dv_addr = String.Concat(base_code, "-Nmbr", lineNumber)

        Case DEV_LAST_CALLER_ATTR.ToString
          '
          ' Create Last Caller Attributes device
          '
          dv_root_name = String.Format("{0} Line #{1}", "Modem", lineNumber)
          dv_root_addr = String.Format("{0}-Line{1}", base_code, lineNumber)

          dv_name = "Last Caller Attributes"
          dv_type = "Last Caller Attributes"
          dv_addr = String.Concat(base_code, "-Attr", lineNumber)

        Case DEV_LAST_CALLER_RING.ToString
          '
          ' Create Last Caller Ring device
          '
          dv_root_name = String.Format("{0} Line #{1}", "Modem", lineNumber)
          dv_root_addr = String.Format("{0}-Line{1}", base_code, lineNumber)

          dv_name = "Last Caller Rings"
          dv_type = "Last Caller Rings"
          dv_addr = String.Concat(base_code, "-Ring", lineNumber)

        Case Else
          Throw New Exception(String.Format("Unable to create plug-in device for unsupported device name: {0}", dv_name))
      End Select

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = "Plug-ins"
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      If dv.Relationship(hs) <> Enums.eRelationship.Child Then

        If bCreateRootDevice = True Then
          dv.AssociatedDevice_ClearAll(hs)
          Dim dvp_ref As Integer = CreateRootDevice(dv_root_name, dv_root_addr, dv_ref)
          If dvp_ref > 0 Then
            dv.AssociatedDevice_Add(hs, dvp_ref)
          End If
          dv.Relationship(hs) = Enums.eRelationship.Child
        End If

      End If

      '
      ' Clear the value status pairs
      '
      hs.DeviceVSP_ClearAll(dv_ref, True)
      hs.DeviceVGP_ClearAll(dv_ref, True)
      hs.SaveEventsDevices()

      '
      ' Update the last change date
      ' 
      If bDeviceExists = False Then
        dv.Last_Change(hs) = DateTime.Now
      End If

      Dim VSPair As VSPair
      Dim VGPair As VGPair

      Select Case dv_type
        Case "Modem Connection"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -3
          VSPair.Status = ""
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -2
          VSPair.Status = "Disconnect"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -1
          VSPair.Status = "Reconnect"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "Disconnected"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "Connected"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VGPairs
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = -3
          VGPair.RangeEnd = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "modem_disconnected.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 1
          VGPair.RangeEnd = 1
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "modem_connected.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Dim dev_status As Integer = 0
          hs.SetDeviceValueByRef(dv_ref, dev_status, False)

          DeviceShowValues = True

        Case "Database"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -2
          VSPair.Status = "Close"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -1
          VSPair.Status = "Open"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "Closed"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "Open"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          DeviceShowValues = True

        Case "Phone Extension"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "On Hook"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "Off Hook"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VGPairs
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "on_hook.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = 1
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "off_hook.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          DeviceShowValues = False

        Case "Last Caller Rings"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 9
          'VSPair.Status = "Rings"
          VSPair.RangeStatusSuffix = " Rings"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          For iCount As Integer = 0 To 9
            '
            ' Add VGPairs
            '
            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = iCount
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, String.Format("count_{0}.png", iCount))
            hs.DeviceVGP_AddPair(dv_ref, VGPair)
          Next

          DeviceShowValues = False

        Case "Last Caller Name", "Last Caller Number"
          '
          ' Add VGPairs
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "caller_contact.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          If bDeviceExists = False Then
            SetDeviceString(dv_addr, "No caller")
          End If

          DeviceShowValues = False

        Case "Last Caller Attributes"
          '
          ' Add VGPairs
          '
          DevicePairs.Add(New hspi_device_pairs(-1, "", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Removal All", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(1, "Add Block Caller Attribute", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(2, "Add Telemarketer Attribute", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(4, "Add Announce Attribute", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(8, "Add Business Attribute", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(16, "Add Family Attribute", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(32, "Add Friends Attribute", "caller_attributes.png", HomeSeerAPI.ePairStatusControl.Control))

          For Each Pair As hspi_device_pairs In DevicePairs

            VSPair = New VSPair(Pair.Type)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = Pair.Value
            VSPair.Status = Pair.Status
            VSPair.Render = Enums.CAPIControlType.Values
            hs.DeviceVSP_AddPair(dv_ref, VSPair)

            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = Pair.Value
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
            hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Next

          '
          ' Set the default device value
          '
          If bDeviceExists = False Then
            SetDeviceString(dv_addr, "No caller")
          End If
          SetDeviceValue(dv_addr, -1)

          DeviceShowValues = True

      End Select

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Clear(hs, Enums.dvMISC.STATUS_ONLY)         ' When set, the device cannot be controlled
      dv.MISC_Clear(hs, Enums.dvMISC.NO_STATUS_TRIGGER)   ' When set, the device status values will Not appear in the device change trigger
      dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)  ' When set, this device is included in the voice recognition context for device commands

      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)         ' When set, device control options will be displayed
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)              ' When set, no logging to the log for this device
      End If

      '
      ' This property indicates (when True) that the device supports the retrieval of its status on-demand through the "Poll" feature on the device utility page.
      '
      dv.Status_Support(hs) = False

      '
      ' If an event or device was modified by a script, this function should be called to update HomeSeer with the changes.
      '
      hs.SaveEventsDevices()

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreatePluinDevice()")
      Return "Failed to create HomeSeer device due to error."
    End Try

  End Function

  ''' <summary>
  ''' Updates the Modem Connection Device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateModemConnectionStatus(ByVal lineNumber As Integer, ByVal modemConnected As Boolean)

    Try

      Dim dv_addr = String.Format("{0}-{1}{2}", IFACE_NAME, "Conn", lineNumber)
      Dim dv_value As Integer = IIf(modemConnected = True, 1, 0)

      SetDeviceValue(dv_addr, dv_value)
    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdateModemConnectionStatus()")
    End Try

  End Function

  ''' <summary>
  ''' Locates device by device code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_value"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal dv_addr As String, _
                            ByVal dv_value As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_value), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        If IsNumeric(dv_value) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblSensorValue As Double = Double.Parse(dv_value)

          If dblDeviceValue <> dblSensorValue Then
            hs.SetDeviceValueByRef(dv_ref, dblSensorValue, True)
          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Sets the HomeSeer device string
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_string"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceString(ByVal dv_addr As String, _
                             ByVal dv_string As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_string), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        hs.SetDeviceString(dv_ref, dv_string, True)

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceString()")

    End Try

  End Sub

End Module
