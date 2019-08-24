Imports System.Text
Imports System.Web
Imports Scheduler
Imports System.Collections.Specialized
Imports System.Text.RegularExpressions

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder

      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)
      End If

      Dim Header As New StringBuilder
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultracid3/css/jquery.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultracid3/css/dataTables.editor.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultracid3/css/editor.dataTables.min.css"" rel=""stylesheet"" />")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultracid3/js/jquery.dataTables.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultracid3/js/dataTables.editor.min.js""></script>")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultracid3/js/hspi_ultracid3_utility.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultracid3/js/hspi_ultracid3_caller_log.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultracid3/js/hspi_ultracid3_caller_details.js""></script>")
      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      Me.RefreshIntervalMilliSeconds = 3000
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus(False) & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Caller Log"
      tab.tabDIVID = "tabCallerLog"
      tab.tabContent = "<div id='divCallerLog'>" & BuildTabCallerLog(False) & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Caller Details"
      tab.tabDIVID = "tabCallerDetails"
      tab.tabContent = "<div id='divCallerDetails'>" & BuildTabCallerDetails() & "</div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("    <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Database Statistics </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Inserts:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics(1, "DBInsSuccess")).ToString("N0"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Failures:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics(1, "DBInsFailure")).ToString("N0"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Size:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", GetDatabaseSize("DBConnectionMain"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      For lineNumber As Integer = 1 To gModemInterfaces
        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td>")
        stb.AppendLine("   <fieldset>")
        stb.AppendFormat("    <legend> Line #{0} Modem Status </legend>", lineNumber.ToString)
        stb.AppendLine("    <table style=""width: 100%"">")
        stb.AppendLine("     <tr>")
        stb.AppendLine("      <td style=""width: 20%""><strong>Connection:</strong></td>")
        stb.AppendFormat("    <td align=""right"">{0}</td>", GetModemStatus(lineNumber))
        stb.AppendLine("     </tr>")
        stb.AppendLine("     <tr>")
        stb.AppendLine("      <td style=""width: 20%""><strong>Ring Count:</strong></td>")
        stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics(lineNumber, "RingCount")).ToString())
        stb.AppendLine("     </tr>")
        stb.AppendLine("     <tr>")
        stb.AppendLine("      <td style=""width: 20%""><strong>Call Count:</strong></td>")
        stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics(lineNumber, "CallCount")).ToString())
        stb.AppendLine("     </tr>")
        stb.AppendLine("     <tr>")
        stb.AppendLine("      <td style=""width: 20%""><strong>Drop&nbsp;Count:</strong></td>")
        stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics(lineNumber, "DropCount")).ToString())
        stb.AppendLine("     </tr>")
        stb.AppendLine("    </table>")
        stb.AppendLine("   </fieldset>")
        stb.AppendLine("  </td>")
        stb.AppendLine(" </tr>")
      Next

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Modem Configuration</td>")
      stb.AppendLine(" </tr>")

      '
      ' Configure Line Number
      '
      Dim selModemInterfaces As New clsJQuery.jqDropList("selModemInterfaces", PageName, True)
      selModemInterfaces.autoPostBack = True
      For modemNumber As Integer = 1 To 2 Step 1
        selModemInterfaces.AddItem(modemNumber.ToString, modemNumber.ToString, (gModemInterfaces.ToString = modemNumber.ToString))
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Connected Modems</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selModemInterfaces.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      For lineNumber As Integer = 1 To gModemInterfaces
        '
        ' Configure Interface
        '
        Dim [interface] As String = String.Format("Interface{0}", lineNumber)
        '
        ' Modem Configuration Line #1
        '
        stb.AppendLine(" <tr>")
        stb.AppendFormat("  <td class='tableheader' colspan='2'>Modem Configuration Line {0}</td>", lineNumber)
        stb.AppendLine(" </tr>")

        '
        ' Modem Configuration (Serial Com Port)
        '
        Dim selInterfaceSerialId As String = String.Format("selInterfaceSerial{0}", lineNumber)
        Dim selInterfaceSerial As New clsJQuery.jqDropList(selInterfaceSerialId, Me.PageName, False)
        selInterfaceSerial.id = selInterfaceSerialId
        selInterfaceSerial.toolTip = "Specify the serial port number used to connect your caller ID modem."

        Dim txtInterfaceSerial As String = GetSetting([interface], "Serial", "0")
        selInterfaceSerial.AddItem("Disabled", "Disabled", txtInterfaceSerial = "0")
        Dim PortNames As Specialized.StringCollection = hspi_plugin.GetSerialPortNames()
        For Each strPortName As String In PortNames
          selInterfaceSerial.AddItem(strPortName, strPortName, txtInterfaceSerial = strPortName)
        Next

        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class='tablecell'>Serial Com Port</td>")
        stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selInterfaceSerial.Build, vbCrLf)
        stb.AppendLine(" </tr>")

        '
        ' Modem Configuration (Modem Initialization)
        '
        Dim selButton1Id As String = String.Format("btnSaveModemInit{0}", lineNumber)
        Dim jqButton1 As New clsJQuery.jqButton(selButton1Id, "Save", Me.PageName, True)
        stb.AppendLine()

        Dim chkModemInitId As String = String.Format("chkModemInit{0}", lineNumber)
        Dim chkModemInit As New clsJQuery.jqCheckBox(chkModemInitId, "&nbsp;Reset To Default", Me.PageName, True, False)
        chkModemInit.checked = False

        Dim txtModemInit As String = GetSetting([interface], "ModemInit", MODEM_INIT)
        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Modem&nbsp;Initialization</td>")
        stb.AppendFormat("  <td class='tablecell'><textarea rows='4' cols='50' name='txtModemInit{0}'>{1}</textarea>{2}{3}</td>{4}", lineNumber,
                                                                                                                                     txtModemInit.Trim.Replace("~", vbCrLf),
                                                                                                                                     jqButton1.Build(),
                                                                                                                                     chkModemInit.Build,
                                                                                                                                     vbCrLf)
        stb.AppendLine(" </tr>")

        '
        ' Modem Configuration (Drop Caller)
        '
        Dim selButton2Id As String = String.Format("btnSaveDropCaller{0}", lineNumber)
        Dim jqButton2 As New clsJQuery.jqButton(selButton2Id, "Save", Me.PageName, True)

        Dim chkDropCallerId As String = String.Format("chkDropCaller{0}", lineNumber)
        Dim chkDropCaller As New clsJQuery.jqCheckBox(chkDropCallerId, "&nbsp;Reset To Default", Me.PageName, True, False)
        chkDropCaller.checked = False

        Dim txtDropCaller As String = GetSetting([interface], "DropCaller", DROP_CALLER)
        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Drop&nbsp;Caller</td>")
        stb.AppendFormat("  <td class='tablecell'><textarea rows='4' cols='50' name='txtDropCaller{0}'>{1}</textarea>{2}{3}</td>{4}", lineNumber,
                                                                                                                                      txtDropCaller.Trim.Replace("~", vbCrLf),
                                                                                                                                      jqButton2.Build(),
                                                                                                                                      chkDropCaller.Build,
                                                                                                                                      vbCrLf)
        stb.AppendLine(" </tr>")

      Next

      '
      ' New Caller Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>New Caller Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' New Caller Options (Use Number Block Mask)
      '
      Dim selNmbrBlockMask As New clsJQuery.jqDropList("selNmbrBlockMask", Me.PageName, False)
      selNmbrBlockMask.id = "selNmbrBlockMask"
      selNmbrBlockMask.toolTip = "The Number of a new caller will be tested to see if the block attribute should be added."

      Dim txtNmbrBlockMask As String = CBool(GetSetting("CallerAttr", "UseNmbrBlockMask", "True")).ToString
      selNmbrBlockMask.AddItem("No", "0", txtNmbrBlockMask = "False")
      selNmbrBlockMask.AddItem("Yes", "1", txtNmbrBlockMask = "True")

      Dim txtNmbrBlockMaskDefault As String = GetSetting("CallerAttr", "NmbrBlockMask", "^8(00|88|77|66)")
      Dim tbNmbrBlockMask As New clsJQuery.jqTextBox("tbNmbrBlockMask", "text", txtNmbrBlockMaskDefault, PageName, 60, False)
      tbNmbrBlockMask.id = "tbNmbrBlockMask"
      tbNmbrBlockMask.promptText = "Regular expression used to determine if the block attribute should be added to a new caller."
      tbNmbrBlockMask.toolTip = tbNmbrBlockMask.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Use Number Block Mask</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}{1}</td>{2}", selNmbrBlockMask.Build, tbNmbrBlockMask.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' New Caller Options (Use Number Block Mask)
      '
      Dim selNameBlockMask As New clsJQuery.jqDropList("selNameBlockMask", Me.PageName, False)
      selNameBlockMask.id = "selNameBlockMask"
      selNameBlockMask.toolTip = "The Number of a new caller will be tested to see if the block attribute should be added."

      Dim txtNameBlockMask As String = CBool(GetSetting("CallerAttr", "UseNameBlockMask", "True")).ToString
      selNameBlockMask.AddItem("No", "0", txtNameBlockMask = "False")
      selNameBlockMask.AddItem("Yes", "1", txtNameBlockMask = "True")

      Dim txtNameBlockMaskDefault As String = GetSetting("CallerAttr", "NameBlockMask", "TOLL FREE CALL")
      Dim tbNameBlockMask As New clsJQuery.jqTextBox("tbNameBlockMask", "text", txtNameBlockMaskDefault, PageName, 60, False)
      tbNameBlockMask.id = "tbNameBlockMask"
      tbNameBlockMask.promptText = "Regular expression used to determine if the block attribute should be added to a new caller."
      tbNameBlockMask.toolTip = tbNameBlockMask.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Use Name Block Mask</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}{1}</td>{2}", selNameBlockMask.Build, tbNameBlockMask.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' New Caller Options (Add Announce Attribute)
      '
      Dim selAnnounceAttr As New clsJQuery.jqDropList("selAnnounceAttr", Me.PageName, False)
      selAnnounceAttr.id = "selAnnounceAttr"
      selAnnounceAttr.toolTip = "Specifies if a new caller should automatically have the announce attribute added."

      Dim txtAnnounceAttr As String = CBool(GetSetting("CallerAttr", "Announce", "True")).ToString
      selAnnounceAttr.AddItem("No", "0", txtAnnounceAttr = "False")
      selAnnounceAttr.AddItem("Yes", "1", txtAnnounceAttr = "True")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Add Announce Attribute</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selAnnounceAttr.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Caller Number Queries
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Caller Number Queries</td>")
      stb.AppendLine(" </tr>")

      '
      ' Caller Number Queries (Caller Number Query URL)
      '
      Dim txtCIDQueryDefault As String = GetSetting("NmbrQuery", "URL", CID_QUERY_URL)
      Dim tbCIDQuery As New clsJQuery.jqTextBox("tbCIDQuery", "text", txtCIDQueryDefault, PageName, 80, False)
      tbCIDQuery.id = "tbCIDQuery"
      tbCIDQuery.promptText = "Enter the URL that will be used when looking up a caller.  Use $nmbr for the number variable."
      tbCIDQuery.toolTip = tbCIDQuery.promptText

      Dim chkCIDQuery As New clsJQuery.jqCheckBox("chkCIDQuery", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkCIDQuery.checked = False

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Caller Number Query URL</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}{1}</td>{2}", tbCIDQuery.Build, chkCIDQuery.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Caller Number Formatting
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Caller Number Formatting</td>")
      stb.AppendLine(" </tr>")

      '
      ' Caller Number Formatting (9 Digit Format)
      '
      Dim txtNumberFormat09 As String = GetSetting("NmbrFormat", "09", NMBR_FORMAT_09)
      Dim tbNumberFormat09 As New clsJQuery.jqTextBox("tbNumberFormat09", "text", txtNumberFormat09, PageName, 80, False)
      tbNumberFormat09.id = "tbNumberFormat09"
      tbNumberFormat09.promptText = "Specifies the format of a 9 digit number."
      tbNumberFormat09.toolTip = tbNumberFormat09.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>9 Digit Format</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbNumberFormat09.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Caller Number Formatting (10 Digit Format)
      '
      Dim txtNumberFormat10 As String = GetSetting("NmbrFormat", "10", NMBR_FORMAT_10)
      Dim tbNumberFormat10 As New clsJQuery.jqTextBox("tbNumberFormat10", "text", txtNumberFormat10, PageName, 80, False)
      tbNumberFormat10.id = "tbNumberFormat10"
      tbNumberFormat10.promptText = "Specifies the format of a 10 digit number."
      tbNumberFormat10.toolTip = tbNumberFormat10.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>10 Digit Format</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbNumberFormat10.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Caller Number Formatting (11 Digit Format)
      '
      Dim txtNumberFormat11 As String = GetSetting("NmbrFormat", "11", NMBR_FORMAT_11)
      Dim tbNumberFormat11 As New clsJQuery.jqTextBox("tbNumberFormat11", "text", txtNumberFormat11, PageName, 80, False)
      tbNumberFormat11.id = "tbNumberFormat11"
      tbNumberFormat11.promptText = "Specifies the format of a 11 digit number."
      tbNumberFormat11.toolTip = tbNumberFormat11.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>11 Digit Format</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbNumberFormat11.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Web Page Access (Authorized User Roles)
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Options (Logging Level)
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Caller Log Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabCallerLog(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmCallerLog", "frmCallerLog", "Post"))

      stb.AppendLine("<div id='accordion-caller-log'>")
      stb.AppendLine(" <span class='accordion-toggle'>Caller Log Search</span>")
      stb.AppendLine(" <div class='accordion-content default'>")
      stb.AppendLine("  <table border='1' width='100%' style='border-collapse: collapse'>")
      stb.AppendLine("   <tbody>")

      Dim beginDate As DateTime = DateAdd(DateInterval.Month, -1, Date.Now)
      Dim tbBeginDate As New clsJQuery.jqTextBox("txtBeginDate", "text", beginDate.ToLongDateString, PageName, 30, True)
      tbBeginDate.id = "txtBeginDate"
      tbBeginDate.toolTip = ""
      tbBeginDate.editable = True

      Dim selBeginTime As New clsJQuery.jqDropList("selBeginTime", Me.PageName, True)
      selBeginTime.id = "selBeginTime"
      selBeginTime.toolTip = ""
      For hour As Integer = 0 To 23
        For min As Integer = 0 To 45 Step 15
          Dim time As String = String.Format("{0}:{1}", hour.ToString.PadLeft(2, "0"), min.ToString.PadLeft(2, "0"))
          selBeginTime.AddItem(time, time, time = "00:00")
        Next
      Next

      Dim tbEndDate As New clsJQuery.jqTextBox("txtEndDate", "text", Date.Now.ToLongDateString, PageName, 30, True)
      tbEndDate.id = "txtEndDate"
      tbEndDate.toolTip = ""
      tbEndDate.editable = True

      Dim selEndTime As New clsJQuery.jqDropList("selEndTime", Me.PageName, False)
      selEndTime.id = "selEndTime"
      selEndTime.toolTip = ""
      For hour As Integer = 0 To 23
        For min As Integer = 0 To 45 Step 15
          Dim time As String = String.Format("{0}:{1}", hour.ToString.PadLeft(2, "0"), min.ToString.PadLeft(2, "0"))
          selEndTime.AddItem(time, time, time = "00:00")
        Next
        selEndTime.AddItem("23:59", "23:59", True)
      Next

      stb.AppendLine("    <tr>")
      stb.AppendLine("     <td>")
      stb.AppendLine("    <strong>From:</strong>&nbsp;" & tbBeginDate.Build & "&nbsp;at&nbsp;" & selBeginTime.Build & "&nbsp;")
      stb.AppendLine("    <strong>To:</strong>&nbsp;" & tbEndDate.Build & "&nbsp;at&nbsp;" & selEndTime.Build & "&nbsp;")
      '
      ' Build Sort Order
      '
      Dim selSortOrder As New clsJQuery.jqDropList("selSortOrder", Me.PageName, True)
      selSortOrder.id = "selSortOrder"
      selSortOrder.toolTip = ""

      Dim strSortOrder As String = GetSetting("Webpage", "SortOrder", "desc")
      selSortOrder.AddItem("Sort Ascending", "asc", "asc" = strSortOrder)
      selSortOrder.AddItem("Sort Descending", "desc", "desc" = strSortOrder)
      stb.AppendLine(selSortOrder.Build())

      Dim btnFilterLogs As New clsJQuery.jqButton("btnFilterLogs", "Filter", Me.PageName, True)
      stb.AppendLine(btnFilterLogs.Build())

      stb.AppendLine("     </td>")
      stb.AppendLine("    </tr>")
      stb.AppendLine("   </tbody>")
      stb.AppendLine(" </table>")

      stb.AppendLine(" </div>")
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      stb.AppendLine("<div class='clear'><br /></div>")
      stb.AppendLine("<div id='divCallerLogTable'>")
      stb.AppendLine(BuildCallerLogTable(strSortOrder, beginDate.ToLongDateString, "00:00", Date.Now.ToLongDateString, "23:59"))
      stb.AppendLine("</div>")

      If Rebuilding Then Me.divToUpdate.Add("divCallerLog", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabLogs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Caller Log Table
  ''' </summary>
  ''' <param name="sortOrder"></param>
  ''' <param name="beginDate"></param>
  ''' <param name="beginTime"></param>
  ''' <param name="endDate"></param>
  ''' <param name="endTime"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildCallerLogTable(ByVal sortOrder As String,
                               ByVal beginDate As String,
                               ByVal beginTime As String,
                               ByVal endDate As String,
                               ByVal endTime As String) As String

    Try

      Dim stb As New StringBuilder

      Dim pageSize As Integer = 5000
      Dim pageCur As Integer = 1
      Dim pageCount As Integer = 1
      Dim recordCount As Integer = 0

      Dim filterValue As String = "%"
      sortOrder = IIf(sortOrder.Length = 0, "asc", sortOrder)

      '
      ' Calculate the beginDate
      '            
      beginTime = IIf(beginTime = "", "00:00:01", beginTime & ":00")
      If IsDate(beginDate) = False Then
        beginDate = Date.Now.ToLongDateString
      Else
        beginDate = Date.Parse(beginDate).ToLongDateString
      End If
      beginDate = String.Format("{0} {1}", beginDate, beginTime)

      '
      ' Calculate the endDate
      '  
      endTime = IIf(endTime = "", "23:59:59", endTime & ":59")
      If IsDate(endDate) = False Then
        endDate = Date.Now.ToLongDateString
      Else
        endDate = Date.Parse(endDate).ToLongDateString
      End If
      endDate = String.Format("{0} {1}", endDate, endTime)

      Dim dbFields As String = "datetime(ts, 'unixepoch', 'localtime') as ts, nmbr, name"
      Dim strSQL As String = BuildCallerLogSQL("tblCallerLog", dbFields, beginDate, endDate, filterValue, sortOrder)

      stb.AppendLine("<table id='table_caller_log' class='display compact' style='width:100%'>")
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Caller Date</th>")
      stb.AppendLine("   <th>Caller Number</th>")
      stb.AppendLine("   <th>Caller Name</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")
      stb.AppendLine("<tbody>")

      Using MyDataTable As DataTable = ExecuteSQL(strSQL, recordCount, pageSize, pageCount, pageCur)

        Dim iRowIndex As Integer = 1
        If MyDataTable.Columns.Contains("ts") Then

          For Each row As DataRow In MyDataTable.Rows
            Dim ts As String = row("ts")
            Dim nmbr As String = row("nmbr")
            Dim name As String = row("name")
            Dim caller_alias As String = hspi_database.GetCallerAlias(nmbr, name)

            stb.AppendLine("  <tr>")
            stb.AppendFormat("   <td>{0}</td>{1}", ts, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", FormatNmbr(nmbr, True), vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", caller_alias, vbCrLf)
            stb.AppendLine("  </tr>")

          Next
        End If

      End Using

      stb.AppendLine("</tbody>")
      stb.AppendLine(" <tfoot>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Caller Date</th>")
      stb.AppendLine("   <th>Caller Number</th>")
      stb.AppendLine("   <th>Caller Name</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </tfoot>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildCallerLogTable")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Log Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabCallerDetails(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmCallerDetail", "frmCallerDetail", "Post"))

      stb.AppendLine("<div id='accordion-caller-detail'>")
      stb.AppendLine(" <span class='accordion-toggle'>Caller Log Search</span>")
      stb.AppendLine(" <div class='accordion-content default'>")
      stb.AppendLine("  <table border='1' width='100%' style='border-collapse: collapse'>")
      stb.AppendLine("   <tbody>")
      stb.AppendLine("    <tr>")
      stb.AppendLine("     <td>")

      '
      ' Build Filter Field
      '
      Dim selFilterField As New clsJQuery.jqDropList("selFilterField", Me.PageName, True)
      selFilterField.id = "selFilterField"
      selFilterField.toolTip = ""

      Dim strFilterField As String = GetSetting("WebPage", "FilterField", "log_data")
      selFilterField.AddItem("Caller Name", "name", "name" = strFilterField)
      selFilterField.AddItem("Caller Number", "nmbr", "nmbr" = strFilterField)
      stb.AppendLine(selFilterField.Build())

      '
      ' Build Filter Compare
      '
      Dim selFilterCompare As New clsJQuery.jqDropList("selFilterCompare", Me.PageName, True)
      selFilterCompare.id = "selFilterCompare"
      selFilterCompare.toolTip = ""

      Dim strFilterCompare As String = GetSetting("WebPage", "FilterCompare", "is")
      selFilterCompare.AddItem("Is", "is", "is" = strFilterCompare)
      selFilterCompare.AddItem("Starts With", "starts with", "starts with" = strFilterCompare)
      selFilterCompare.AddItem("Ends With", "ends with", "ends with" = strFilterCompare)
      selFilterCompare.AddItem("Contains", "contains", "contains" = strFilterCompare)
      stb.AppendLine(selFilterCompare.Build())

      '
      ' Build Filter Value
      '
      Dim tbFilterValue As New clsJQuery.jqTextBox("txtFilterValue", "text", "", PageName, 30, True)
      tbFilterValue.id = "txtFilterValue"
      tbFilterValue.toolTip = ""
      tbFilterValue.editable = True
      stb.AppendLine(tbFilterValue.Build())

      '
      ' Build Sort Order
      '
      Dim selSortOrder As New clsJQuery.jqDropList("selSortOrder", Me.PageName, True)
      selSortOrder.id = "selSortOrder"
      selSortOrder.toolTip = ""

      Dim strSortOrder As String = GetSetting("Webpage", "SortOrder", "desc")
      selSortOrder.AddItem("Sort Ascending", "asc", "asc" = strSortOrder)
      selSortOrder.AddItem("Sort Descending", "desc", "desc" = strSortOrder)
      stb.AppendLine(selSortOrder.Build())

      '
      ' Build Sort Field
      '
      Dim selSortField As New clsJQuery.jqDropList("selSortField", Me.PageName, True)
      selSortField.id = "selSortField"
      selSortField.toolTip = ""

      Dim strSortField As String = GetSetting("WebPage", "SortField", "name")
      selSortField.AddItem("Caller Name", "name", "name" = strSortField)
      selSortField.AddItem("Caller Number", "nmbr", "nmbr" = strSortField)
      selSortField.AddItem("Caller Count", "call_count", "call_count" = strSortField)
      selSortField.AddItem("Last Call Date", "last_ts", "last_ts" = strSortField)
      stb.AppendLine(selSortField.Build())

      Dim btnFilterDetails As New clsJQuery.jqButton("btnFilterDetails", "Filter", Me.PageName, True)
      stb.AppendLine(btnFilterDetails.Build())

      stb.AppendLine("     </td>")
      stb.AppendLine("    </tr>")
      stb.AppendLine("   </tbody>")
      stb.AppendLine(" </table>")

      stb.AppendLine(" </div>")
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      stb.AppendLine("<div class='clear'><br /></div>")
      stb.AppendLine("<div id='divCallerDetailsTable'>")
      stb.AppendLine(BuildCallerDetailsTable(strFilterField, strFilterCompare, "", strSortOrder, strSortField))
      stb.AppendLine("</div>")

      If Rebuilding Then Me.divToUpdate.Add("divLog", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabLogs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build Caller Details Table
  ''' </summary>
  ''' <param name="filterField"></param>
  ''' <param name="filterCompare"></param>
  ''' <param name="filterValue"></param>
  ''' <param name="sortOrder"></param>
  ''' <param name="sortField"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildCallerDetailsTable(ByVal filterField As String,
                                   ByVal filterCompare As String,
                                   ByVal filterValue As String,
                                   ByVal sortOrder As String,
                                   ByVal sortField As String) As String

    Try

      Dim stb As New StringBuilder

      Dim pageSize As Integer = 5000
      Dim pageCur As Integer = 1
      Dim pageCount As Integer = 1
      Dim recordCount As Integer = 0

      Dim strSQL As String = "SELECT id, nmbr, name, attr, notes, datetime(last_ts, 'unixepoch', 'localtime') as last_ts, call_count FROM tblCallerDetails"

      If filterValue.Length > 0 Then

        If filterField = "nmbr" Then
          filterValue = Regex.Replace(filterValue, "[^\d]", "")
        End If

        Select Case filterCompare
          Case "is"
            filterCompare = "="
          Case "starts with"
            filterCompare = "LIKE"
            filterValue = String.Format("{0}%", filterValue)
          Case "ends with"
            filterCompare = "LIKE"
            filterValue = String.Format("%{0}", filterValue)
          Case "contains"
            filterCompare = "LIKE"
            filterValue = String.Format("%{0}%", filterValue)
        End Select

        Dim strWhere As String = String.Format(" WHERE {0} {1} '{2}'", filterField, filterCompare, filterValue)
        strSQL = String.Concat(strSQL, strWhere)

      End If

      '
      ' Add in the Order By and Sort Order
      '
      sortOrder = IIf(sortOrder.Length = 0, "asc", sortOrder)
      strSQL = String.Format("{0} ORDER BY {1} {2}", strSQL, sortField, sortOrder)

      stb.AppendLine("<table id='table_caller_details' class='display compact' style='width:100%'>")
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Caller Number</th>")
      stb.AppendLine("   <th>Caller Name</th>")
      stb.AppendLine("   <th>Caller Notes</th>")
      stb.AppendLine("   <th>Caller Attributes</th>")
      stb.AppendLine("   <th>Caller Count</th>")
      stb.AppendLine("   <th>Caller Date</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")
      stb.AppendLine("<tbody>")

      Using MyDataTable As DataTable = ExecuteSQL(strSQL, recordCount, pageSize, pageCount, pageCur)
        'id, nmbr, name, attr, notes, last_ts, call_count
        Dim iRowIndex As Integer = 1
        If MyDataTable.Columns.Contains("id") Then

          For Each row As DataRow In MyDataTable.Rows
            Dim id As Integer = row("id")
            Dim last_ts As String = row("last_ts")
            Dim nmbr As String = row("nmbr")
            Dim name As String = row("name")
            Dim attr As Integer = row("attr")
            Dim notes As String = row("notes") & ""
            Dim call_count As Integer = row("call_count")

            stb.AppendFormat("  <tr id='{0}'>{1}", id.ToString, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", FormatNmbr(nmbr, True), vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", name, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", notes, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", FormatCallerAttrs(attr.ToString), vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", call_count, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", last_ts, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
            stb.AppendLine("  </tr>")

          Next
        End If

      End Using

      stb.AppendLine("</tbody>")
      stb.AppendLine(" <tfoot>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Caller Number</th>")
      stb.AppendLine("   <th>Caller Name</th>")
      stb.AppendLine("   <th>Caller Notes</th>")
      stb.AppendLine("   <th>Caller Attributes</th>")
      stb.AppendLine("   <th>Caller Count</th>")
      stb.AppendLine("   <th>Caller Date</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </tfoot>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildCallerLogTable")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If

      Select Case postData("action")
        Case "edit"
          Dim id As String = postData("id")
          Dim name As String = postData("data[name]").Trim
          Dim notes As String = postData("data[notes]").Trim
          Dim attr As String = postData("data[attr]").Trim
          Dim iCallerAttrs As Integer = 0

          If attr.Length > 0 Then
            Dim chkCallerAttr As String() = attr.Split(",")
            For Each callerAttr As String In chkCallerAttr
              Dim iAttr As CallerAttrs = System.Enum.Parse(GetType(CallerAttrs), callerAttr)
              iCallerAttrs = iCallerAttrs Or iAttr
            Next
          End If

          Dim strSQL1 As String = String.Format("UPDATE tblCallerDetails SET name='{0}', notes='{1}', attr={2} WHERE id={3}", name.ToUpper, notes, iCallerAttrs, id)
          ExecuteSQL(strSQL1, 0, 0, 0, 0)

          Dim nmbr As String = hspi_database.GetCallerNumberById(id)
          If nmbr.Length > 0 Then
            Dim strSQL2 As String = String.Format("UPDATE tblCallerLog SET name='{0}' WHERE nmbr={1}", name.ToUpper, nmbr)
            ExecuteSQL(strSQL2, 0, 0, 0, 0)
          End If

          Dim bSuccess As Boolean = True

          If bSuccess = False Then
            Dim sb As New StringBuilder

            sb.AppendLine("{")
            sb.AppendFormat(" ""error"": {0}{1}", "Unable to update caller due to an error.", vbCrLf)
            sb.AppendLine("}")

            Return sb.ToString
          Else
            'BuildTabCallerDetails(True)
            'Me.pageCommands.Add("executefunction", "reDrawCallerDetails()")
            Return "{ }"
          End If

        Case "remove"
          Dim id As String = Val(postData("id[]"))

          Dim strSQL As String = String.Format("DELETE FROM tblCallerDetails WHERE id={0}", id)
          ExecuteSQL(strSQL, 0, 0, 0, 0)
          Dim bSuccess As Boolean = True

          If bSuccess = False Then
            Dim sb As New StringBuilder

            sb.AppendLine("{")
            sb.AppendFormat(" ""error"": {0}{1}", "Unable to delete caller due to an error.", vbCrLf)
            sb.AppendLine("}")

            Return sb.ToString
          Else
            'BuildTabCallerDetails(True)
            'Me.pageCommands.Add("executefunction", "reDrawCallerDetails()")
            Return "{ }"
          End If

      End Select

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabCallerLog"

        Case "btnFilterLogs"
          Dim sortOrder As String = postData("selSortOrder")
          Dim beginDate As String = postData("txtBeginDate")
          Dim beginTime As String = postData("selBeginTime")
          Dim endDate As String = postData("txtEndDate")
          Dim endTime As String = postData("selEndTime")

          Me.divToUpdate.Add("divCallerLogTable", BuildCallerLogTable(sortOrder, beginDate, beginTime, endDate, endTime))
          Me.pageCommands.Add("executefunction", "reDrawCallerLog()")

        Case "btnFilterDetails"
          Dim filterField As String = postData("selFilterField")
          Dim filterCompare As String = postData("selFilterCompare")
          Dim filterValue As String = postData("txtFilterValue")
          Dim sortOrder As String = postData("selSortOrder")
          Dim sortField As String = postData("selSortField")

          Me.divToUpdate.Add("divCallerDetailsTable", BuildCallerDetailsTable(filterField, filterCompare, filterValue, sortOrder, sortField))
          Me.pageCommands.Add("executefunction", "reDrawCallerDetails()")

        Case "selModemInterfaces"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Interface", "Count", strValue)

          gModemInterfaces = Integer.Parse(strValue)

          PostMessage("The Connected Modem Count has been updated.  A restart of HomeSeer may be required.")

        Case "selInterfaceSerial1"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Interface1", "Serial", strValue)

          PostMessage("The Line #1 Modem Serial Port has been updated.  A restart of HomeSeer may be required.")

        Case "selInterfaceSerial2"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Interface2", "Serial", strValue)

          PostMessage("The Line #2 Modem Serial Port has been updated.  A restart of HomeSeer may be required.")

        Case "chkModemInit1"
          SaveSetting("Interface1", "ModemInit", MODEM_INIT)
          BuildTabOptions(True)

          PostMessage("The Line #1 Modem Initialization option has been reset.")

        Case "chkModemInit2"
          SaveSetting("Interface2", "ModemInit", MODEM_INIT)
          BuildTabOptions(True)

          PostMessage("The Line #2 Modem Initialization option has been reset.")

        Case "chkDropCaller1"
          SaveSetting("Interface1", "DropCaller", DROP_CALLER)
          BuildTabOptions(True)

          PostMessage("The Line #1 Drop Caller option has been reset.")

        Case "chkDropCaller2"
          SaveSetting("Interface2", "DropCaller", DROP_CALLER)
          BuildTabOptions(True)

          PostMessage("The Line #2 Drop Caller option has been reset.")

        Case "btnSaveModemInit1"
          Dim strValue As String = postData("txtModemInit1").Trim.Replace(vbCrLf, "~")
          SaveSetting("Interface1", "ModemInit", strValue)

          PostMessage("The Line #1 Modem Initialization option has been updated.")

        Case "btnSaveModemInit2"
          Dim strValue As String = postData("txtModemInit1").Trim.Replace(vbCrLf, "~")
          SaveSetting("Interface2", "ModemInit", strValue)

          PostMessage("The Line #2 Modem Initialization option has been updated.")

        Case "btnSaveDropCaller1"
          Dim strValue As String = postData("txtDropCaller1").Trim.Replace(vbCrLf, "~")
          SaveSetting("Interface1", "DropCaller", strValue)

          PostMessage("The Line #1 Drop Caller option has been updated.")

        Case "btnSaveDropCaller2"
          Dim strValue As String = postData("txtDropCaller2").Trim.Replace(vbCrLf, "~")
          SaveSetting("Interface2", "DropCaller", strValue)

          PostMessage("The Line #2 Drop Caller option has been updated.")

        Case "selNmbrBlockMask"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("CallerAttr", "UseNmbrBlockMask", strValue)

          PostMessage("The Number Block Mask option has been updated.")

        Case "tbNmbrBlockMask"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("CallerAttr", "NmbrBlockMask", strValue)

          PostMessage("The Number Block Regular Expression option has been updated.")

        Case "selNameBlockMask"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("CallerAttr", "UseNameBlockMask", strValue)

          PostMessage("The Name Block Mask option has been updated.")

        Case "tbNameBlockMask"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("CallerAttr", "NameBlockMask", strValue)

          PostMessage("The Name Block Mask option has been updated.")

        Case "selAnnounceAttr"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("CallerAttr", "Announce", strValue)

          PostMessage("The Announce option has been updated.")

        Case "tbCIDQuery"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("NmbrQuery", "URL", strValue)

          PostMessage("The Caller Number Query URL option has been updated.")

        Case "chkCIDQuery"
          SaveSetting("NmbrQuery", "URL", CID_QUERY_URL)
          BuildTabOptions(True)

          PostMessage("The Caller Number Query URL option has been reset.")

        Case "tbNumberFormat09"
          Dim strValue As String = postData(postData("id"))
          If ValidateNmbrFormat(postData("id"), strValue) = True Then
            SaveSetting("NmbrFormat", "09", strValue)
            PostMessage("The 9 digit number option has been updated.")
          Else
            PostMessage("The 9 digit number value is not valid.")
          End If

        Case "tbNumberFormat10"
          Dim strValue As String = postData(postData("id"))
          If ValidateNmbrFormat(postData("id"), strValue) = True Then
            SaveSetting("NmbrFormat", "10", strValue)
            PostMessage("The 10 digit number option has been updated.")
          Else
            PostMessage("The 10 digit number value is not valid.")
          End If

        Case "tbNumberFormat11"
          Dim strValue As String = postData(postData("id"))
          If ValidateNmbrFormat(postData("id"), strValue) = True Then
            SaveSetting("NmbrFormat", "11", strValue)
            PostMessage("The 11 digit number option has been updated.")
          Else
            PostMessage("The 11 digit number value is not valid.")
          End If

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class
