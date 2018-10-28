Imports System.Text.RegularExpressions
Imports System.IO

Module common
    '*** VERSION 2.5
    '*** 2014-05-28 rewritten as ASP.NET 2
    '*** 2014-06-15 modified to simplify CheckSession and to make it work with a general local permission table
    '*** 2014-06-27 bug fix in SendMail addresses need to be added individually.  semicolon separated addresses not valid.
    '*** 2014-09-03 CP810 cookie timeout added in CheckSession
    '*** 2014-11-04 StreamFile added
    '*** 2014-11-04 exportRFC4180table added
    '*** 2014-11-06 fixed bug in CheckSession, it needs to check all session vars are present
    '*** 2014-11-17 updated setDropDown to work like doBindDataRow, changed connection string to sConn
    '*** 2014-11-20 added sortACDC overload to handle datagrids
    '*** 2014-11-25 added setGVcolVis
    '*** 2014-11-17 Note: doBindDataRow and doUnbindDataRow work with a lookHere object, this could be a page, a DataGridRow or a GridView row
    '*** 2015-04-29 added checkNTLMUser
    '*** 2015-05-04 added setColVis overload to handle column visibility on datagrids and gridview.  setDGcolVis/ setGVcolVis are deprecated
    '*** 2015-06-09 added blPathInfo to checkNTLMuser as a means of cleaning up URLs with appended target info
    '*** 2015-06-19 checkNTLMUser supports tblName as the permissions table, must be Read/Write
    '*** 2015-07-03 bug fix.  Users > 120 days inactive OR with NULL LastLoginUTC will return false
    '*** 2015-09-17 checkNTLMuser will return false if AccountLocked=true
    '*** 2015-09-21 streamfile updated to write to log and not reveal path info on error
    '*** 2016-09-09 bugfix checkNTLMuser where lastLoginUTC is null 
    '*** 2018-03-21 added HttpUtility.HTMLEncode(string) in the doBindDataRow routine to protect against cross site script attack XSS
    '*** [have i doubled up?  i use doBindDataRow on form fields such as fld_ so there is no need to use fnAsHTML on these too, however i do
    '*** need to use fnAsHTML on a datagrid as this is bound from a dataset, not thro fnASHTML]
    '*** 2018-03-26 added urldecode to pathinfo in checkNTLMuser
    '*** 2018-04-27 added checkLogonUser
    '*** 2018-10-18 added dtColumnOrder

    '*** Important:  you must be consistent with the web.config name for the connection string, as its called from this routine as well as the main pages
    Dim sConn As String = ConfigurationManager.ConnectionStrings("sConn").ConnectionString


#Region "...AUTHENTICATION AND LOGGING"
    Public Function checkLogonUser(ByVal myPage As Page, Optional ByVal blRefresh As Boolean = False, Optional ByVal blPathInfo As Boolean = False, Optional ByVal tblName As String = "tblUserPermission") As Boolean
        '*** 2018-04-27 supports use of IIS7 useRemoteWorkerProcess as the actual user identity at an ACL level.  We then check the AD authenticated user, who is
        '*** the LOGON_USER and this is the user we check is credentialled in our permitted user list.  This avoids need to ACL every user against the web directory
        '*** but does leverage AD authentication. It also allows self registration.
        '*** LOGON_USER is the AD authenticated user.  AUTH_USER and REMOTE_USER will be the app pool
        With myPage
            If blPathInfo Then
                If .Request.PathInfo.Length > 0 Then
                    'capture the pathInfo and add it to the Session object.  This is the easiest way to make it survive the 
                    'the redirect which we use to clean up the URL
                    '*** 2018-03-26 added URL decode to protect against XSS
                    .Session.Add("PATHINFO", HttpUtility.UrlDecode(.Request.PathInfo))
                    .Response.Redirect(.Request.ServerVariables("URL"))
                End If
            End If


            '1/ If session is still valid, exit true unless we are forcing a session vars refresh
            If (Not .Session("AUTHUSER") Is Nothing) And blRefresh = False Then
                '*** 2015-07-03 bug fix.  Users > 120 days inactive OR with no LastLoginUTC will return false
                If Not IsDate(.Session("LASTLOGINUTC")) Then Return False

                '*** users >120 days since last login will return false
                If (DateDiff(DateInterval.Day, .Session("LASTLOGINUTC"), Date.UtcNow)) > 120 Then Return False

                '*** 2015-09-17  If AccountLocked return false
                If CBool(.Session("ACCOUNTLOCKED")) Then Return False

                Return True
            End If

            '2/ if NTLM/AD fails to authenticate, terminate the app.  Don't want to use the test user in this scenario because its a security risk
            'any NTLM failure will lead to all users becoming the test user.
            If .User.Identity.IsAuthenticated = False Then
                '*** terminate the page
                .Response.Write("<b=""red"">FATAL ERROR:  Active Directory user cannot be identified.</b>  Contact system administrator and report this error.")
                .Response.End()
                Return False
            End If


            '2a/ possible alt user
            '**** does User.Identity.Name match into altUser?
            If ConfigurationManager.AppSettings("altUser") Is Nothing Then
                '*** no altUser so go with LOGON_USER
                .Session("AUTHUSER") = .Request("LOGON_USER")
                '*** look for LOGON_USER in altUser, use Instr because regex will be confused by the \

                '*** new block will match full string prior to the colon
            ElseIf String.Equals(.Request("LOGON_USER"), Regex.Replace(ConfigurationManager.AppSettings("altUser").ToString, "([^\x5c]+\x5c\w+):(\w+)$", "$1"), _
                                 StringComparison.CurrentCultureIgnoreCase) = True Then

                '*** matches, so substitute the second group vID with the part after the colon. \x5C is a \ char
                .Session("AUTHUSER") = Regex.Replace(ConfigurationManager.AppSettings("altUser").ToString, "([^\x5c]+)\x5c(\w+):(\w+)$", "$1\$3")
            Else
                '*** does not match, use the LOGON_USER
                .Session("AUTHUSER") = .Request("LOGON_USER")
            End If

            '3/  refresh the system vars now that we have set a valid Session("AUTHUSER")
            Dim oConn As New OleDb.OleDbConnection(sConn)

            Try
                Dim objCmd As New OleDb.OleDbCommand("SELECT * FROM " & tblName & " WHERE AUTHUSER=@p1", oConn)
                objCmd.Parameters.Add("@p1", OleDb.OleDbType.VarChar).Value = .Session("AUTHUSER")
                oConn.Open()
                Dim objRead As OleDb.OleDbDataReader = objCmd.ExecuteReader(CommandBehavior.CloseConnection)

                If objRead.Read Then
                    Dim n As Integer
                    For n = 0 To objRead.FieldCount - 1
                        '*** note that all session vars are UPPER CASE to avoid case problems, even though in asp.net
                        '*** session keys are case insensitive
                        If objRead.Item(n) Is DBNull.Value Then
                            '*** map null to string.empty to help with regex tests later
                            .Session(objRead.GetName(n).ToUpper) = String.Empty
                        Else
                            .Session(objRead.GetName(n).ToUpper) = objRead.Item(n)
                        End If
                    Next
                    objRead.Close()
                End If

                '*** 2015-09-17  If AccountLocked return false
                If CBool(.Session("ACCOUNTLOCKED")) Then Return False


                '*** 2015-05-11 For CPS108 compliance, we should deny users with >120 day access. To do this we'd simply test last login UTC
                '*** and return false at this step, not update the lastlogin value.
                '  .Trace.Warn("chckNTLM " & DateDiff(DateInterval.Day, .Session("LASTLOGINUTC"), Date.UtcNow))

                '*** 2015-07-03 bug fix.  Users > 120 days inactive OR with no LastLoginUTC will return false
                If Not IsDate(.Session("LASTLOGINUTC")) Then Return False

                If (DateDiff(DateInterval.Day, .Session("LASTLOGINUTC"), Date.UtcNow)) > 120 Then Return False
                '*** these session vars are not changed.  Main program code must look at the LASTLOGINUTC as a possible reason for the reject.
                '*** to reset the 120 day lockout, the admin must re-update a locked user.

                '*** 2015-05-11 update the lastlogin value
                objCmd = New OleDb.OleDbCommand("Update " & tblName & " SET LastLoginUTC=@p1 WHERE AUTHUSER=@p2", oConn)

                objCmd.Parameters.Add("@p1", OleDb.OleDbType.Date).Value = DateTime.UtcNow
                objCmd.Parameters.Add("@p2", OleDb.OleDbType.VarChar).Value = .Session("AUTHUSER")
                oConn.Open()
                objCmd.ExecuteNonQuery()
                oConn.Close()

                Return True
            Catch ex As Exception
                .Trace.Warn(ex.ToString)
                Return False
            Finally
                oConn.Dispose()
            End Try

        End With
    End Function

    Public Function writeAudit(ByVal buffer As String, ByVal sNTID As String) As String
        '*** write to audit.txt in the files folder
        '*** needs reference imports system.IO
        'http://www.builderau.com.au/program/windows/soa/Reading-and-writing-text-files-with-VB-NET/0,339024644,320267367,00.htm

        Dim oWrite As StreamWriter
        Try
            oWrite = File.AppendText(System.Configuration.ConfigurationManager.AppSettings("strUserFiles") & "\audit.txt")
            buffer = String.Concat(Format(DateTime.UtcNow, "u"), vbTab, sNTID, vbTab, buffer)
            oWrite.WriteLine(buffer)
            'oWrite.WriteLine("{0,10:dd MMMM}{0,10:hh:mm tt}{1,25:C}", Now(), 13455.33)
            oWrite.Close()
            Return True
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
#End Region


#Region "...DATATABLE RELATED..."
    Sub doBindDataRow(ByRef myRow As DataRow, ByVal lookHere As Object, Optional ByVal sPrefix As String = "fld_")
        '*** this code takes a data row, and searches the lookHere object (page or datagriditem) for controls with an id of fld_field1 (this prefix can be overidden) etc
        '*** where field1 is a dataset field
        '*** additional attribute you can put on the control are;
        '*** optional DFS="{0:c}" on TextBox controls; this formats the presentation
        '*** 2009-12-14 attribute bind usage; bind="text|value [nobind] [legacy] [blank]" where;
        '*** text|value signifies how to bind to the dropdownlist options to the database
        '*** optional nobind signifies this sub is to ignore the control (useful if other code sets value and you still want to unbind)
        '*** optional legacy will add an extra dropdown option and select it to support legacy data
        '*** optional blank will add a blank entry to end of list
        '*** CheckBox controls are rendered to checked or not based on boolean value
        '*** IMPORTANT findbyValue is case sensitive
        '*** 2018-03-21 text and literals are HTMLencoded to protect against XSS attacks.

        Dim myColumn As DataColumn

        '*** find parent DataTable and its schema
        For Each myColumn In myRow.Table.Columns

            Try
                '*** pick up the data item, don't force type conversion yet
                Dim s As Object = myRow.Item(myColumn.ColumnName)
                '*** now lets populate the form. First find the control
                Dim myControl As Object = lookHere.FindControl(sPrefix & myColumn.ColumnName)
                '*** 2009-12-21 some controls, such as literals have no attributes and will cause an error if you try to access these
                '*** at this point, so wait until we are working with the DDlist
                If TypeOf (myControl) Is TextBox Then
                    Dim myTextBox As TextBox = myControl
                    '*** don't force type conversion yet, unless s is NULL
                    If s Is DBNull.Value Then s = String.Empty

                    '*** look for optional DFS attribute which will contain a formatting string
                    If myTextBox.Attributes("DFS") Is Nothing Then
                        '*** if no format string present, then convert to a string

                        myTextBox.Text = HttpUtility.HtmlEncode(s.ToString)
                    ElseIf myTextBox.Attributes("DFS") = "percent" Then
                        '*** special setting to convert a text field to a percent figure
                        If IsNumeric(s) Then
                            If s < 2 Then myTextBox.Text = FormatPercent(s, 2, TriState.True)
                        Else
                            myTextBox.Text = HttpUtility.HtmlEncode(s.ToString)
                        End If
                    Else
                        '*** the formatting will convert to string, but this formatting only works if the object is still
                        '*** intact e.g. a date, time, double etc, hence we did not force to a string earlier
                        myTextBox.Text = HttpUtility.HtmlEncode(String.Format(myTextBox.Attributes("DFS"), s))
                    End If

                    '*** for checkboxes, we only need worry about the boolean value
                ElseIf TypeOf (myControl) Is CheckBox Then
                    Dim myCheckbox As CheckBox = myControl
                    myCheckbox.Checked = CBool(s)

                ElseIf TypeOf (myControl) Is DropDownList Then
                    '*** For dropdowns, list items must be bound prior.  We are looking to select a text or value
                    '*** entry as found in our target datarow field
                    '*** switch options are bind="text|value [nobind] [legacy] [blank]"
                    '*** default is text if bind attrib not present
                    If Not Regex.IsMatch(myControl.attributes("bind") & String.Empty, "nobind", RegexOptions.IgnoreCase) Then
                        Dim myDropdown As DropDownList = myControl
                        Dim oItem As ListItem
                        Dim blLegacy As Boolean = Regex.IsMatch(myDropdown.Attributes("bind") & String.Empty, "legacy", RegexOptions.IgnoreCase)
                        Dim blBlank As Boolean = Regex.IsMatch(myDropdown.Attributes("bind") & String.Empty, "blank", RegexOptions.IgnoreCase)
                        If s Is DBNull.Value Then s = String.Empty '*** we can't bind dbNULL do convert to string.empty
                        '*** 2009-12-14 some changes to the bind parameter and legacy parameter
                        If Regex.IsMatch(myDropdown.Attributes("bind") & String.Empty, "value", RegexOptions.IgnoreCase) Then
                            '*** find by value, but first check wether we are supporting legacy values (i.e. those not bound in the list)
                            If myDropdown.Items.FindByValue(CType(s, String)) Is Nothing And blLegacy Then
                                oItem = New ListItem(HttpUtility.HtmlEncode(s.ToString), HttpUtility.HtmlEncode(s.ToString))
                                myDropdown.Items.Add(oItem) '*** add a legacy item
                                '*** now also add a blank if required (and hasn't just been added as a legacy item)
                            Else
                                oItem = myDropdown.Items.FindByValue(CType(s, String))
                            End If
                            myDropdown.SelectedIndex = myDropdown.Items.IndexOf(oItem)
                        Else
                            '*** default is to find by text
                            If myDropdown.Items.FindByText(CType(s, String)) Is Nothing And blLegacy Then
                                oItem = New ListItem(CType(s, String), CType(s, String))
                                myDropdown.Items.Add(oItem)
                            Else
                                oItem = myDropdown.Items.FindByText(CType(s, String))
                            End If
                            myDropdown.SelectedIndex = myDropdown.Items.IndexOf(oItem)
                        End If  '*** value test

                        '*** now add a blank if required, and select it if required
                        If blBlank Then
                            '*** add a blank if one does not already exist
                            If CBool(myDropdown.Items.FindByValue(String.Empty) Is Nothing) Then myDropdown.Items.Add(New ListItem(String.Empty, String.Empty))
                            '*** If we do not have a valid oItem from before, then select this blank value
                            If oItem Is Nothing Then
                                oItem = myDropdown.Items.FindByValue(String.Empty)
                                myDropdown.SelectedIndex = myDropdown.Items.IndexOf(oItem)
                            End If
                        End If '*** blank 

                    End If  '*** end nobind test

                ElseIf TypeOf (myControl) Is Literal Then
                    Dim myLiteral As Literal = myControl
                    If s Is DBNull.Value Then
                        myLiteral.Text = String.Empty
                    Else
                        myLiteral.Text = HttpUtility.HtmlEncode(s.ToString)
                    End If

                    '*** end of control types
                End If


            Catch
            End Try
        Next myColumn
    End Sub
    Sub doUnbindDataRow(ByRef myRow As DataRow, ByVal lookHere As Object, Optional ByVal sPrefix As String = "fld_")
        '*** based on the db-datarow column names, looks for fields fld_field1 etc on the page
        '*** lookHere could be the page object or a datagriditem
        '*** so whilst original page may have been populated with one query, its possible to bind
        '*** page control values back to a different query
        '*** NOTE do not rebind the controls until AFTER you have called this sub
        '*** NOTE: Dates are a problem.  If you force them to display as MM/DD/YYYY, you have a problem writing
        '*** them back to the db, as its locale might expect DD/MM/YYYY
        '*** also, if you don't want createDates being trunkated back in the db, need to make them display readonly or disabled
        '*** 2010-03-30 optional attribute bind="nounbind" meaning that we won't unbind the data back to the database the
        '*** reason for this feature is that sometimes we want to save a record, but have say a status dropdown written to db by other logic

        '*** 2015-05-05 IMPORTANT: do not use highlevel grid controls such as <asp:checkboxfield> in conjunction with <asp:templatefield> because the first
        '*** will unbind the control value and stop the templatefield from working.  templatefields also don't enumerate to useful IDs

        '*** 2018-03-21 to protect against ingesting data that is potentially a cross site script attack XSS, ensure that validaterequest=true (the page default)
        '*** there is no benefit to HTMLdecoding here because the page should trap potentially unsafe text strings
        '*** https://www.apexhost.com.au/knowledgebase.php?action=displayarticle&id=66

        Dim myColumn As DataColumn

        '*** REMEMBER that reserved names such as Currency cause problems for the Updatebuilder

        '*** find parent DataTable and its schema
        For Each myColumn In myRow.Table.Columns
            Try
                '*** pick up the data item, don't force type conversion yet
                Dim s As Object = myRow.Item(myColumn.ColumnName)
                '*** now lets populate the datarow from the form. First find the control
                Dim myControl As Object = lookHere.FindControl(sPrefix & myColumn.ColumnName)
                If myControl Is Nothing Then
                    '*** do nothing, control cannot be found
                ElseIf Regex.IsMatch(myControl.attributes("bind") & String.Empty, "nounbind", RegexOptions.IgnoreCase) Then
                    '*** do nothing, user has disabled unbinding of the control
                ElseIf Not myControl.Enabled Then
                    '*** do nothing - enabled is a property found on ALL controls
                    '*** else what type of control is it?
                ElseIf TypeOf (myControl) Is TextBox Then
                    Dim myTextBox As TextBox = myControl
                    '*** handle zero length strings - allowDBNull
                    If myTextBox.ReadOnly Then
                        '*** do nothing - this property is specific to textboxes
                    ElseIf myColumn.AllowDBNull And myTextBox.Text = "" Then
                        '*** write a null
                        myRow.Item(myColumn.ColumnName) = DBNull.Value
                    ElseIf myColumn.DataType.ToString = "System.Int32" Then
                        '*** convert numeric types first
                        myRow.Item(myColumn.ColumnName) = CLng(myTextBox.Text)
                    ElseIf myColumn.DataType.ToString = "System.DateTime" Then
                        myRow.Item(myColumn.ColumnName) = CDate(myTextBox.Text)
                    ElseIf myColumn.DataType.ToString = "System.Double" Then
                        '*** check to see if we need to handle a percentage
                        If Regex.IsMatch(myTextBox.Text, "^[-\d\.]+%$") Then
                            myRow.Item(myColumn.ColumnName) = CDbl(myTextBox.Text.ToString.Replace("%", String.Empty)) / 100
                        Else
                            myRow.Item(myColumn.ColumnName) = CDbl(myTextBox.Text)
                        End If

                    Else
                        '*** write any text, including zero len
                        '*** 2014-03-20 modified this to trim the spaces first
                        myRow.Item(myColumn.ColumnName) = myTextBox.Text.Trim
                    End If

                    '*** for checkboxes, we only need worry about the boolean value
                ElseIf TypeOf (myControl) Is CheckBox Then
                    Dim myCheckbox As CheckBox = myControl
                    myRow.Item(myColumn.ColumnName) = myCheckbox.Checked
                ElseIf TypeOf (myControl) Is DropDownList Then
                    '*** For dropdowns, we are looking for selecteditemValue only
                    Dim myDropdown As DropDownList = myControl
                    If myColumn.DataType.ToString() = "System.Boolean" Then
                        myRow.Item(myColumn.ColumnName) = CBool(myDropdown.SelectedValue)
                    ElseIf myColumn.AllowDBNull And myDropdown.SelectedValue = String.Empty Then
                        '*** write a null if we cannot have zero len strings
                        myRow.Item(myColumn.ColumnName) = DBNull.Value
                    Else
                        myRow.Item(myColumn.ColumnName) = myDropdown.SelectedValue
                    End If

                    '*** end control type tests
                End If

            Catch

            Finally
            End Try

        Next myColumn
    End Sub
    Sub dtColumnOrder(ByRef dt As DataTable, ByVal colNames() As String, Optional ByVal blDropRestOfCols As Boolean = False)
        '*** modifies and re-orders columns in a datatable. colNames is an array of existing names in desired order
        '*** Note in net 3 you need to declare the colNames() array before you call this routine
        '*** dim a() as string={"what","the"}   dtcolumnOrder(dt,a)
        '*** whereas in net 4, you can do directly in the function call dtcolumnOrder(dt,{"what","the"})
        Try

            Dim i As Int16 = 0
            For Each cn In colNames
                If dt.Columns.Contains(cn) Then
                    dt.Columns(cn).SetOrdinal(i)
                    i += 1
                End If
            Next

            '*** now remove ones we don't need.  pain, easy to check if dt.contains, but not easy to check if colNames() contains because this is not a function of array
            'https://www.dreamincode.net/forums/topic/102273-how-to-use-arrayexist-method/
            'https://forums.asp.net/t/2085592.aspx?check+value+exist+in+an+array

            If blDropRestOfCols = False Then Exit Sub
            '*** need to enumerate backwards when removing
            For i = dt.Columns.Count - 1 To 0 Step -1
                If Not colNames.Contains(dt.Columns(i).ColumnName) Then
                    dt.Columns.RemoveAt(i)
                End If
            Next

        Catch ex As Exception
            dt = Nothing
        End Try

    End Sub
    Sub exportRFC4180table(ByVal myPage As Page, ByVal dTbl As DataTable, Optional ByVal myFilename As String = "test.csv", Optional ByVal noSLYK As Boolean = False)
        '*** exports a datatable to XL. Originally this used vbTab chars to separate the variables, however if you do this then
        '*** XL throws an annoying error "content does not match the description" i.e. the thing is full of tabs but the filename is .xls or .csv
        '*** to work around it use commas as the separator, and use .csv as the filename.  Then it opens straight in XL with no fuss.
        '*** http://tools.ietf.org/html/rfc4180
        '*** WELL ALMOST.  you have to escape commas by enclosing in quotes.  I am not bothering to escape quotes.
        '*** GOTCHA. ID as the first column header will cause a problem as XL thinks the file is an SYLK file, remove this from your datatable

        '*** 2016-06-30 updated the tool to delete first col if noSLYK is true
        Try
            With myPage
                If noSLYK Then
                    dTbl.Columns.RemoveAt(0)
                End If

                .Trace.IsEnabled = False
                Dim attachment As String = "attachment; filename=" & myFilename
                .Response.ClearContent()
                .Response.AddHeader("content-disposition", attachment)
                .Response.ContentType = "application/vnd.ms-excel"
                Dim tb As String = String.Empty
                For Each dtcol As DataColumn In dTbl.Columns

                    .Response.Write(tb & dtcol.ColumnName)
                    tb = ","
                Next

                .Response.Write(vbCrLf)
                For Each dr As DataRow In dTbl.Rows

                    tb = ""
                    Dim j As Integer
                    For j = 0 To dTbl.Columns.Count - 1

                        If Regex.IsMatch(dr(j).ToString, ",") Then
                            'RFC4180 requires any field containing a comma to be enclosed in quotes at each end of field value
                            .Response.Write(tb & """" & dr(j).ToString & """")
                        Else
                            .Response.Write(tb & dr(j).ToString)
                        End If

                        tb = ","
                    Next
                    .Response.Write(vbCrLf)
                Next
                .Response.End()

            End With
        Catch ex As Exception
            '*** report errors only if trace is enabled.  in production you will always get a thread abort error and we don't want this in the report
            If myPage.Trace.IsEnabled Then myPage.Response.Write(ex.ToString)

        End Try
    End Sub

#End Region

#Region "...CONTROLS AND BINDING..."
    Public Function sortACDC(ByVal dg As DataGrid, ByVal sSortExp As String)
        '*** modifies col header to include &uarr; or &darr; as well as toggling the
        '*** return state.  If ASC or DESC is provided in sSortExp this is the default starting dir
        Dim myC As DataGridColumn = Nothing
        Dim myC1 As DataGridColumn
        For Each myC1 In dg.Columns
            If myC1.SortExpression = sSortExp Then
                myC = myC1
            Else
                myC1.HeaderText = System.Text.RegularExpressions.Regex.Replace(myC1.HeaderText, "&uarr;|&darr;", "")
            End If
        Next

        If System.Text.RegularExpressions.Regex.IsMatch(sSortExp, ",") Then Return sSortExp '*** bail for complex sort expressions
        If myC Is Nothing Then Return sSortExp

        '*** If col has existing arrow, swap direction and toggle sort expression direction
        '*** Note sSortExp may NOT contain ASC or DESC so you cannot simply search for one of these
        '*** and replace it.  istead you have to append ASC or DESC
        If System.Text.RegularExpressions.Regex.IsMatch(myC.HeaderText, "&uarr;") Then
            myC.HeaderText = System.Text.RegularExpressions.Regex.Replace(myC.HeaderText, "&uarr;", "&darr;")
            sSortExp = System.Text.RegularExpressions.Regex.Replace(sSortExp, " ASC| DESC", "")
            sSortExp += " DESC"
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(myC.HeaderText, "&darr;") Then
            myC.HeaderText = System.Text.RegularExpressions.Regex.Replace(myC.HeaderText, "&darr;", "&uarr;")
            sSortExp = System.Text.RegularExpressions.Regex.Replace(sSortExp, " ASC| DESC", "")
            sSortExp += " ASC"
            '*** if no arrows are present, means this is the first time we sort this col, so use its default
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(sSortExp, "DESC") Then
            myC.HeaderText += "&nbsp;&darr;"
        Else '*** ASC, or no sort direction provided
            myC.HeaderText += "&nbsp;&uarr;"
        End If
        Return sSortExp
    End Function
    Public Function sortACDC(ByVal gv As GridView, ByVal sSortExp As String)
        '*** overload version. modifies col header to include &uarr; or &darr; as well as toggling the
        '*** return state.  If ASC or DESC is provided in sSortExp this is the default starting dir

        '*** IMPORTANT:  you need to set htmlEncode="false" on the column, else the gridview will escape the &uarr; and it will 
        '*** appear as &ampuarr; instead of the arrow character.  This will have a flow on impact for any characters displayed in the
        '*** column data fields also.
        '*** http://codeverge.com/asp.net.presentation-controls/gridview-and-special-characters/470523

        Dim myC As DataControlField = Nothing  'column in gridview
        Dim myC1 As DataControlField
        For Each myC1 In gv.Columns
            If myC1.SortExpression = sSortExp Then
                myC = myC1
            Else
                myC1.HeaderText = System.Text.RegularExpressions.Regex.Replace(myC1.HeaderText, "&uarr;|&darr;", "")
            End If
        Next

        If System.Text.RegularExpressions.Regex.IsMatch(sSortExp, ",") Then Return sSortExp '*** bail for complex sort expressions
        If myC Is Nothing Then Return sSortExp

        '*** If col has existing arrow, swap direction and toggle sort expression direction
        '*** Note sSortExp may NOT contain ASC or DESC so you cannot simply search for one of these
        '*** and replace it.  istead you have to append ASC or DESC
        If System.Text.RegularExpressions.Regex.IsMatch(myC.HeaderText, "&uarr;") Then
            myC.HeaderText = System.Text.RegularExpressions.Regex.Replace(myC.HeaderText, "&uarr;", "&darr;")
            sSortExp = System.Text.RegularExpressions.Regex.Replace(sSortExp, " ASC| DESC", "")
            sSortExp += " DESC"
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(myC.HeaderText, "&darr;") Then
            myC.HeaderText = System.Text.RegularExpressions.Regex.Replace(myC.HeaderText, "&darr;", "&uarr;")
            sSortExp = System.Text.RegularExpressions.Regex.Replace(sSortExp, " ASC| DESC", "")
            sSortExp += " ASC"
            '*** if no arrows are present, means this is the first time we sort this col, so use its default
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(sSortExp, "DESC") Then
            myC.HeaderText += "&nbsp;&darr;"
        Else '*** ASC, or no sort direction provided
            myC.HeaderText += "&nbsp;&uarr;"
        End If
        Return sSortExp
    End Function
    Sub setColVis(ByVal g As DataGrid, ByVal c As String, ByVal v As Boolean)
        '*** 2015-05-04 overloaded version
        '*** show/hides the column with headerText=c
        Dim myC As DataGridColumn
        For Each myC In g.Columns
            If myC.HeaderText.ToUpper = c.ToUpper Then myC.Visible = v : Exit For
        Next

    End Sub
    Sub setColVis(ByVal g As GridView, ByVal c As String, ByVal v As Boolean)
        '*** 2015-05-04 overloaded version
        '*** show/hides the column with headerText=c
        Dim myC As DataControlField
        For Each myC In g.Columns
            If myC.HeaderText.ToUpper = c.ToUpper Then myC.Visible = v : Exit For
        Next
    End Sub
    Sub setDropDown(ByVal oDD As DropDownList, Optional ByVal oTable As DataTable = Nothing, Optional ByVal sVal As String = "")
        '*** 2014-11-17 modified to support optional 'bind' attribute on the DD, making it work same way as the doBindDataRow function
        '***  attribute bind usage; bind="text|value [nobind] [legacy] [blank]" where;
        '*** oldV attribute will hold existing value, oldT existing text.  these are bound on the aspx page, but sVal on the call will override
        '*** if no bind attribute is provided, oldT|oldV are used and a blank row is added

        If Not oTable Is Nothing Then
            oDD.DataSource = oTable
            oDD.DataBind()
        End If

        '*** backwards compatibility support
        If sVal = String.Empty Then
            If oDD.Attributes("oldV") & String.Empty <> String.Empty Then
                sVal = oDD.Attributes("oldV")
                If Not oDD.Items.FindByValue(sVal) Is Nothing Then oDD.Items.FindByValue(sVal).Selected = True 'backwards compatability support
            ElseIf oDD.Attributes("oldT") & String.Empty <> String.Empty Then
                sVal = oDD.Attributes("oldT") 'backwards compatability support
                If Not oDD.Items.FindByText(sVal) Is Nothing Then oDD.Items.FindByText(sVal).Selected = True
            End If
        End If

        '*** new code block, if bind attribute was provided, follow this
        If Not Regex.IsMatch(oDD.Attributes("bind") & String.Empty, "nobind", RegexOptions.IgnoreCase) Then
            Dim oItem As ListItem
            Dim blLegacy As Boolean = Regex.IsMatch(oDD.Attributes("bind") & String.Empty, "legacy", RegexOptions.IgnoreCase)
            Dim blBlank As Boolean = Regex.IsMatch(oDD.Attributes("bind") & String.Empty, "blank", RegexOptions.IgnoreCase)
            '*** 2009-12-14 some changes to the bind parameter and legacy parameter
            If Regex.IsMatch(oDD.Attributes("bind") & String.Empty, "value", RegexOptions.IgnoreCase) Then
                '*** find by value, but first check wether we are supporting legacy values (i.e. those not bound in the list)
                If oDD.Items.FindByValue(sVal) Is Nothing And blLegacy Then
                    oItem = New ListItem(sVal, sVal)
                    oDD.Items.Add(oItem) '*** add a legacy item
                    '*** now also add a blank if required (and hasn't just been added as a legacy item)
                Else
                    oItem = oDD.Items.FindByValue(sVal)
                End If
                oDD.SelectedIndex = oDD.Items.IndexOf(oItem)
            Else
                '*** default is to find by text
                If oDD.Items.FindByText(sVal) Is Nothing And blLegacy Then
                    oItem = New ListItem(sVal, sVal)
                    oDD.Items.Add(oItem)
                Else
                    oItem = oDD.Items.FindByText(sVal)
                End If
                oDD.SelectedIndex = oDD.Items.IndexOf(oItem)
            End If  '*** value test

            '*** now add a blank if required, and select it if required
            If blBlank Then
                '*** add a blank if one does not already exist
                If CBool(oDD.Items.FindByValue(String.Empty) Is Nothing) Then oDD.Items.Add(New ListItem(String.Empty, String.Empty))
                '*** If we do not have a valid oItem from before, then select this blank value
                If oItem Is Nothing Then
                    oItem = oDD.Items.FindByValue(String.Empty)
                    oDD.SelectedIndex = oDD.Items.IndexOf(oItem)
                End If
            End If '*** blank 
        End If
    End Sub
#End Region


#Region "...HELPER FUNCTIONS..."

    Function fnSpaceToNull(ByVal sText As String) As Object
        If sText = String.Empty Then Return DBNull.Value
        Return sText
    End Function
    Function fnWkgDays(ByVal d1 As Date, ByVal d2 As Date) As Long
        Dim n As Long
        Dim res As Long = 0
        For n = 1 To DateDiff("d", d1, d2)
            If Weekday(DateAdd("d", n, d1)) > 1 And Weekday(DateAdd("d", n, d1)) < 7 Then res += 1
        Next
        Return res
    End Function
    Function adjTZO(ByVal x As Object, ByVal timeOffset As String) As Object
        '*** adjusts for timezone
        Try
            If IsDate(x) Then
                x = CDate(x).AddHours(-1 * timeOffset)
            End If
        Catch
        End Try
        Return x
    End Function
    Function sqlSafe(ByVal s As String) As String
        '*** call when using dataview.filter this routine will escape apostrophies and other problmmatic chars
        Try
            s = s.Replace("'", "''")
            Return s.Replace("""", """""")

            '*** we don't need to worry about ; because .filter does not support termination of sql commands with a following one
            'e.g.  command 1; command 2
        Catch
            Return String.Empty
        End Try


    End Function
    Sub StreamFile(ByVal sFullPath As String, ByVal pg As Page)
        '*** will stream the target file to the client browser.  This has the advantage that .xlsx files will work properly over https://
        '*** the problem we had with hyperlinks was that the .xlsx file got corrupted over https and we had to server over http instead

        '*** 2014-10-31 fix for Firefox
        '*** http://techblog.procurios.nl/k/news/view/15872/14863/mimetype-corruption-in-firefox.html
        '*** the fix is below.  THIS WORKS on IE, FF and Chrome

        Dim tgByte() As Byte = Nothing
        With pg.Response
            Try
                Dim tgFStream As New IO.FileStream(sFullPath, IO.FileMode.Open, IO.FileAccess.Read)
                Dim tgBinaryReader As New IO.BinaryReader(tgFStream)
                tgByte = tgBinaryReader.ReadBytes(Convert.ToInt32(tgFStream.Length))
                '*** write the response
                .Clear()
                .OutputStream.Write(tgByte, 0, tgByte.Length)

                '*** strange bug with Firefox downloads.  If the filename contains spaces, FF will break the filename at the first space, lose the file suffix
                '*** and won't know how to open the file.  If we server.URLencode it, the spaces come across as + symbols, the file is correctly recognised but
                '*** we obviously have + now instead of space.
                '*** http://stackoverflow.com/questions/93551/how-to-encode-the-filename-parameter-of-content-disposition-header-in-http
                '*** The theoretically correct syntax for use of UTF-8 in Content-Disposition is just crazy: filename*=UTF-8''foo%c3%a4 (yes, that's an asterisk, and no quotes except an empty single quote in the middle)
                '*** YES, this does work with FireFox but we also need to use Uri.EscapeDataString to encode spaces as %20 rather than + (which is what Server.EncodeURL does).
                .AddHeader("Content-Disposition", "attachment; filename*=UTF-8''" & Uri.EscapeDataString(System.IO.Path.GetFileName(tgFStream.Name)))
                .AddHeader("Content-Length", tgByte.Length.ToString())
                .ContentType = "application/octet-stream"

                .End()
                tgBinaryReader.Close()
                tgFStream.Close()

            Catch ex As Exception
                writeAudit(ex.ToString, "streamfile")
                .Write("Sorry an error occured retreiving the file")
            Finally

            End Try
        End With
    End Sub

#End Region


#Region "...EMAIL..."
    Function sendMail(ByVal sTo As String, ByVal sCC As String, ByVal sFrom As String, ByVal sSubject As String, ByVal sBody As String) As Boolean
        '*** 2014-05-28 re-written for net 2.  Need to build an overload version to handle attachments
        '*** 2014-06-27 the mail server host name is only accessible through the client object, so the TEST routine has been moved lower down

        Using myMail As New System.Net.Mail.MailMessage
            Try
                If sFrom.Trim = String.Empty Then
                    myMail.From = New System.Net.Mail.MailAddress("noReplies@au.verizon.com", "PCM server no replies")
                Else
                    myMail.From = New System.Net.Mail.MailAddress(sFrom)
                End If

                '*** 2014-06-27 sTo and sCC might be ; separated strings.  We have to process each member and add separately to the mail object
                For Each s As String In Split(sTo, ";")
                    If s.Trim <> String.Empty Then myMail.To.Add(New System.Net.Mail.MailAddress(s))
                Next

                For Each s As String In Split(sCC, ";")
                    If s.Trim <> String.Empty Then myMail.CC.Add(New System.Net.Mail.MailAddress(s))
                Next

                myMail.Subject = sSubject
                myMail.Body = sBody
                myMail.BodyEncoding = System.Text.Encoding.ASCII
                ' MyMail.BodyFormat = System.Web.Mail.MailFormat.Text
                myMail.Priority = System.Net.Mail.MailPriority.High
                Dim myClient As New System.Net.Mail.SmtpClient
                '*** the client will pick up the smtp server address from the web.config file
                If myClient.Host = "TEST" Then
                    writeAudit(String.Concat("to:", sTo, vbCrLf, "cc:", sCC, vbCrLf, "From:", sFrom, vbCrLf, "subject:", sSubject, vbCrLf, "body:", sBody), "TEST_email")
                    Return True
                Else
                    myClient.Send(myMail)
                    Return True
                End If


            Catch ex As Exception
                writeAudit(ex.ToString, "sendmail")
                writeAudit(String.Concat("to:", sTo, vbCrLf, "cc:", sCC, vbCrLf, "From:", sFrom, vbCrLf, "subject:", sSubject, vbCrLf, "body:", sBody), "TEST_email")
                Return False
            Finally
            End Try
        End Using
    End Function
    Function sendMail(ByVal sTo As String, ByVal sCC As String, ByVal sFrom As String, ByVal sSubject As String, ByVal sBody As String, ByVal sAttachmentPaths As String) As Boolean
        '*** overload version of sendMail to handle attachments.  These must be full paths separated by a comma.
        '*** 2014-06-27 the mail server host name is only accessible through the client object, so the TEST routine has been moved lower down
        Using myMail As New System.Net.Mail.MailMessage
            Try
                If sFrom.Trim = String.Empty Then
                    myMail.From = New System.Net.Mail.MailAddress("noReplies@au.verizon.com", "PCM server no replies")
                Else
                    myMail.From = New System.Net.Mail.MailAddress(sFrom)
                End If

                '*** 2014-06-27 sTo and sCC might be ; separated strings.  We have to process each member and add separately to the mail object
                For Each s As String In Split(sTo, ";")
                    If s.Trim <> String.Empty Then myMail.To.Add(New System.Net.Mail.MailAddress(s))
                Next

                For Each s As String In Split(sCC, ";")
                    If s.Trim <> String.Empty Then myMail.CC.Add(New System.Net.Mail.MailAddress(s))
                Next

                myMail.Subject = sSubject
                myMail.Body = sBody
                myMail.BodyEncoding = System.Text.Encoding.ASCII
                ' MyMail.BodyFormat = System.Web.Mail.MailFormat.Text
                myMail.Priority = System.Net.Mail.MailPriority.High

                For Each sPath As String In sAttachmentPaths.Split(",")
                    myMail.Attachments.Add(New System.Net.Mail.Attachment(sPath))
                Next

                Dim myClient As New System.Net.Mail.SmtpClient
                '*** the client will pick up the smtp server address from the web.config file
                If myClient.Host = "TEST" Then
                    writeAudit(String.Concat("to:", sTo, vbCrLf, "cc:", sCC, vbCrLf, "From:", sFrom, vbCrLf, "subject:", sSubject, vbCrLf, "body:", sBody), "TEST_email")
                    Return True
                Else
                    myClient.Send(myMail)
                    Return True
                End If

            Catch ex As Exception
                writeAudit(ex.ToString, "sendMailWithAttachment")
                writeAudit(String.Concat("to:", sTo, vbCrLf, "cc:", sCC, vbCrLf, "From:", sFrom, vbCrLf, "subject:", sSubject, vbCrLf, "body:", sBody, vbCrLf, "sAttachmentPaths:", sAttachmentPaths), "TEST_email")
                writeAudit(myMail.To.ToString, "myMail.To.ToString")
                writeAudit(myMail.From.ToString, "myMail.From.ToString")
                writeAudit(myMail.CC.ToString, "myMail.cc.ToString")
                Return False
            Finally
            End Try
        End Using
    End Function
#End Region




#Region "...DEPRECATED CODE..."

    Public Function URLbeautifier(ByVal myPage As Page) As String
        '*** 2015-06-09 Deprecated, kept here for backwards compatibility in the CPE Workflow app

        '*** strips infopath and or querystring, redirects to clean the url 
        '*** else, returns the cleaned up param

        URLbeautifier = String.Empty '*** default return value to keep net2 compiler happy

        '*** Page onLoad event needs to check the beautifier session objects
        '2006-06-26 URL pathinfo rewritten
        '*** if you enter with CPEworkflow_Main.aspx/SP1234 then page will redirect to a beautified URL
        'http://weblogs.asp.net/scottgu/archive/2007/02/26/tip-trick-url-rewriting-with-asp-net.aspx
        With myPage
            .Trace.Warn("URL beautifier entry point. Postback=" & .IsPostBack)
            '*** deal with query strings, remove these by redirecting to a clean page
            If .Request("q") & String.Empty <> String.Empty Then
                '*** legacy support for ?q=SP1235, but we need to strip the querystring
                .Session.Add("CPEworkflow_param", .Request("q"))
                .Response.Redirect("CPEworkflow_Main.aspx", True)
            ElseIf .Request("t") & String.Empty <> String.Empty Then
                '*** legacy support for ?t=9999, but we need to strip the querystring
                .Session.Add("CPEworkflow_param", .Request("t"))
                .Response.Redirect("CPEworkflow_Editor.aspx", True)
            End If

            '*** look at the pathinfo
            If .Request.PathInfo.Length > 0 Then
                '*** Test for pathinfo, e.g. _Main.aspx/SP1234 or _Main.aspx/1234
                Dim m As Match = Regex.Match(.Request.PathInfo, "(SP)?(\d{2,})", RegexOptions.IgnoreCase)
                If m.Success Then
                    If m.Groups(1).Value.ToUpper = "SP" Then
                        '*** beautify the URL and invoke a redirect
                        .Session.Add("CPEworkflow_param", "SP" & m.Groups(2).ToString)
                        .Response.Redirect("CPEworkflow_Main.aspx", True)
                    Else
                        .Session.Add("CPEworkflow_param", m.Groups(2).ToString)
                        .Response.Redirect("CPEworkflow_Editor.aspx", True)
                    End If
                Else
                    '*** we have pathinfo but must be spurious data, eg user did not hack URL correctly
                    .Response.Redirect("CPEworkflow_Main.aspx", True)
                End If
            End If

            '*** return CPEworkflow_param if found
            If Not .Session("CPEworkflow_param") Is Nothing Then
                URLbeautifier = .Session("CPEworkflow_param")
                .Session.Remove("CPEworkflow_param")
                .Trace.Warn("URL beautifier exit1" & URLbeautifier)
            End If
            .Trace.Warn("URL beautifier exit2")
        End With
    End Function
    Sub setDGcolVis(ByVal dg As DataGrid, ByVal c As String, ByVal v As Boolean)
        '*** 2015-05-04 Deprecated
        '*** show/hides the column with headerText=c
        Dim myC As DataGridColumn
        For Each myC In dg.Columns
            If myC.HeaderText.ToUpper = c.ToUpper Then myC.Visible = v : Exit For
        Next
    End Sub
    Sub setGVcolVis(ByVal gv As GridView, ByVal c As String, ByVal v As Boolean)
        '*** 2015-05-04 deprecated
        '*** show/hides the column with headerText=c
        Dim myC As DataControlField
        For Each myC In gv.Columns
            If myC.HeaderText.ToUpper = c.ToUpper Then myC.Visible = v : Exit For
        Next
    End Sub
    Sub setDropDownLegacy(ByVal oDD As DropDownList, Optional ByVal oTable As DataTable = Nothing)
        '*** DEPRECATED.  Use setDropDown instead
        '*** pass a DDlist, with oldV or oldT attribute and this sub will set the dropdown
        '*** note that x will be case sensitive.  Also binds the dataset to the list.
        '*** If current OldV/D is not in the bound list, it is created thus allowing us to display retired values
        Dim x As String
        If Not oTable Is Nothing Then
            oDD.DataSource = oTable
            oDD.DataBind()
            '*** add a blank value as an option
            oDD.Items.Add(New ListItem(String.Empty, String.Empty))
        End If

        Try
            x = oDD.Attributes("oldT")
            Dim y As ListItem = Nothing
            If x <> "" Then y = oDD.Items.FindByText(x)
            x = oDD.Attributes("oldV")
            If x <> "" Then y = oDD.Items.FindByValue(x)

            If x = "" Then
                oDD.Items.FindByValue("").Selected = True
            ElseIf y Is Nothing Then
                y = New ListItem(x, x)
                y.Selected = True
                oDD.Items.Add(y)
            Else
                y.Selected = True
            End If
        Catch
        End Try

    End Sub

    Public Function checkNTLMUser(ByVal myPage As Page, Optional ByVal blRefresh As Boolean = False, Optional ByVal blPathInfo As Boolean = False, Optional ByVal tblName As String = "tblUserPermission") As Boolean
        '*** verifies NTLM user is fully validated and refreshes the session variables if required.
        '*** Will substitute a test user for a valid user as controlled by the altUser string in web.config
        '*** for security purposes will only substitute the altUser if the parent user is authenticated
        '*** if validation fails, it will halt page execution
        '*** blRefresh will force routine to refresh session variables from the database
        '*** blPathInfo will strip path info and redirect the page to itself thus cleaning the URL; it puts the path info into a session var.

        '*** BUG 2015-07-03 think about dates, when user is registering LastLoginUTC is null, but should this stop us refreshing all other 
        '*** session vars, and should we return true or false?
        '*** 2015-09-17 AccountLocked now will force a false return value



        With myPage
            '*** 2015-06-09 deal with PathInfo first
            If blPathInfo Then
                If .Request.PathInfo.Length > 0 Then
                    'capture the pathInfo and add it to the Session object.  This is the easiest way to make it survive the 
                    'the redirect which we use to clean up the URL
                    '*** 2018-03-26 added URL decode to protect against XSS
                    .Session.Add("PATHINFO", HttpUtility.UrlDecode(.Request.PathInfo))
                    .Response.Redirect(.Request.ServerVariables("URL"))
                End If
            End If



            '1/ If session is still valid, exit true unless we are forcing a session vars refresh
            If (Not .Session("AUTHUSER") Is Nothing) And blRefresh = False Then
                '*** 2015-07-03 bug fix.  Users > 120 days inactive OR with no LastLoginUTC will return false
                If Not IsDate(.Session("LASTLOGINUTC")) Then Return False

                '*** users >120 days since last login will return false
                If (DateDiff(DateInterval.Day, .Session("LASTLOGINUTC"), Date.UtcNow)) > 120 Then Return False

                '*** 2015-09-17  If AccountLocked return false
                If CBool(.Session("ACCOUNTLOCKED")) Then Return False

                Return True
            End If

            '2/ if NTLM fails to authenticate, terminate the app.  Don't want to use the test user in this scenario because its a security risk
            'any NTLM failure will lead to all users becoming the test user.
            If .User.Identity.IsAuthenticated = False Then
                'terminate the page
                .Response.Write("<b=""red"">FATAL ERROR:  NTLM User cannot be identified.</b>  Contact system administrator and report this error.")
                .Response.End()
            End If

            '2a/ possible alt user
            '**** does User.Identity.Name match into altUser?
            If ConfigurationManager.AppSettings("altUser") Is Nothing Then
                '*** no altUser so go with Identity
                .Session("AUTHUSER") = .User.Identity.Name.ToString
                '*** look for User.Identity.Name in altUser, use Instr because regex will be confused by the \
                '*** old block would match short strings rather than exact
                'ElseIf InStr(ConfigurationManager.AppSettings("altUser").ToString, .User.Identity.Name.ToString, CompareMethod.Text) = 1 Then

                '*** new block will match full string prior to the colon
            ElseIf String.Equals(.User.Identity.Name.ToString, Regex.Replace(ConfigurationManager.AppSettings("altUser").ToString, "([^\x5c]+\x5c\w+):(\w+)$", "$1"), _
                                 StringComparison.CurrentCultureIgnoreCase) = True Then


                '*** matches, so substitute the second group vID with the part after the colon. \x5C is a \ char
                .Session("AUTHUSER") = Regex.Replace(ConfigurationManager.AppSettings("altUser").ToString, "([^\x5c]+)\x5c(\w+):(\w+)$", "$1\$3")
            Else
                '*** does not match, use the NTLM user 
                .Session("AUTHUSER") = .User.Identity.Name.ToString
            End If

            '3/  refresh the system vars
            '*** all good, load session vars
            Dim oConn As New OleDb.OleDbConnection(sConn)

            Try
                Dim objCmd As New OleDb.OleDbCommand("SELECT * FROM " & tblName & " WHERE AUTHUSER=@p1", oConn)
                objCmd.Parameters.Add("@p1", OleDb.OleDbType.VarChar).Value = .Session("AUTHUSER")
                oConn.Open()
                Dim objRead As OleDb.OleDbDataReader = objCmd.ExecuteReader(CommandBehavior.CloseConnection)

                If objRead.Read Then
                    Dim n As Integer
                    For n = 0 To objRead.FieldCount - 1
                        '*** note that all session vars are UPPER CASE to avoid case problems, even though in asp.net
                        '*** session keys are case insensitive
                        If objRead.Item(n) Is DBNull.Value Then
                            '*** map null to string.empty to help with regex tests later
                            .Session(objRead.GetName(n).ToUpper) = String.Empty
                        Else
                            .Session(objRead.GetName(n).ToUpper) = objRead.Item(n)
                        End If
                    Next
                    objRead.Close()
                End If

                '*** 2015-09-17  If AccountLocked return false
                If CBool(.Session("ACCOUNTLOCKED")) Then Return False


                '*** 2015-05-11 For CPS108 compliance, we should deny users with >120 day access. To do this we'd simply test last login UTC
                '*** and return false at this step, not update the lastlogin value.
                '  .Trace.Warn("chckNTLM " & DateDiff(DateInterval.Day, .Session("LASTLOGINUTC"), Date.UtcNow))

                '*** 2015-07-03 bug fix.  Users > 120 days inactive OR with no LastLoginUTC will return false
                If Not IsDate(.Session("LASTLOGINUTC")) Then Return False

                If (DateDiff(DateInterval.Day, .Session("LASTLOGINUTC"), Date.UtcNow)) > 120 Then Return False
                '*** these session vars are not changed.  Main program code must look at the LASTLOGINUTC as a possible reason for the reject.
                '*** to reset the 120 day lockout, the admin must re-update a locked user.

                '*** 2015-05-11 update the lastlogin value
                objCmd = New OleDb.OleDbCommand("Update " & tblName & " SET LastLoginUTC=@p1 WHERE AUTHUSER=@p2", oConn)

                objCmd.Parameters.Add("@p1", OleDb.OleDbType.Date).Value = DateTime.UtcNow
                objCmd.Parameters.Add("@p2", OleDb.OleDbType.VarChar).Value = .Session("AUTHUSER")
                oConn.Open()
                objCmd.ExecuteNonQuery()
                oConn.Close()

                Return True
            Catch ex As Exception
                .Trace.Warn(ex.ToString)
                Return False
            Finally
                oConn.Dispose()
            End Try

        End With
    End Function
    Function checkSession(ByVal myPage As Page, ByVal appName As String, ByVal sTablePermission As String) As String
        '*** re-written 2010-06-17.  Fixed bugs relating to cookie test, you must test cookie object exists before testing an attribute
        '*** this app does not support GUEST access

        '*** MAKE GENERIC 2014-06-15
        '*** appName is the name used for the application specific cookies
        '*** sTablePermission is the app local permissions table.  It MUST contain NTID and LastLoginUTC fields.
        '*** all field names and values from sTablePermission will be loaded into the session object
        '*** sTablePermission can be an updateable query

        '*** 2014-09-03 added cookieTimeout

        '*** 2015-06-09 this is a forms based login, and will be superceeded by checkNTLMuser

        Dim ObjConn As New OleDb.OleDbConnection(sConn)

        With myPage
            .Trace.Warn("checkSession entry")
            Dim cookieTimeout As Long = 44640 'one month in mins

            '*** CP810 override, can set a specific cookie timeout
            If IsNumeric(ConfigurationManager.AppSettings("cookieTimeout") & String.Empty) Then
                cookieTimeout = CLng(ConfigurationManager.AppSettings("cookieTimeout"))
            End If

            Try
                If .Session("NTID") & String.Empty <> String.Empty Then
                    '*** session ok, so persist in the cookie
                    .Response.Cookies(appName).Item("NTID") = .Session("NTID")
                    .Response.Cookies(appName).Expires = DateTime.Now.AddMinutes(cookieTimeout)
                Else
                    '*** no session, so check cookie
                    If .Request.Cookies(appName) Is Nothing Then
                        '*** no cookie to test, we have no session either so force a login
                        Return False
                    Else
                        '*** cookie was ok, so use this to reset the session
                        .Response.Cookies.Set(.Request.Cookies(appName))
                        .Session("NTID") = .Response.Cookies(appName).Item("NTID")
                        .Response.Cookies(appName).Expires = DateTime.Now.AddMinutes(cookieTimeout)
                    End If
                End If

                '*** so now we have a good session(NTID), what about the other session vars?
                .Trace.Warn("checkSession have a good NTID")


                '*** At this point, Session NTID will be valid.  Reload other session vars if these have expired
                '*** Because we don't know what field names are in sTablePermission, we test on the mandatory LastLoginUTC field
                '*** 2014-11-06 bug fix, we need to test ALL the table field names because some, such as LastLoginUTC are shared across apps.
                '*** since we wish to avoid reloading the table if we can, we instead use an appName marker

                If .Session(appName) & String.Empty = String.Empty Then

                    '*** re-run the query on sTablePermission and load all fields into the session object
                    ObjConn.Open()
                    Dim objCmd As New OleDb.OleDbCommand("SELECT * FROM " & sTablePermission & " WHERE NTID=@p1", ObjConn)
                    objCmd.Parameters.Add("@p1", OleDb.OleDbType.VarChar).Value = .Session("NTID")
                    Dim objRead As OleDb.OleDbDataReader = objCmd.ExecuteReader(CommandBehavior.CloseConnection)

                    If objRead.Read Then
                        Dim n As Integer
                        For n = 0 To objRead.FieldCount - 1
                            '*** note that all session vars are UPPER CASE to avoid case problems, even though in asp.net
                            '*** session keys are case insensitive
                            If objRead.Item(n) Is DBNull.Value Then
                                '*** map null to string.empty to help with regex tests later
                                .Session(objRead.GetName(n).ToUpper) = String.Empty
                            Else
                                .Session(objRead.GetName(n).ToUpper) = objRead.Item(n)
                            End If
                        Next
                        objRead.Close()

                        ObjConn.Open()
                        '*** 2013-05-29 modified log the user in tblOwner, this is done once per expired session
                        objCmd = New OleDb.OleDbCommand("UPDATE " & sTablePermission & " SET LastLoginUTC=@p1 WHERE NTID=@p2", ObjConn)
                        objCmd.Parameters.Add("@p1", OleDb.OleDbType.Date).Value = DateTime.UtcNow
                        objCmd.Parameters.Add("@p2", OleDb.OleDbType.VarChar).Value = .Session("NTID")
                        objCmd.ExecuteNonQuery()
                        '*** 2014-11-06 add the marker
                        .Session(appName) = appName

                    ElseIf .Session("NTID") = String.Empty Then
                        '*** user has not registered at all so must do so. Returning false will bounce them back to the login page where they
                        '*** can self register
                        Return False
                    Else
                        '*** if we can't find a valid user NTID entry in permissions table so exit as not valid.
                        Return ("REGISTER")
                    End If

                    ObjConn.Close()
                End If

                .Trace.Warn("checkSession8")
                Return True

            Catch ex As Exception
                writeAudit(ex.ToString, "ERROR_Checksession")
                If Not .Request.Cookies(appName).Item("NTID") Is Nothing Then .Response.Cookies(appName).Expires = DateTime.Now.AddMonths(-1)
                Return False
            Finally
                ObjConn.Dispose()
            End Try
        End With
    End Function


#End Region

End Module
