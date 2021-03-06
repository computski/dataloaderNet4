﻿Imports System.Data.OleDb
Imports System.Text
Imports System.IO

Public Class Analytics_dataLoad
    Inherits System.Web.UI.Page
    Dim sConn As String = ConfigurationManager.ConnectionStrings("sConn").ConnectionString
    Dim sData As String = ConfigurationManager.ConnectionStrings("sData").ConnectionString
    Dim sUserFiles As String = ConfigurationManager.AppSettings("strUserFiles")

    'notes
    'have a big problem with data import
    'reading csv into a dataset is fast, but the dataadaptor update is really slow at 80 rec/sec
    'options are
    '
    'link the table, then use bulk DAO3.6 update, however type coertion is a problem here
    'so we'd need to properly specify an import spec
    '
    'read a dataset then perform a series of discrete DAO record updates.  need to test this in access
    'as want to test whether you can wrap in a transaction
    '
    '





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnUpload_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnUpload.Click
        '*** run through the file upload list
        Dim sb As New StringBuilder
        sb.Append("Start upload...<br/>")
        Dim ImageFiles As HttpFileCollection = Request.Files
        Dim j As Integer = 1
        For i As Integer = 0 To ImageFiles.Count - 1
            Dim file As HttpPostedFile = ImageFiles(i)

            Trace.Warn(file.FileName)
            Trace.Warn(file.ContentType)




            If file.ContentLength = 0 Then
                sb.Append(" zero len file, ignored<br/>")
            ElseIf Regex.IsMatch(file.FileName, "\.CSV$", RegexOptions.IgnoreCase) Then
                sb.Append("[")
                sb.Append(j)
                j += 1
                sb.Append("] ")
                sb.Append(Regex.Match(file.FileName, "[^\\]+$").ToString)

                Dim m As String = processPBfile(file.InputStream)
                sb.Append(" result:")
                sb.Append(m)
                sb.Append("<br/>")
            ElseIf Regex.IsMatch(file.FileName, "\.ZIP$", RegexOptions.IgnoreCase) Then
                'process a zip file.  do we just handle the first file therein, or all of them (better)
                Dim zip As New System.IO.Compression.ZipArchive(file.InputStream)
                sb.Append("found zip file, processing contents...<br/>")
                For Each f As System.IO.Compression.ZipArchiveEntry In zip.Entries
                    sb.Append("[")
                    sb.Append(j)
                    j += 1
                    sb.Append("] ")
                    sb.Append(f.Name)
                    Dim st As Stream = f.Open()
                    Dim m As String = processPBfile(st)
                    sb.Append(" result:")
                    sb.Append(m)
                    sb.Append("<br/>")

                Next
                sb.Append("zipfile processed.")
                sb.Append("<br/>")
            Else
                sb.Append("[")
                sb.Append(j)
                j += 1
                sb.Append("] ")
                sb.Append(Regex.Match(file.FileName, "[^\\]+$").ToString)
                sb.Append(" not a CSV or ZIP, ignored<br/>")
            End If

        Next
        sb.Append("Done...<br/>")

        litDataResult.Text = sb.ToString

    End Sub
    Function processPBfileTABLEDEF(ByVal fc As System.IO.Stream) As String
        '*** look at file, what's the majority billing period?  does that database table exist?  if no, create it, otherwise consider appending to it based on
        '*** metadata
        'WARNING: there is a max request size (and on IIS7 you need to set on server too).  system will bomb with no error page if it is exceeded.
        'in principle, we can demand all files over 10M be sent up as zips
        'https://stackoverflow.com/questions/288612/how-to-increase-the-max-upload-file-size-in-asp-net

        'ARGH, there are a lot of type conversion errors happening, really we need to load the ultimate target table
        'schema (tblSource) and parse these header rows against it.  we should also drop some cols to reduce data
        'e.g. "" wont convert to null when we try to copy the data over to a double later.

        'can either pull in the schema, apply this to column heads (via schema.contains) and then type convert as we load each row
        'or leave as text, and type convert at end, running down each column and converting values before finally
        'setting the type on the col (though we may not be able to do this on a loaded recordset)
        'Dim oDA As New OleDbDataAdapter("SELECT * FROM " & rMeta("Charge Period") & " WHERE ID=0", dConn)
        ' oDA.Fill(oDS, "target")

        'REVISED
        'LOAD a copy of the target table schema
        Dim oDS As New DataSet
        Dim oDA As New OleDbDataAdapter("SELECT * FROM tblSource WHERE ID=0", sData)
        oDA.Fill(oDS, "template")

        'now open the db file  https://stackoverflow.com/questions/411592/how-do-i-save-a-stream-to-a-file-in-c
        Dim fs As FileStream = File.Create(sUserFiles & "data.csv")
        fc.Seek(0, SeekOrigin.Begin)
        fc.CopyTo(fs)
        fs.Close()

        'now use DAO 3.6, create a linked tabledef in our db
        'https://documentation.help/MS-DAO-3.60/damthopendatabase.htm
        'https//social.msdn.microsoft.com/Forums/office/en-US/933f6025-be11-4444-8949-ccc2d5315283/vbnet-writing-to-access-tables-ado-vs-dao-vs-oledb?forum=accessdev

        Dim dbEngine As New DAO.DBEngine
        Dim db As DAO.Database = dbEngine.OpenDatabase("C:\Users\Julian\Documents\Visual Studio 2017\source\PCManalytics\primebillerdata.mdb", False)

        '2018-10-31 usiing oledb for updates is slow.  dao 3.6 is faster, but we have a problem.
        'if we link to the table, the datatypes are not co-erced for us.  therefore i cannot just
        'select into the target table due to type conversion errors.
        '
        'several poss fixes;
        '1. is to define all the types in MS access when setting up the table link.  i am not sure this can cope with date formats
        '2. is to import the data as csv and co-erce into dtI, from here we build a long SQL string of individual insert-values
        'statements each separated by a semi colon.  we then execute these with db.execte  (or can try executenonquery)
        'this should be faster.





        Try
            db.TableDefs.Delete("linkData")
        Catch ex As Exception
        End Try

        Dim tdfLink As DAO.TableDef = db.CreateTableDef("linkData")
        tdfLink.SourceTableName = "data.csv"
        tdfLink.Connect = "Text;FMT=Delimited;HDR=YES;IMEX=2;CharacterSet=850;DATABASE=" & sUserFiles
        'tdfLink.Connect = "Text;FMT=Delimited;HDR=NO;IMEX=2; CharacterSet=437;DATABASE=C:\Users\v817353\Documents\invoice control csv\Prime biller\ICR_2018\ICR_201806\"
        db.TableDefs.Append(tdfLink)


        'build an array of fieldnames
        Dim c As Integer
        Dim fl As New ArrayList()
        For c = 0 To tdfLink.Fields.Count - 1
            fl.Add(tdfLink.Fields(c).Name)
        Next

        Trace.Warn(fl(2).ToString)


        'tdfLink = Nothing

        'Now find meta data in this link table from a stored query
        oDA = New OleDbDataAdapter("SELECT * FROM qryMetaData", sData)
        oDA.Fill(oDS, "meta")
        gvDebug.DataSource = oDS.Tables("meta")
        gvDebug.DataBind()
        'for each target table, if does not exist, create it




        Dim dConn As New OleDbConnection(sData)
        Dim restrictions(3) As String
        restrictions(3) = "TABLE"

        '*** pull list of existing tables, there may be some temp tables in there
        'https://stackoverflow.com/questions/1699897/retrieve-list-of-tables-in-ms-access-file
        dConn.Open()
        Dim dtTable As DataTable = dConn.GetSchema("tables", restrictions)
        'dConn.Close()
        Dim thisPeriod As String
        For Each rMeta In oDS.Tables("meta").Rows
            '*** does this table exist already?
            Trace.Warn("processing records for " & rMeta("chargePeriod"))
            thisPeriod = rMeta("chargePeriod")

            Dim foundrows() As DataRow = dtTable.Select("TABLE_NAME='" & thisPeriod & "'")
            If foundrows.Count = 0 Then
                '*** create the table
                Dim oCmd As New OleDb.OleDbCommand(String.Concat("SELECT tblSource.* INTO ", thisPeriod, " FROM tblSource WHERE ID=0"), dConn)

                'dConn.Open()
                'oCmd.Transaction = myTrans
                oCmd.ExecuteNonQuery()
                Trace.Warn("table created")
                '*** now set a primary key on the ID field, else commmandbuilder update will later fail
                'https://stackoverflow.com/questions/2470681/how-to-define-a-vb-net-datatable-column-as-primary-key-after-creation
                oCmd = New OleDbCommand("ALTER Table " & thisPeriod & " ADD PRIMARY KEY (ID);", dConn)
                oCmd.ExecuteNonQuery()
            End If
            '*** at this point we have either found the table, or created it, so either way continue because it exists

        Next

        'now build a set of queries based on our template table columnnames, matching these to the linkdata columns
        'do this within a transaction wrapper

        For Each rMeta In oDS.Tables("meta").Rows
            thisPeriod = rMeta("chargePeriod")

            Dim sb As New StringBuilder
            Dim sb2 As New StringBuilder
            sb.Append("INSERT INTO ")
            sb.Append(thisPeriod)
            sb.Append(" (")
            'target table fields here []
            For Each myC As DataColumn In oDS.Tables("template").Columns
                If Not Regex.IsMatch(myC.ColumnName, "ID|isfullperiod", RegexOptions.IgnoreCase) Then
                    sb.Append(myC.ColumnName)
                    sb.Append(",")
                    'look for this columnname in the linkData


                End If
            Next
            sb.Remove(sb.Length - 1, 1)
            sb.Append(") SELECT (")

            'nope, the order is all screwy maybe use triplets as a structure?  or define a mapping table
            'or loop it through from linkData, and use indexOf template.

            For c = 0 To fl.Count - 1
                Dim fname As String = Regex.Replace(fl(c).ToString, "[\s\W\.\?]+", String.Empty)
                If oDS.Tables("template").Columns.Contains(fname) Then
                    sb.Append("[")
                    sb.Append(fl(c).ToString)
                    sb.Append("],")
                End If

            Next
            sb.Remove(sb.Length - 1, 1)
            'source table fields here []
            sb.Append(") FROM linkData;")
            Trace.Warn(sb.ToString)
        Next


        'and execute


        Return 0








        'https://stackoverflow.com/questions/5065086/vb-net-how-can-i-check-if-a-primary-key-exists-in-an-access-db
        'you need to drill the table schema



        '*** attempted fixes
        'A. add sequencial ID after loading dtI
        '


        'Trace.Warn("primarykeyCount " & oDS.Tables("source").PrimaryKey.Count)

        '*** the target table will contain column data types and also allow us to cut down on the columns captured
        '*** prior textparser code could handle multiple instances of the same columnName.  we do not need to support this

        '1/ load dataset.  This is a revised text parser that will load data with reference to a target table schema
        'and co-erce data types as required, or map to dbnull.  If the target table does not have a matching column, the
        'text input data is ignored.  we load the data into a separate recordset dtI rather than directly into the target table
        '*** because we need to direct where the various records will go.
        Dim sR As StreamReader = New StreamReader(fc)
        Dim afile As FileIO.TextFieldParser = New FileIO.TextFieldParser(New StringReader(sR.ReadToEnd().ToString()))
        sR.Dispose()

        Dim CurrentRecord As String() ' this array will hold each line of data
        afile.TextFieldType = FileIO.FieldType.Delimited
        afile.Delimiters = New String() {","}
        afile.HasFieldsEnclosedInQuotes = True

        Dim dtI As New DataTable("source")
        Dim r As Long = 0
        'Dim c As Integer = 0
        Dim dr As DataRow

        Do While Not afile.EndOfData
            Try
                CurrentRecord = afile.ReadFields
                c = 0
                If r > 0 Then
                    dr = dtI.NewRow
                End If

                For Each s As String In CurrentRecord
                    '*** r=0 is the table header
                    If r = 0 Then
                        Dim inx As Integer
                        '*** remove spaces and non word chars from the 
                        s = Regex.Replace(s, "[\s\W\.\?]+", String.Empty)
                        inx = oDS.Tables("source").Columns.IndexOf(s.ToString)
                        If inx = -1 Then
                            dtI.Columns.Add(s & "_IGNORE")
                        Else
                            dtI.Columns.Add(s, oDS.Tables("source").Columns(inx).DataType)
                        End If

                    Else
                        'we are handling a row of data, need to process per column datatype
                        'skip the _IGNORE columns
                        If Not Regex.IsMatch(dtI.Columns(c).ColumnName, "IGNORE$") Then
                            dr.Item(c) = coerceType(s, dtI.Columns(c).DataType, dtI.Columns(c).MaxLength)
                        End If

                        c += 1
                    End If

                Next
                If r > 0 Then dtI.Rows.Add(dr)
                r += 1

            Catch ex As FileIO.MalformedLineException
                statusBar.InnerText = "ERROR the CSV file does Not conform to RFC8140"
                Exit Function
            Catch ex As Exception
                statusBar.InnerText = "Sorry an error occured. please check source file Is MC-08"
                Trace.Warn("processPBfile " & ex.ToString)
                writeAudit("processPBfile " & ex.ToString, Request("LOGON_USER"))
                Exit Function
            End Try
        Loop
        '*** done processing the file
        afile.Dispose()

        '*** drop the _IGNORE columns
        For c = dtI.Columns.Count - 1 To 0 Step -1
            If Regex.IsMatch(dtI.Columns(c).ColumnName, "IGNORE$") Then dtI.Columns.RemoveAt(c)
        Next
        '*** add blank ID column
        Dim dc As New DataColumn("ID", GetType(Integer))
        dc.AutoIncrement = True
        dtI.Columns.Add(dc)
        dc.SetOrdinal(0)

        '*** add a ssequence number
        r = 1
        For Each myR As DataRow In dtI.Rows
            myR("ID") = r
            r += 1
        Next

        dtI.Columns.Add(New DataColumn("isFullPeriod", GetType(Boolean)))
        dtI.AcceptChanges()


        '*** check columncounts match
        Trace.Warn(dtI.Columns.Count)
        Trace.Warn(oDS.Tables("source").Columns.Count)
        If dtI.Columns.Count <> oDS.Tables("source").Columns.Count Then Throw New ArgumentException("column count on import does Not match tblSource")

        '*** calculate metadata for this file
        Dim myView As DataView = dtI.DefaultView
        Dim a() As String = {"BillRunID", "ChargePeriod"}
        Dim dtMeta As DataTable = myView.ToTable(True, a)

        '*** calculate record counts
        dtMeta.Columns.Add(New DataColumn("count", GetType(Integer)))

        For Each myR As DataRow In dtMeta.Rows
            myR("count") = dtI.Compute("count([ChargePeriod])", "[ChargePeriod]='" & myR("ChargePeriod") & "'")
        Next

        '*** so now we have record counts per charge period.  find the largest and this is the base table we need to work with.
        '*** we will also back fill prior months, but only going back 2 months.

        myView = dtMeta.DefaultView
        myView.Sort = "[ChargePeriod] DESC"
        statusBar.InnerText = myView.Item(0).Item("ChargePeriod")



        '*** loop for the the top three entries in dtMeta, create a table as required, or look at the target table
        '*** metadata before writing to it
        r = 0

        '*** need to this next part as a TRANSACTION where we update multiple tables at once, or not at all
        '*** this means each table has to be given a new name
        dConn.Open()
        'Dim myTrans = dConn.BeginTransaction(IsolationLevel.ReadCommitted)

        'ARGH, for transactions, its ok to create empty tables even if we don't ultimately write to them
        'so really the transaction is only needed at the table-write time



        For Each rMeta In dtMeta.Rows
            '*** does this table exist already?
            Trace.Warn("processing records for " & rMeta("chargePeriod"))
            'Dim thisPeriod As String = rMeta("chargePeriod")

            Dim foundrows() As DataRow = dtTable.Select("TABLE_NAME='" & thisPeriod & "'")
            If foundrows.Count = 0 Then
                '*** create the table
                Dim oCmd As New OleDb.OleDbCommand(String.Concat("SELECT tblSource.* INTO ", thisPeriod, " FROM tblSource WHERE ID=0"), dConn)

                'dConn.Open()
                'oCmd.Transaction = myTrans
                oCmd.ExecuteNonQuery()
                Trace.Warn("table created")
                '*** now set a primary key on the ID field, else commmandbuilder update will later fail
                'https://stackoverflow.com/questions/2470681/how-to-define-a-vb-net-datatable-column-as-primary-key-after-creation
                oCmd = New OleDbCommand("ALTER Table " & thisPeriod & " ADD PRIMARY KEY (ID);", dConn)
                oCmd.ExecuteNonQuery()
            End If
            '*** at this point we have either found the table, or created it, so either way continue because it exists

            '*** now check the metadata on that table.  The count of records for this [Bill run ID] must be less than the count
            '*** in Meta data, else we would be double-loading
            '*** assumes Bill Run ID is a string
            oDA = New OleDbDataAdapter("SELECT * FROM " & thisPeriod & " WHERE ID=0", dConn)
            oDA.Fill(oDS, thisPeriod)


            Dim oCmd2 = New OleDbCommand(String.Concat("SELECT count([ID]) AS CountOfID FROM ", thisPeriod, " WHERE [BillRunID]=", CLng(rMeta("BillRunID"))), dConn)
            'dConn.Open()
            'oCmd2.Transaction = myTrans
            Dim existingCount As Integer = oCmd2.ExecuteScalar
            Trace.Warn("existing rec " & existingCount)
            'dConn.Close()

            '*** if existing count is less than our rMeta(count) then proceed with writing
            If existingCount < rMeta("count") Then
                'loop to write in the records
                For Each myR As DataRow In dtI.Rows

                    If myR("ChargePeriod").ToString = thisPeriod Then
                        '*** add data to appropriate charge-period table.  note that even though dtI has more cols than
                        '*** the target table, using itemArray will map across only those we need
                        Dim newR As DataRow = oDS.Tables(thisPeriod).NewRow()
                        'copyArray(myR.ItemArray, newR.ItemArray)
                        newR.ItemArray = myR.ItemArray
                        'newR("OPCO") = "YY"
                        oDS.Tables(thisPeriod).Rows.Add(newR)
                        'Trace.Warn("row added")
                        'ARGH the problem is dTI does not have same schmea as tbl(thisPeriod) as it had no IDcol
                        'I might be better off importing the data to tblSource as a dataset but then not writing it

                        'WHAAAAA??
                        'oDS.Tables(thisPeriod).ImportRow(myR)
                        'huh, does not add myR to the table.  

                        If newR.HasErrors Then
                            Trace.Warn("row error = " & r)

                        End If


                        r += 1
                    End If

                Next

                Trace.Warn("added the records " & r)
                '*** update the table, as a transaction
                Dim builder As New OleDb.OleDbCommandBuilder(oDA)
                builder.GetInsertCommand()
                builder.GetUpdateCommand()

                'oDA.UpdateCommand.Transaction = myTrans
                '*** now run the transaction itself
                oDA.ContinueUpdateOnError = False
                oDA.Update(oDS, thisPeriod)
                'oDS.Tables(0).HasErrors
                Trace.Warn("called oDA.update")
            Else
                Trace.Warn("records already exist for " & thisPeriod)
            End If

            Trace.Warn("table done")
            Trace.Warn(oDS.Tables(thisPeriod).HasErrors)

        Next

        gvDebug.DataSource = dtMeta
        gvDebug.DataBind()

        'last action is to commit all changes across multiple tables
        'myTrans.Commit()
        Trace.Warn("committed the transaction. END")
        dConn.Dispose()
        Return "OK, records=" & r


        'gvDebug.DataSource = dConn.GetSchema("tables", restrictions)
        'gvDebug.DataBind()
        ' dConn.Dispose()


        gvDebug.DataSource = dtMeta
        gvDebug.DataBind()


        '2./now find the base table. if it does not exist, create one. [via select into]

        'SELECT tblSource.* INTO tbl2 FROM(tblSource) WHERE (((tblSource.ID) Is Null));
        'so you can create a new table with the same schema by using tbl2 as a name.  the ID will restart at 1 I think.

        '3/ now as a transaction, add these new records after first checking meta data of target table.
        'do this for the top 3 months in the new source data





    End Function

    Function processPBfile(ByVal fc As System.IO.Stream) As String
        '*** look at file, what's the majority billing period?  does that database table exist?  if no, create it, otherwise consider appending to it based on
        '*** metadata
        'WARNING: there is a max request size (and on IIS7 you need to set on server too).  system will bomb with no error page if it is exceeded.
        'in principle, we can demand all files over 10M be sent up as zips
        'https://stackoverflow.com/questions/288612/how-to-increase-the-max-upload-file-size-in-asp-net
        Dim dtStart As DateTime = Now()


        'LOAD a copy of the target table schema
        Dim oConn As New OleDbConnection(sConn)
        Dim oDS As New DataSet
        Dim oDA As New OleDbDataAdapter("SELECT * FROM tblSource WHERE ID=0", sData)
        oDA.Fill(oDS, "template")



        'https://stackoverflow.com/questions/5065086/vb-net-how-can-i-check-if-a-primary-key-exists-in-an-access-db
        'you need to drill the table schema



        '*** get field lengths
        'https://blogs.technet.microsoft.com/heyscriptingguy/2007/10/03/hey-scripting-guy-how-can-i-retrieve-the-field-size-and-a-sample-record-for-all-the-tables-and-fields-in-an-office-access-database/
        '*** oledb and acceess don't populate fieldSize property, instead you have to pull from schema and then write it back
        '*** to the template table
        oConn.Open()
        Dim T As New DataTable
        T = getColLengths("tblSource", oConn)

        For Each myR As DataRow In T.Rows
            Dim myC = oDS.Tables("template").Columns.Item(myR("COLUMN_NAME").ToString)
            Trace.Warn(myC.ColumnName & " " & myC.DataType.ToString)
            If myC.DataType.ToString = "System.String" Then
                myC.MaxLength = myR("CHARACTER_MAXIMUM_LENGTH")
                Trace.Warn(myR("CHARACTER_MAXIMUM_LENGTH"))
            End If
        Next
        Trace.Warn("field lengths set")


        '*** the target table will contain column data types and also allow us to cut down on the columns captured
        '*** prior textparser code could handle multiple instances of the same columnName.  we do not need to support this

        '1/ load dataset.  This is a revised text parser that will load data with reference to a target table schema
        'and co-erce data types as required, or map to dbnull.  If the target table does not have a matching column, the
        'text input data is ignored.  we load the data into a separate recordset dtI rather than directly into the target table
        '*** because we need to direct where the various records will go.
        Dim sR As StreamReader = New StreamReader(fc)
        Dim afile As FileIO.TextFieldParser = New FileIO.TextFieldParser(New StringReader(sR.ReadToEnd().ToString()))
        sR.Dispose()

        Dim CurrentRecord As String() ' this array will hold each line of data
        afile.TextFieldType = FileIO.FieldType.Delimited
        afile.Delimiters = New String() {","}
        afile.HasFieldsEnclosedInQuotes = True

        Dim dtI As New DataTable("import")
        Dim r As Long = 0
        Dim c As Integer = 0
        Dim dr As DataRow

        Do While Not afile.EndOfData
            Try
                CurrentRecord = afile.ReadFields
                c = 0
                If r > 0 Then
                    dr = dtI.NewRow
                End If

                For Each s As String In CurrentRecord
                    '*** r=0 is the table header
                    If r = 0 Then
                        Dim inx As Integer
                        '*** remove spaces and non word chars from the 
                        s = Regex.Replace(s, "[\s\W\.\?]+", String.Empty)
                        inx = oDS.Tables("template").Columns.IndexOf(s.ToString)
                        If inx = -1 Then
                            dtI.Columns.Add(s & "_IGNORE")
                        Else
                            dtI.Columns.Add(s, oDS.Tables("template").Columns(inx).DataType)
                        End If

                    Else
                        'we are handling a row of data, need to process per column datatype
                        'skip the _IGNORE columns
                        If Not Regex.IsMatch(dtI.Columns(c).ColumnName, "IGNORE$") Then
                            dr.Item(c) = coerceType(s, dtI.Columns(c).DataType, dtI.Columns(c).MaxLength)
                        End If

                        c += 1
                    End If

                Next
                If r > 0 Then dtI.Rows.Add(dr)
                r += 1

            Catch ex As FileIO.MalformedLineException
                statusBar.InnerText = "ERROR the CSV file does not conform to RFC8140"
                Exit Function
            Catch ex As Exception
                statusBar.InnerText = "Sorry an error occured. please check source file is MC-08"
                Trace.Warn("processPBfile " & ex.ToString)
                writeAudit("processPBfile " & ex.ToString, Request("LOGON_USER"))
                Exit Function
            End Try
        Loop
        '*** done processing the file
        afile.Dispose()

        '*** drop the _IGNORE columns
        For c = dtI.Columns.Count - 1 To 0 Step -1
            If Regex.IsMatch(dtI.Columns(c).ColumnName, "IGNORE$") Then dtI.Columns.RemoveAt(c)
        Next
        '*** add blank ID column
        Dim dc As New DataColumn("ID", GetType(Integer))
        dc.AutoIncrement = True
        dtI.Columns.Add(dc)
        dc.SetOrdinal(0)
        dtI.Columns.Add(New DataColumn("isFullPeriod", GetType(Boolean)))

        '*** add a ssequence number ONLY NEEDED IF USING DATAADAAPTOR later
        r = 1
        For Each myR As DataRow In dtI.Rows
            myR("ID") = r
            r += 1

            If Regex.IsMatch(myR("CompanionProduct") & String.Empty, "Location", RegexOptions.IgnoreCase) Then
                '*** special case, mark the location identifiers as true
                myR("isFullPeriod") = True
            Else
                '*** work it out from the dates
                myR("isFullPeriod") = checkFullPeriod(myR("UsageDateFrom"), myR("UsageDateTo"))

            End If
        Next


        dtI.AcceptChanges()


        '*** check columncounts match
        Trace.Warn(dtI.Columns.Count)
        Trace.Warn(oDS.Tables("template").Columns.Count)
        If dtI.Columns.Count <> oDS.Tables("template").Columns.Count Then Throw New ArgumentException("column count on import does Not match tblSource")

        '*** calculate metadata for this file
        Dim myView As DataView = dtI.DefaultView
        Dim a() As String = {"BillRunID", "ChargePeriod"}
        Dim dtMeta As DataTable = myView.ToTable(True, a)

        '*** calculate record counts
        dtMeta.Columns.Add(New DataColumn("count", GetType(Integer)))

        For Each myR As DataRow In dtMeta.Rows
            myR("count") = dtI.Compute("count([ChargePeriod])", "[ChargePeriod]='" & myR("ChargePeriod") & "'")
        Next

        '*** so now we have record counts per charge period.  find the largest and this is the base table we need to work with.
        '*** we will also back fill prior months, but only going back 2 months.

        'CHECK there are no empty string names, remove any found

        For i = dtMeta.Rows.Count - 1 To 0 Step -1
            If String.IsNullOrEmpty(dtMeta.Rows(i).Item("chargeperiod")) Then dtMeta.Rows.RemoveAt(i)
        Next

        myView = dtMeta.DefaultView
        myView.Sort = "[ChargePeriod] DESC"
        'statusBar.InnerText = myView.Item(0).Item("ChargePeriod")


        Dim dConn As New OleDbConnection(sData)
        Dim restrictions(3) As String
        restrictions(3) = "TABLE"

        '*** pull list of existing tables, there may be some temp tables in there
        'https://stackoverflow.com/questions/1699897/retrieve-list-of-tables-in-ms-access-file
        dConn.Open()
        Dim dtTable As DataTable = dConn.GetSchema("tables", restrictions)
        dConn.Close()

        '*** loop for the the top three entries in dtMeta, create a table as required, or look at the target table
        '*** metadata before writing to it
        r = 0
        dConn.Open()

        'LOOP FOR EACH REQUIRED TABLE, CREATE IF NECESSARY and add a primary key

        For Each rMeta In dtMeta.Rows
            '*** does this table exist already?
            Trace.Warn("checking table exists for " & rMeta("chargePeriod"))
            Dim thisPeriod As String = rMeta("chargePeriod")

            Dim foundrows() As DataRow = dtTable.Select("TABLE_NAME='" & thisPeriod & "'")
            If foundrows.Count = 0 Then
                '*** create the table
                Dim oCmd As New OleDb.OleDbCommand(String.Concat("SELECT tblSource.* INTO ", thisPeriod, " FROM tblSource WHERE ID=0"), dConn)
                oCmd.ExecuteNonQuery()
                Trace.Warn("table created")
                '*** now set a primary key on the ID field, else commmandbuilder update will later fail
                'https://stackoverflow.com/questions/2470681/how-to-define-a-vb-net-datatable-column-as-primary-key-after-creation
                oCmd = New OleDbCommand("ALTER Table " & thisPeriod & " ADD PRIMARY KEY (ID);", dConn)
                oCmd.ExecuteNonQuery()
            End If
            '*** at this point we have either found the table, or created it, so either way continue because it exists

        Next

        Trace.Warn("table checks complete...processing data")

        '*** AGAIN LOOP thro each table, and check metadata on that table
        '*** WE ARE USING DAO recordsets not an oledb.dataadaptor.   this seems to run at about 60 records per sec
        '*** 28208 records in 500 sec, yup so about 50-60 records per sec
        '*** I think the only way to go faster is a linked table, properly datatyped and then SQL bulk insert queries

        Dim rs As DAO.Recordset
        Dim dbEngine As New DAO.DBEngine
        Dim ws As DAO.Workspace = dbEngine.Workspaces(0)
        r = 0
        Try
            'start TRANSACTION HERE
            ws.BeginTrans()
            'Dim db As DAO.Database = dbEngine.OpenDatabase("C:\Users\Julian\Documents\Visual Studio 2017\source\PCManalytics\primebillerdata.mdb", False)
            Dim m As Match = Regex.Match(sConn, "(.+)=(.+);")
            Trace.Warn(m.Groups(2).ToString)

            Dim db As DAO.Database = dbEngine.OpenDatabase((m.Groups(2).ToString), False)


            'https://docs.microsoft.com/en-us/office/vba/access/concepts/data-access-objects/use-transactions-in-a-dao-recordset
            Dim cols As Integer

            'using rs addnew and update, we can get 5000 updates per sec in native acccess.  nope, it runs at about 50 per sec
            'interop services also seem slow

            For Each rMeta In dtMeta.Rows
                Dim thisPeriod As String = rMeta("chargePeriod")

                '*** now check the metadata on that table.  The count of records for this [Bill run ID] must be less than the count
                '*** in Meta data, else we would be double-loading
                '*** assumes Bill Run ID is a string

                Dim oCmd2 = New OleDbCommand(String.Concat("SELECT count([ID]) AS CountOfID FROM ", thisPeriod, " WHERE [BillRunID]=", CLng(rMeta("BillRunID"))), dConn)

                Dim existingCount As Integer = oCmd2.ExecuteScalar
                Trace.Warn("existing rec " & existingCount)


                '*** if existing count is less than our rMeta(count) then proceed with writing
                If existingCount < rMeta("count") Then
                    'loop to write in the records
                    Dim strSQL As String = String.Empty

                    rs = db.OpenRecordset(thisPeriod)
                    cols = rs.Fields.Count
                    For Each myR As DataRow In dtI.Rows
                        If myR("ChargePeriod").ToString = thisPeriod Then
                            r += 1
                            '2018-11-05 approach, use rs.addnew and rs.update
                            rs.AddNew()
                            For c = 1 To cols - 1
                                'ignore the ID column in  posn 0
                                rs.Fields(c).Value = myR(c)
                            Next
                            rs.Update()
                        End If
                    Next
                    rs.Close()
                    Trace.Warn("added the records " & r)

                Else
                    Trace.Warn("records already exist for " & thisPeriod)
                End If

                Trace.Warn("table done")

            Next

            gvDebug.DataSource = dtMeta
            gvDebug.DataBind()

            'last action is to commit all changes across multiple tables
            ws.CommitTrans()
            Trace.Warn("committed the transaction. END")
            dConn.Dispose()
            Return String.Concat("OK, records=", r, " in ", DateDiff(DateInterval.Second, dtStart, Now), " seconds")

        Catch ex2 As System.Runtime.InteropServices.COMException
            '**** will be a recordset error of some type
            Trace.Warn(ex2.Message)
            If Not rs Is Nothing Then Trace.Warn("FIELD=" & rs.Fields(c).Name & " @ CSV row " & r)
            ws.Rollback()
            Return ("FAILED: transactions rolled back. " & ex2.Message)
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            ws.Rollback()
            Return ("FAILED: transactions rolled back")

        Finally
            dbEngine = Nothing

        End Try

        gvDebug.DataSource = dtMeta
        gvDebug.DataBind()


        Return 0



    End Function

    Function coerceType(ByVal o As Object, ByVal t As Type, maxLen As Object) As Object
        '**** coerce value to the supplied type, or return dbnull.value
        Try

            Select Case t.ToString
                Case "System.String"
                    '*** if we co-erce a string its because we want to return null instead of string.empty
                    '*** truncate at 255 chars
                    ' Trace.Warn(CLng(maxLen))
                    ' Return strTruncate(o.ToString, 255)  'else too slow!!
                    Return strTruncate(o.ToString, CLng(maxLen))
                   ' Return (o.ToString.Substring(0, 255))
                   'SUBSTRING is very slow?

                Case "System.Double"
                    If Not IsNumeric(o) Then Return DBNull.Value
                    Return CDbl(o)

                Case "System.Int32"
                    If Not IsNumeric(o) Then Return DBNull.Value
                    Return CLng(o)

                Case "System.Integer"
                    If Not IsNumeric(o) Then Return DBNull.Value
                    Return CInt(o)

                Case "System.DateTime"
                    '*** assume we have received as dd-mm-yyyy, we need to return as a date type
                    If String.IsNullOrEmpty(o) Then Return DBNull.Value
                    'convert to yyyy-mm-dd per ISO 6-4,3-2,0-2
                    Dim s As String = o.ToString.Substring(6, 4)
                    s &= "-"
                    s &= o.ToString.Substring(3, 2)
                    s &= "-"
                    s &= o.ToString.Substring(0, 2)
                    Return CDate(s)

                Case "System.Boolean"
                    Return Regex.IsMatch(o.ToString, "true|yes|1", RegexOptions.IgnoreCase)
                Case Else
                    Trace.Warn("coerce unsupported " & t.ToString)
                    Return DBNull.Value
            End Select

        Catch ex As Exception
            Return DBNull.Value
        End Try


    End Function
    Function strTruncate(s As String, ByVal maxLen As Integer) As String
        If maxLen <= 0 Then Return s

        If s.Length < maxLen Then Return s
        Return s.Substring(0, maxLen)

    End Function

    Public Shared Function getKeyNames(tableName As [String], conn As OleDbConnection) As List(Of String)
        'https://stackoverflow.com/questions/5065086/vb-net-how-can-i-check-if-a-primary-key-exists-in-an-access-db
        'https://support.microsoft.com/en-au/help/309488/how-to-retrieve-schema-information-by-using-getoledbschematable-and-vi
        Dim returnList = New List(Of String)()


        Dim mySchema As DataTable = TryCast(conn, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Primary_Keys, New [Object]() {Nothing, Nothing, tableName})


        ' following is a lengthy form of the number '3' :-)
        Dim columnOrdinalForName As Integer = mySchema.Columns("COLUMN_NAME").Ordinal

        For Each r As DataRow In mySchema.Rows
            returnList.Add(r.ItemArray(columnOrdinalForName).ToString())
        Next

        Return returnList
    End Function
    Public Shared Function getColLengths(tableName As [String], conn As OleDbConnection) As DataTable
        Dim returnList = New List(Of String)()
        Dim mySchema As DataTable = TryCast(conn, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New [Object]() {Nothing, Nothing, tableName})
        Return mySchema

        'CHARACTER_MAXIMUM_LENGTH for COLUMN_NAME
        'https://support.microsoft.com/en-au/help/309488/how-to-retrieve-schema-information-by-using-getoledbschematable-and-vi
    End Function



    Function checkFullPeriod(sFrom As Object, sTo As Object) As Boolean
        Try
            '*** checks the date markers are a full period.  special case for 
            If String.IsNullOrEmpty(sFrom) Or String.IsNullOrEmpty(sTo) Then Return False
            Dim dt As Date = CDate(sFrom)
            '*** add a month, which we expect will be the first of the next
            dt = dt.AddMonths(1)
            '*** we expect the day to the first 
            If dt.Day <> 1 Then Return False
            '*** fail if the end date differs from the expected first of the following month by more than a day
            If Math.Abs(DateDiff(DateInterval.Day, dt, CDate(sTo))) > 1 Then Return False
            '*** all tests passed, return true
            Return True
        Catch
            Return False
        End Try



    End Function

#Region "...NOT USED..."
    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'after hours of messing around, here's the easy way to set the primary key
        Dim oConn As New OleDbConnection(sData)
        Dim ocmd As New OleDbCommand("ALTER Table 201810 ADD PRIMARY KEY (ID);", oConn)
        oConn.Open()
        ocmd.ExecuteNonQuery()
        oConn.Close()

        'but its not magic because the dti table schmea is not identical to the target table
        'hence row import is giving me grief

    End Sub

    Sub copyArray(ByVal source(), ByRef target())
        '*** copy array but not first element as this is ID
        For i = 1 To target.Count - 1
            target(i) = source(i)
        Next

    End Sub


    Protected Sub bTest_Click(sender As Object, e As EventArgs) Handles bTest.Click
        'pull the database table list
        Dim oConn As New OleDbConnection(sData)
        Dim oDS As New DataSet
        'Dim oDA As New OleDbDataAdapter("SELECT * FROM tables", oConn)
        'oDA.Fill(oDS)
        Dim restrictions(3) As String
        restrictions(3) = "TABLE"

        oConn.Open()
        'gvDebug.DataSource = oConn.GetSchema("tables", restrictions)
        'gvDebug.DataBind()


        gvDebug.DataSource = getKeyNames("201810", oConn)
        gvDebug.DataBind()
        oConn.Dispose()
    End Sub

    Function buildSQLinsert(dcol As DataColumnCollection, dr As DataRow, target As String) As String
        'build an INSERT VALUES query.  this assumes all cols are present and in correct order.
        Dim sb As New StringBuilder
        sb.Append("INSERT INTO ")
        sb.Append(target)
        sb.Append(" VALUES (")

        'loop through data.  string values need double quotes
        'which in turn is a pain because we need to check datatype
        For Each dc As DataColumn In dcol
            Trace.Warn(dc.ColumnName & " " & (dr(dc.ColumnName)))
            If dc.DataType.ToString = "System.String" Then
                sb.Append(Chr(34))
                'sb.Append("'")
                sb.Append(dr(dc.ColumnName))
                sb.Append(Chr(34))
                sb.Append(",")
            ElseIf (dr(dc.ColumnName)) Is DBNull.Value Then
                sb.Append("NULL")
                sb.Append(",")
            Else
                sb.Append(dr(dc.ColumnName))
                sb.Append(",")
            End If
        Next
        'drop trailing comma
        sb.Remove(sb.Length - 1, 1)
        sb.Append(");")
        Return Regex.Replace(sb.ToString, " 12:00:00 AM", String.Empty)



    End Function

#End Region
End Class