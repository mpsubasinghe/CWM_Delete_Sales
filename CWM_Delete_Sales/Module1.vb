Imports System.Data.SqlClient
Imports System.Data.OleDb

Imports System.Text.RegularExpressions


Module Module1

    Public DownLoad_R_Code As String = ""
    Public DownLoad_RmCod_Code As String = ""

    Public Colecction_S_Code As String = ""
    Public Colecction_Retailer As String = ""

    Public Areacode As String = ""
    Public AreaName As String = ""
    Public Head As String = ""
    Public RegMan As String = ""
    Public SupsCode As String = ""
    Public Region1 As String = ""
    Public Stockist1 As String = ""


    Public comcode As String = ""
    Public PDAInv As String = ""
    Public Sectors As String = ""
    Public Route As String = ""
    Public RetailerName As String = ""
    Public RepName As String = ""
    Public StkName As String = ""
    Public ItemName As String = ""
    Public SEQ As String = ""


    Public D_DailySaleID As Integer = 0
    Public D_DailySaleDate As String = ""



    Public CATE As String = ""

    Public conn As New SqlConnection

    ' Public AS400Str As String = "Provider=IBMDA400;Data Source=192.168.190.2;User Id=" & Trim(Login.UIDtxt.Text) & ";Password=" & Trim(Login.PWDtxt.Text) & ";"
    'Public AS400Str As String = ""
    'Public InvNo As String
    'Public Dattime As String
    'Public prg As Integer
    'Public mon As String
    'Public SecID As String
    'Public Route As String
    'Public online As Boolean'
    Public directory As String = My.Application.Info.DirectoryPath

    ' Public aceesscon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.4.0;Data Source=" & directory & "\database\DBCMarketing.mdb;Jet OLEDB:Database Password=redrock;")

    Public aceesscon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & directory & "\database\SFA.mdb;Jet OLEDB:Database Password=dbclmis321;")

    'Public aceesscon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & directory & "\database\DBCMarketing1.mdb;Jet OLEDB:System Database=system.mdw;User ID=admin;Password=redrock;")
    'Public aceesscon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & directory & "\database\SFA.mdb;Jet OLEDB:Database Password=dbclmis321;")


    ' Public cona As New OleDbConnection("Provider=IBMDA400;Data Source=192.168.190.2;User Id=CALS;Password=CALS123;")
    'Public cona As New OleDbConnection(AS400Str)
    '  Public con As New SqlConnection("Data Source=SQL5021.Smarterasp.net;Initial Catalog=DB_9AAA2F_PhamaNew;User Id=DB_9AAA2F_PhamaNew_admin;Password=mano1234;")

    ' Public sqlite As New SQLiteConnection("Data Source=" & directory & "\database\Android\SFAD.sqlite;Version=3;")
    '  Public con As New SqlConnection("Data Source=192.168.0.43,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeDBC;User ID=admin;Password=a;")


    ' Public SFAHeadoffice_SMARTER As New SqlConnection("Data Source=SQL5016.Smarterasp.net;Initial Catalog=DB_9AAA2F_SFAHeadoffice;User Id=DB_9AAA2F_SFAHeadoffice_admin;Password=dbcl1234;")
    Public SFAHeadoffice As New SqlConnection("Data Source=SQL5016.Smarterasp.net;Initial Catalog=DB_9AAA2F_SFAHeadoffice;User Id=DB_9AAA2F_SFAHeadoffice_admin;Password=dbcl1234;")
    'Public SFAHeadoffice As New SqlConnection("Data Source=SQL5016.Smarterasp.net;Initial Catalog=DB_9AAA2F_SFAHeadoffice;User Id=DB_9AAA2F_SFAHeadoffice_admin;Password=dbcl1234;")
    'Public SFAHeadoffice As New SqlConnection("Data Source=SQL5016.Smarterasp.net;Initial Catalog=DB_9AAA2F_SFAHeadoffice;User Id=DB_9AAA2F_SFAHeadoffice_admin;Password=dbcl1234;")
    'Public SFAHeadoffice As New SqlConnection("Data Source=SQL5016.Smarterasp.net;Initial Catalog=DB_9AAA2F_SFAHeadoffice;User Id=DB_9AAA2F_SFAHeadoffice_admin;Password=dbcl1234;")
    Public SFAHeadofficeLoal As New SqlConnection("Data Source=10.1.6.36,1433;Network Library=DBMSSOCN;Initial Catalog=SFAHeadofficeCWM;User ID=sa;Password=admin@Sfa99;")
    ' Public con As New SqlConnection("Data Source=localhost,1433;Initial Catalog=SFAHeadofficeDBC;User ID=sa;Password=mano1234;")

    Function GetDataSQL(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String, ByVal sqlcon As SqlConnection) As System.Data.DataSet
        Try

            ' con.Open()
            Dim queryString As String = sql
            Dim dbCommand As System.Data.IDbCommand = New System.Data.SqlClient.SqlCommand
            dbCommand.CommandText = queryString
            dbCommand.Connection = sqlcon

            Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p1.ParameterName = pn1
            dbParam_p1.Value = p1
            dbParam_p1.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p1)

            Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p2.ParameterName = pn2
            dbParam_p2.Value = p2
            dbParam_p2.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p2)

            Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p3.ParameterName = pn3
            dbParam_p3.Value = p3
            dbParam_p3.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p3)

            Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            dataAdapter.SelectCommand = dbCommand
            Dim dataSet As System.Data.DataSet = New System.Data.DataSet
            dataAdapter.Fill(dataSet)

            'con.Close()
            Return dataSet

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            '  con.Close()
        End Try

    End Function


    Function GetDataACC(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String) As System.Data.DataSet
        ' MsgBox(sql)
        Dim queryString As String = sql
        Dim dbCommand As System.Data.IDbCommand = New System.Data.OleDb.OleDbCommand
        dbCommand.CommandText = queryString
        dbCommand.Connection = aceesscon

        Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p1.ParameterName = pn1
        dbParam_p1.Value = p1
        dbParam_p1.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p1)

        Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p2.ParameterName = pn2
        dbParam_p2.Value = p2
        dbParam_p2.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p2)

        Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p3.ParameterName = pn3
        dbParam_p3.Value = p3
        dbParam_p3.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p3)

        Dim dbParam_p4 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p4.ParameterName = pn4
        dbParam_p4.Value = p4
        dbParam_p4.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p4)

        Dim dbParam_p5 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p5.ParameterName = pn5
        dbParam_p5.Value = p5
        dbParam_p5.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p5)

        Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter
        dataAdapter.SelectCommand = dbCommand
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet
        dataAdapter.Fill(dataSet)


        Return dataSet

        dataSet.Dispose()

    End Function


    'Function GetSQLITE(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String) As System.Data.DataSet
    '    ' MsgBox(sql)
    '    Dim queryString As String = sql
    '    Dim dbCommand As System.Data.IDbCommand = New System.Data.SQLite.SQLiteCommand
    '    dbCommand.CommandText = queryString
    '    dbCommand.Connection = sqlite

    '    Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.SQLite.SQLiteParameter
    '    dbParam_p1.ParameterName = pn1
    '    dbParam_p1.Value = p1
    '    dbParam_p1.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p1)

    '    Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.SQLite.SQLiteParameter
    '    dbParam_p2.ParameterName = pn2
    '    dbParam_p2.Value = p2
    '    dbParam_p2.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p2)

    '    Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.SQLite.SQLiteParameter
    '    dbParam_p3.ParameterName = pn3
    '    dbParam_p3.Value = p3
    '    dbParam_p3.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p3)

    '    Dim dbParam_p4 As System.Data.IDataParameter = New System.Data.SQLite.SQLiteParameter
    '    dbParam_p4.ParameterName = pn4
    '    dbParam_p4.Value = p4
    '    dbParam_p4.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p4)

    '    Dim dbParam_p5 As System.Data.IDataParameter = New System.Data.SQLite.SQLiteParameter
    '    dbParam_p5.ParameterName = pn5
    '    dbParam_p5.Value = p5
    '    dbParam_p5.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p5)

    '    Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.SQLite.SQLiteDataAdapter
    '    dataAdapter.SelectCommand = dbCommand
    '    Dim dataSet As System.Data.DataSet = New System.Data.DataSet
    '    dataAdapter.Fill(dataSet)


    '    Return dataSet

    '    dataSet.Dispose()

    'End Function


    Public Function SqlComand(con1 As SqlConnection, SQL As String) As DataSet
        Dim ds As New DataSet()
        Try
            ' SqlConnection conn = new SqlConnection(ConnectionString);
            Dim da As New SqlDataAdapter()
            Dim cmd As SqlCommand = con1.CreateCommand()
            cmd.CommandText = SQL
            da.SelectCommand = cmd


            da.Fill(ds)
        Catch e As Exception
            Console.WriteLine(e.Message)
        Finally


        End Try

        Return ds


    End Function



    Public Function isNumeric(input As String) As Boolean
        Return Regex.IsMatch(input.Trim, "\A-{0,1}[0-9.]*\Z")
    End Function
    Function ComboSelectedValue(ByVal sql As String, ByVal select_val As String)
        Dim ID As String = ""
        Dim ds As DataSet
        ds = GetDataACC("", "p2", "p3", "p4", "p5", "@id", "pn2", "pn3", "pn4", "pn5", sql)
        '    MsgBox(ds.Tables(0).Rows.Count)
        If ds.Tables(0).Rows.Count > 0 Then
            For a = 0 To ds.Tables(0).Rows.Count - 1
                ID = (ds.Tables(0).Rows(a).Item(select_val))
            Next
        End If

        Return ID

    End Function
    Function LoadCombo(ByVal ds As DataSet, ByVal combo As ComboBox, ByVal p As String)
        combo.Items.Clear()
        If ds.Tables(0).Rows.Count > 0 Then
            For a = 0 To ds.Tables(0).Rows.Count - 1
                combo.Items.Add(ds.Tables(0).Rows(a).Item(p))
            Next
        End If
    End Function
    'Function InsertDataACC(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String) As System.Data.DataSet
    '     MsgBox(sql)
    '    Dim queryString As String = sql
    '    Dim dbCommand As System.Data.IDbCommand = New System.Data.OleDb.OleDbCommand
    '    dbCommand.CommandText = queryString
    '    dbCommand.Connection = aceesscon

    '    Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
    '    dbParam_p1.ParameterName = pn1
    '    dbParam_p1.Value = p1
    '    dbParam_p1.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p1)

    '    Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
    '    dbParam_p2.ParameterName = pn2
    '    dbParam_p2.Value = p2
    '    dbParam_p2.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p2)

    '    Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
    '    dbParam_p3.ParameterName = pn3
    '    dbParam_p3.Value = p3
    '    dbParam_p3.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p3)

    '    Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter
    '    dataAdapter.SelectCommand = dbCommand
    '    Dim dataSet As System.Data.DataSet = New System.Data.DataSet
    '    dataAdapter.Fill(dataSet)

    '       Return dataSet

    'End Function

End Module

