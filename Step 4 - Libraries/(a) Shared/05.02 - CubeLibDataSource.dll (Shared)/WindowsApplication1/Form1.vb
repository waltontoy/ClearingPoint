
Imports CubeLibDataSource

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Debug.Print(CDatasource.SadbelTableType.BOX_DEFAULT_IMPORT_ADMIN.ToString())

        CDatasource.SadbelTableType.BOX_DEFAULT_IMPORT_ADMIN.ToString().Replace("_", " ")

        Dim DataSource As CDatasource = New CDatasource()
        'Dim success As Integer
        Dim rstRecord As New ADODB.Recordset

        DataSource.SetPersistencePath("C:\Cubepoint\Unit Test Transactions")

        Debug.Print(DataSource.GetEnumFromTableName("BOX DEFAULT EXPORT ADMIN", CDatasource.DBInstanceType.DATABASE_SADBEL))

        rstRecord = DataSource.ExecuteQuery("SELECT * FROM [USERS] ", CDatasource.DBInstanceType.DATABASE_OTHER, , , "UsersDB.mdb")

        rstRecord.AddNew()
        rstRecord.Fields("User_Name").Value = "Raymondxx"
        rstRecord.Fields("User_Password").Value = "1xx"
        rstRecord.Update()

        DataSource.InsertOtherDB(rstRecord, "Users", "UsersDB")

        DataSource.BeginTransaction(CDatasource.DBInstanceType.DATABASE_OTHER, vbNullString, "UsersDB.mdb")

        rstRecord.AddNew()
        rstRecord.Fields("User_Name").Value = "xx05"
        rstRecord.Fields("User_Password").Value = "yy05"
        rstRecord.Update()

        'objRecord = New CubeLibDataSource.DNetRecordset
        'objRecord.InitializeClass(rstRecord, rstRecord.Bookmark)

        DataSource.InsertOtherDB(rstRecord, "Users", "UsersDB")

        DataSource.CommitTransaction(CDatasource.DBInstanceType.DATABASE_OTHER, vbNullString, "UsersDB.mdb")

        rstRecord.AddNew()
        rstRecord.Fields("User_Name").Value = "xx06"
        rstRecord.Fields("User_Password").Value = "yy06"
        rstRecord.Update()

        DataSource.InsertOtherDB(rstRecord, "Users", "UsersDB")

        Dim x As Long
        x = rstRecord.Fields("User_ID").Value

        DataSource.BeginTransaction(CDatasource.DBInstanceType.DATABASE_OTHER, vbNullString, "UsersDB.mdb")

        rstRecord.AddNew()
        rstRecord.Fields("User_Name").Value = "6666"
        rstRecord.Fields("User_Password").Value = "9999"
        rstRecord.Update()

        DataSource.InsertOtherDB(rstRecord, "Users", "UsersDB")
        DataSource.RollbackTransaction(CDatasource.DBInstanceType.DATABASE_OTHER, vbNullString, "UsersDB.mdb")
        rstRecord.Delete()
        rstRecord.MoveLast()

        rstRecord.Fields("User_Name").Value = "hala"
        rstRecord.Fields("User_Password").Value = "ka"
        rstRecord.Update()

        DataSource.UpdateOtherDB(rstRecord, "Users", "UsersDB.mdb")


        DataSource.RollbackTransaction(CDatasource.DBInstanceType.DATABASE_OTHER, vbNullString, "UsersDB.mdb")

        DataSource.BeginTransaction(CDatasource.DBInstanceType.DATABASE_OTHER, vbNullString, "UsersDB.mdb")

        rstRecord.AddNew()
        rstRecord.Fields("User_Name").Value = "xx07"
        rstRecord.Fields("User_Password").Value = "yy07"
        rstRecord.Update()

        DataSource.InsertOtherDB(rstRecord, "Users", "UsersDB")

        DataSource.CommitTransaction(CDatasource.DBInstanceType.DATABASE_OTHER, vbNullString, "UsersDB.mdb")

        'ADORecordsetOpen("SELECT [Lic_SerialNumber] AS [Lic_SerialNumber] FROM [Licensee] ", ADOConnection, rstLicense, adOpenKeyset, adLockOptimistic, , True)

        'success = DataSource.UpdateSadbel(objRecord, DataSource.GetEnumFromTableName("BOX DEFAULT EXPORT ADMIN", CDatasource.DBInstanceType.DATABASE_SADBEL))

        'Dim uh As New DNetRecordset
        'x.UpdateSadbel(uh, CDatasource.SadbelTableType.BOX_DEFAULT_IMPORT_ADMIN)
    End Sub
End Class
