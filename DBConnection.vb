Imports System.Data.OleDb

''' <summary>
''' Stores information used by the db connection
''' </summary>
Public Module DBData
	Public connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "/Resources/Database.accdb;"
	Public CachedDbs As Dictionary(Of String, DBConnection) = New Dictionary(Of String, DBConnection)
End Module

''' <summary>
''' This is a custom Oledb database handling class that simplifies a lot of the process
''' and makes it a lot less repetitive allowing me to focus on more key features
''' it should also *hopefully* increase performance with it's in-built db caching
''' </summary>
Public Class DBConnection
	Public Property database As DataTable
	Public Property adapter As OleDbDataAdapter
	Public Property maximumRows As Integer
	Public Property dataSet As DataSet
	Public Property cmdBuilder As OleDbCommandBuilder
	Public Property connection As New OleDbConnection

	Private Property DBName As String
	Private Property Cache As Dictionary(Of String, DBConnection) = DBData.CachedDbs
	Private Property connectionString As String = DBData.connectionString

	''' <summary>
	''' Creates a new database connection to the specific database table
	''' </summary>
	''' <param name="dbName">The table to retrieve information from the database in resources</param>
	''' <param name="Key">The primary key of the table</param>
	Public Sub New(dbName As String, Optional Key As String = Nothing)
		' Setting "new" object to cache if it exists
		If Me.Cache.ContainsKey(dbName) Then
			Dim cachedObject As DBConnection = Me.Cache(dbName)

			' Making sure the cached object isn't self
			If Not cachedObject.Equals(Me) Then
				' Setting all internal properties
				Me.database = cachedObject.database
				Me.adapter = cachedObject.adapter
				Me.maximumRows = cachedObject.maximumRows
				Me.dataSet = cachedObject.dataSet
				Me.cmdBuilder = cachedObject.cmdBuilder
				Me.connection = cachedObject.connection
				Me.DBName = cachedObject.DBName

				Console.WriteLine("Returned Cached Database: " & dbName)
				Return
			End If
		End If

		Console.WriteLine("Create new DB Connection: " & dbName)

		' Creating the connection to the database
		Me.DBName = dbName
		Me.connection.ConnectionString = Me.connectionString
		Me.connection.Open()

		' Using an adapter to communicate with the database.
		Me.dataSet = New DataSet
		Me.adapter = New OleDbDataAdapter("SELECT * FROM [" & dbName & "]", connection)

		' Filling the dataset with entries from the database
		Me.adapter.Fill(Me.dataSet, dbName)

		' Setting the primary key
		If Not IsNothing(Key) Then
			Me.dataSet.Tables(dbName).PrimaryKey = New DataColumn() {Me.dataSet.Tables(dbName).Columns(Key)}
		End If
		Me.database = Me.dataSet.Tables(dbName)

		' Setting up commands
		Me.cmdBuilder = New OleDbCommandBuilder(Me.adapter)
		Me.adapter.InsertCommand = Me.cmdBuilder.GetInsertCommand
		If Not IsNothing(Key) Then
			Me.adapter.DeleteCommand = Me.cmdBuilder.GetDeleteCommand
			Me.adapter.UpdateCommand = Me.cmdBuilder.GetUpdateCommand
		End If

		' Setting the self variables
		Me.maximumRows = Me.database.Rows.Count

		' Closing the connection because it is no longer needed
		Me.connection.Close()

		' Creating a cache entry so this db can be used multiple times
		Me.Cache.Add(dbName, Me)
	End Sub

	''' <summary>
	''' Sets the DeleteCommand for the SQL Database
	''' Updates the tables, and the builder to support the new command.
	''' </summary>
	''' <param name="SQL">The new SQL Command for deletion</param>
	Public Sub DeleteCommand(SQL As String)
		' Running the cached objects function instead
		If Me.Cache.ContainsKey(DBName) Then
			Dim cachedObject As DBConnection = Me.Cache(DBName)

			' Making sure the cached object isn't self
			Console.WriteLine("Use Cached Delete Command (outside)")
			If Not cachedObject.Equals(Me) Then
				Console.WriteLine("Use Cached Delete Command (inside)")
				cachedObject.DeleteCommand(SQL)
			End If
		End If

		' Updating the command
		Me.adapter.DeleteCommand = New OleDbCommand(SQL)
		Me.cmdBuilder.RefreshSchema()

		' Resetting the tables
		Me.dataSet.Tables.Remove(Me.dataSet.Tables(Me.DBName))
		Me.adapter.Fill(Me.dataSet, Me.DBName)
	End Sub

	''' <summary>
	''' Sends the data from the local table to the database
	''' </summary>
	Public Sub Update()
		' Running the cached objects function instead
		If Me.Cache.ContainsKey(DBName) Then
			Dim cachedObject As DBConnection = Me.Cache(DBName)

			' Making sure the cached object isn't self
			Console.WriteLine("Use Cached Update Function (inside)")
			If Not cachedObject.Equals(Me) Then
				Console.WriteLine("Use Cached Update Function (inside)")
				cachedObject.Update()
			End If
		End If

		' Opening the connection and updating
		Me.connection.Open()
		Try
			Me.adapter.Update(Me.dataSet, DBName)
		Catch ex As Exception
			Console.WriteLine("Failed to update database :\n" & ex.Message)
		End Try

		' Updating internal properties
		Me.maximumRows = Me.database.Rows.Count

		' Finishing up
		Me.database.AcceptChanges()
		Me.connection.Close()
	End Sub

End Class