# DBConnection
A class for handling OleDB databases (for college assignment)


# Usage

## Creating a DBConnection:
### 1) Database with a primary key (Recommended and preferred)
```vb
Dim MyNewDatabase As New DBConnection("TableName", "PrimaryKey")
```

### 2) Database without a primary key (Unrecommended)
```vb
Dim MyNewDatabase As New DBConnection("TableName")
```
*The reason I recommend you use a primary key is simply because a primary key allows you to use the .Find function*

## Database Rows

### 1) **Creating** a new row
```vb
Dim NewRow As DataRow = MyNewDatabase.database.NewRow() ' Creates an empty DataRow

NewRow.Item("Key") = "Value"
'... Add any row values you want

' Adding the row to the database
MyNewDatabase.database.Rows.Add(NewRow) ' This adds the row to the internal table
MyNewDatabase.Update() ' This will update the actual physical database
```

### 2) **Deleting** a row 
```vb
' Method A) Sets the delete command to a custom SQL string
MyNewDatabase.DeleteCommand("DELETE FROM TableName WHERE Key='Value'")
MyNewDatabase.Update() ' This will update the actual physical database

' Method B) Deleting a row
MyDataRow.Delete() ' Deleting the datarow object
MyNewDatabase.Update() ' Updating the database to reflect deleted row change
```

### 3) **Reading** a row
There are *two* methods you can use.

Method One (using primary key and .Find()):
```vb
' .Find lets you search for a row with the primary key you pass
Dim MyDataRow As DataRow = MyNewDatabase.database.Rows.Find("PrimaryKeyValue")
Console.WriteLine(MyDataRow.Item("Key")) ' Outputting the value of Column "Key"
```

Method Two (using indexes):
```vb
Dim MyDataRow As DataRow = MyNewDatabase.database.Rows(1)
Console.WriteLine(MyDataRow.Item("Key")) ' Outputting the value of Column "Key"
```
