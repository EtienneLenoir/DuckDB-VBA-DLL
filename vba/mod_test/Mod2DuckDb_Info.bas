Attribute VB_Name = "Mod2DuckDb_Info"
Option Explicit

'===============================================================================
'                    Module :  INFO / ADMIN DuckDB
'
' - Test_VersionDuckDbDll : affiche version() DuckDB (check DLL + QueryFast).
' - Test_CheckBaseDuckDb  : liste tables/vues via TablesInfo (check metadata tables).
' - Test_CheckTable       : liste colonnes via ColumnsInfo (check metadata colonnes).
' - Test_TableExists      : vérifie existence table/vue (check TableExists).
' - Test_ColumnExists     : vérifie existence colonne (check ColumnExists).
' - Test_RenameTableColumn: renomme table + colonne (check DDL Rename*).
' - Demo_PRAGMA_CheatSheet: démo PRAGMA/SET (check Exec + diagnostics).
' - Demo_ParquetInfo      : infos parquet sur feuille (check extension parquet).
' - Demo_CountParquetRows : COUNT(*) sur parquet (check scan parquet)
'
'===============================================================================

'affiche version() DuckDB (check DLL + QueryFast).
Public Sub Test_VersionDuckDbDll()

    On Error GoTo Fail
    Dim db As New cDuck, a As Variant
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    a = db.QueryFast("SELECT 'duckdb_version' AS k UNION ALL SELECT version()")
    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    End With
    db.CloseDuckDb
    MsgBox "OK QueryToArrayV"
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

'liste colonnes via ColumnsInfo (check metadata colonnes).
Public Sub Test_CheckTable()

    On Error GoTo Fail

    'Adapter ces deux valeurs si besoin :
    Dim db As New cDuck, v As Variant, dbPath As String, tablePath As String
    dbPath = ThisWorkbook.Path & "\cache.duckdb"
    tablePath = "main.Instruments"   ' ex: "main.MaTable"
    '1) Session DuckDB via ta classe
    db.Init ThisWorkbook.Path
    db.ErrorMode = 1 ' 0=Raise, 1=MsgBox, 2=LogOnly
    db.OpenDuckDb dbPath
    '2) Appel direct à l’API ColumnsInfoV (comme ton code d’origine)
     v = db.ColumnsInfo(tablePath)
    '3) Dépôt sur feuille
    DumpVariant2D v, "DuckDB_Columns", "A1"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit
       
End Sub

'liste tables/vues via TablesInfo (check metadata tables).
Public Sub Test_CheckBaseDuckDb()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, schemaFilter As String
    
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2           ' 2=log only, 1=msgbox, 0=raise
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"   ' adapte si besoin

    '1) Tous schémas (hors systèmes) -> passer NULL = 0
    
    schemaFilter = ""
    v = db.TablesInfo(schemaFilter)
    DumpVariant2D v, "DuckDB_Tables", "A1"
    
    'If DuckVba_TableInfoV(db.handle, 0, v) = 0 Then
        'Err.Raise 5, , "TableInfo(all) KO: " & Native_LastErrorText()
    'Else
        'DumpVariant2D v, "DuckDB_Tables", "A1"
   ' End If

    '2) Filtrer sur un schéma (ex: "main")
    'If DuckVba_TableInfoV(db.Handle, StrPtr("main"), V) = 0 Then
    '    Err.Raise 5, , "TableInfo(main) KO: " & Native_LastErrorText()
    'Else
    '    DumpVariant2D V, "DuckDB_Tables", "G1"
    'End If
    MsgBox "OK: TableInfo rempli (onglet 'DuckDB_Tables').", vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit
    
End Sub

'vérifie existence table/vue (check TableExists).
Sub Test_TableExists()

    Dim db As New cDuck
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\test.duckdb"   ' adapte le chemin
    Debug.Print "main.clients        ? "; db.TableExists("main.clients")
    Debug.Print "main.table_quiexistepas ? "; db.TableExists("main.table_quiexistepas")
    db.CloseDuckDb

End Sub

'vérifie existence colonne (check ColumnExists).
Sub Test_ColumnExists()

    Dim db As New cDuck
    On Error GoTo Fail
    '1) DuckDB en mémoire
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    '2) Crée une petite table de test
    db.Exec "CREATE TABLE main.Clients (" & _
            "   Id      INTEGER," & _
            "   Nom     TEXT," & _
            "   DateCrea DATE" & _
            ");"
    '3) Tests TableExistsDll
    Debug.Print "Table main.Clients existe ? ", db.TableExists("main.Clients")
    Debug.Print "Table main.Inconnue existe ? ", db.TableExists("main.Inconnue")
    '4) Tests ColumnExists
    Debug.Print "Colonne Id existe ?       ", db.ColumnExists("main.Clients", "Id")
    Debug.Print "Colonne Nom existe ?      ", db.ColumnExists("main.Clients", "Nom")
    Debug.Print "Colonne DateCrea existe ? ", db.ColumnExists("main.Clients", "DateCrea")
    Debug.Print "Colonne Bidon existe ?    ", db.ColumnExists("main.Clients", "Bidon")
Bye:
    db.CloseDuckDb
    Exit Sub
Fail:
    MsgBox "Erreur dans Test_Table_ColumnExists_Simple : " & Err.Description, vbExclamation
    Resume Bye

End Sub

'renomme table + colonne (check DDL Rename*).
Sub Test_RenameTableColumn()

    On Error GoTo Fail

    Dim db As New cDuck

    '1) DuckDB en mémoire
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    '2) Table de départ : main.Clients (Id, Nom)
    db.Exec "CREATE TABLE main.Clients (" & _
            "   Id  INTEGER," & _
            "   Nom TEXT" & _
            ");"

    Debug.Print "Clients existe ?           ", db.TableExists("main.Clients")
    Debug.Print "Colonne Nom existe ?      ", db.ColumnExists("main.Clients", "Nom")

    '3) Renomme la table Clients -> Customers
    db.RenameTable "main.Clients", "Customers"

    Debug.Print "Clients existe ? (après)  ", db.TableExists("main.Clients")
    Debug.Print "Customers existe ?        ", db.TableExists("main.Customers")

    '4) Renomme la colonne Nom -> Name
    db.RenameColumn "main.Customers", "Nom", "Name"

    Debug.Print "Colonne Nom existe ?      ", db.ColumnExists("main.Customers", "Nom")
    Debug.Print "Colonne Name existe ?     ", db.ColumnExists("main.Customers", "Name")
Bye:
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur dans Test_Rename_Table_Column_Simple : " & Err.Description, vbExclamation
    Resume Bye

End Sub

'démo PRAGMA/SET (check Exec + diagnostics).
Public Sub Demo_PRAGMA_CheatSheet()

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, v As Variant, y As Long
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"   'ou ":memory:"

    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear
    y = 1

    'Titre
    ws.Cells(y, 1).Value = "PRAGMA / SET — démo"
    ws.Cells(y, 1).Font.Bold = True
    y = y + 2

    '1)PRAGMA sans résultat (équivalent SET)
    db.Exec "PRAGMA threads=4;"
    v = db.QueryFast("SELECT current_setting('threads') AS threads;")
    ws.Cells(y, 1).Value = "1) Threads courants": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    '2a) PRAGMA database_list
    v = db.QueryFast("PRAGMA database_list;")
    ws.Cells(y, 1).Value = "2a) PRAGMA database_list": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    '2b) PRAGMA show_tables
    v = db.QueryFast("PRAGMA show_tables;")
    ws.Cells(y, 1).Value = "2b) PRAGMA show_tables": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    '2c) Table de démo + table_info
    db.Exec "CREATE TABLE IF NOT EXISTS DemoT(a INT, s TEXT);"
    db.Exec "DELETE FROM DemoT; INSERT INTO DemoT VALUES (1,'x'),(2,'y'),(3,'z');"

    v = db.QueryFast("PRAGMA table_info(" & SqlQ("DemoT") & ");")
    ws.Cells(y, 1).Value = "2c) PRAGMA table_info('DemoT')": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    '2d) storage_info
    v = db.QueryFast("PRAGMA storage_info(" & SqlQ("DemoT") & ");")
    ws.Cells(y, 1).Value = "2d) PRAGMA storage_info('DemoT')": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    '2e) Taille de la base
    v = db.QueryFast("PRAGMA database_size;")
    ws.Cells(y, 1).Value = "2e) PRAGMA database_size": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    '3) Variantes CALL pragma_*
    v = db.QueryFast("CALL pragma_database_size();")
    ws.Cells(y, 1).Value = "3a) CALL pragma_database_size()": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    v = db.QueryFast("CALL pragma_table_info('DemoT');")
    ws.Cells(y, 1).Value = "3b) CALL pragma_table_info('DemoT')": ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    MsgBox "Démo PRAGMA terminée – résultats déposés sur la feuille 1.", vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

'infos parquet sur feuille (check extension parquet).
Public Sub Demo_ParquetInfo()

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, v As Variant, parquetPath As String

    '1) Init + ouverture en mémoire
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    '2) Chemin du fichier (ou motif)
    'parquetPath = ThisWorkbook.Path & "\instruments.parquet"
    parquetPath = ThisWorkbook.Path & "\access_table.parquet"

    '3) Récupération des infos
    v = Duck_ParquetInfo(db, parquetPath)

    '4) Dépôt sur la feuille
    Set ws = ThisWorkbook.Worksheets(1)
    ArrayToSheet v, ws, "A1"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur ParquetInfo: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit
    
End Sub

'COUNT(*) sur parquet (check scan parquet)
Public Sub Demo_CountParquetRows()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, parquetPath As String, sql As String

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    db.TryLoadExt "parquet"
    '--- 2) Fichier cible ---
    parquetPath = ThisWorkbook.Path & "\access_table.parquet"
    parquetPath = Replace(parquetPath, "\", "/")
    '--- 3) SQL simple ---
    sql = "SELECT COUNT(*) AS row_count FROM read_parquet(" & SqlQ(parquetPath) & ");"
    '--- 4) Exécution et affichage ---
    v = db.QueryFast(sql)
    MsgBox "Nombre de lignes dans le parquet : " & v(2, 1), vbInformation, "DuckDB COUNT(*)"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur COUNT Parquet: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit
    
End Sub
