Attribute VB_Name = "Mod1DuckDb_Begin"
Option Explicit

'===============================================================================
'                       Demo_Begin (fonctions de base)

'    DuckDB est un moteur SQL analytique embarqué (in-process), conçu pour exécuter des requêtes
'    OLAP (scan, agrégations, jointures, fenêtrage) avec une exécution vectorisée/colonnaire et un
'    optimiseur de requêtes moderne, tout en restant sans serveur (pas de service à installer, pas de port, pas de déploiement).
'    Il s’utilise soit en base éphémère en mémoire (:memory:) pour des traitements rapides et jetables,
'    soit en base persistante sur fichier (.duckdb) pour conserver tables et index.
'
' DuckDB + VBA Bridge (duckdb_vba_bridge.dll)
' ------------------------------------------
' Ce projet utilise une DLL "bridge" qui expose DuckDB (le moteur C/C++ de DuckDB) à VBA sans driver ODBC
'
' Pourquoi DuckDB est puissant ici ?
' ---------------------------------
' - Mode mémoire : db.OpenDuckDb ":memory:"
'     -> base 100% RAM, très rapide, idéale pour tests, ETL temporaire, calculs
'        intermédiaires, prototypage "python-like" depuis des arrays.
' - Mode fichier : db.OpenDuckDb "...\DbDuckDb.duckdb"
'     -> base persistante, utile pour cache, historisation,
'        data mart léger à côté d'Excel.
'
'----------Ecriture/ ReadONly------------
' - Mode ecriture :   db.OpenDuckDb "...\xxx.duckdb"
        'CREATE, INSERT, UPDATE, DELETE, DROP, CREATE INDEX, etc.
' - Mode lecture seule : db.OpenReadOnly "...\DbDuckDb.duckdb"
'        'SELECT, PRAGMA
'
' Fonctions Principales
' ---------------------------------------
' - db.Exec(sql)
'     -> exécute du SQL non-SELECT (DDL/DML).
' - db.QueryFast(selectSql) As Variant
'     -> exécute un SELECT via le code C du bridge et renvoie un Variant(2D)
'
' Gestion d’erreur (ErrorMode)
' ----------------------------
' - db.ErrorMode = 2  (LogOnly) : écrit dans duckdb_errors.log, pas de MsgBox,pas de Raise(idéal batch/ETL).
' - db.ErrorMode = 1  (MsgBox)  : affiche un MsgBox (debug interactif).
' - db.ErrorMode = 0  (Raise)   : lève une erreur VBA (mode strict).
'
' Démonstrations du module:
' -----------------------
' - Demo_CreateTable      : crée la base + table Instruments + index.
' - Demo_OpenRO           : ouvre en lecture seule et affiche une table.
' - Demo_OpenMemory       : base en RAM (test rapide sans fichiers).
' - Demo_Insert           : INSERT/UPDATE/UPSERT avec transaction.
' - Demo_DisplayPowerQuery: affichage via SelectToSheet (robuste/UTF-8).
' - Demo_DisplayFast      : affichage via QueryFast (rapide Variant(2D)).
' - Demo_ImportCsv        : import CSV (Replace) + affichage sur feuille 1.
' - Demo_Brackets         : test parsing des identifiants [entre crochets].
'
'===============================================================================

Public Sub Demo_CreateTable() 'Démo Create Table

    Dim db As New cDuck
    
    On Error GoTo Fail
    
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenDuckDb ThisWorkbook.Path & "\DbDuckDb.duckdb"

    db.Exec "CREATE TABLE IF NOT EXISTS Instruments(" & _
            "ISIN TEXT, NumeroContrat TEXT, Prix DOUBLE, ModifiedAt TIMESTAMP);"
    db.Exec "CREATE INDEX IF NOT EXISTS ix_inst_isin ON Instruments(ISIN);"
    db.Exec "CREATE INDEX IF NOT EXISTS ix_inst_num ON Instruments(NumeroContrat);"

    MsgBox "Base prête."

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
        db.Rollback
        db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit
    
End Sub

Public Sub Demo_OpenRO() 'Ouverture en lecture

    Dim db As New cDuck, ws As Worksheet, a As Variant
    
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenReadOnly ThisWorkbook.Path & "\Db_DuckDb_Exemple.duckdb" 'Init + ouverture en lecture seule
    
    a = db.QueryFast("SELECT * FROM Instruments")
    Set ws = ThisWorkbook.Worksheets(1)
    Call ArrayToSheet(a, ws, "A1")
    
    db.CloseDuckDb
    
End Sub

Public Sub Demo_OpenMemory() 'ouverture en RAM

    Dim db As New cDuck, ws As Worksheet, arr As Variant
    
    db.Init
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenDuckDb ":memory:"   ' base temporaire 100% en RAM
    ' --- Mini table en mémoire ---
    db.Exec "CREATE TABLE T([ISIN] TEXT, [Nom] TEXT, [Prix] DOUBLE, [ModifiedAt] TIMESTAMP);"
    db.Exec "INSERT INTO T VALUES " & _
            "('FR0000987654','Contrat A',101.25,NOW())," & _
            "('FR0000123456','Contrat B',99.80,NOW());"

    Set ws = ThisWorkbook.Worksheets(1)
    arr = db.QueryFast("SELECT * FROM  T")
    Call ArrayToSheet(arr, ws, "A1")

    db.CloseDuckDb
    MsgBox "OK : test en mémoire terminé", vbInformation
End Sub


Public Sub Demo_Insert() 'Demo Insert

    Dim db As New cDuck

    On Error GoTo Fail
    
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenDuckDb ThisWorkbook.Path & "\DbDuckDb.duckdb"

    db.BeginTx
    db.Exec "INSERT INTO Instruments VALUES ('FR0000123460','C-001',101.25,NOW());"
    db.Exec "UPDATE Instruments SET Prix=103.10, ModifiedAt=NOW() WHERE ISIN='FR0000123456';"
    db.Exec "CREATE TABLE IF NOT EXISTS Quotes(ISIN TEXT PRIMARY KEY, Prix DOUBLE, ModifiedAt TIMESTAMP);"
    db.Exec "INSERT INTO Quotes VALUES ('FR0000123456', 99.9, NOW()) " & _
            "ON CONFLICT(ISIN) DO UPDATE SET Prix=excluded.Prix, ModifiedAt=excluded.ModifiedAt;"
    db.Commit

    MsgBox "CRUD/UPSERT OK"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume CleanExit
    
End Sub

Public Sub Demo_DisplayPowerQuery() 'Demo Display avec Méthode PowerQuery

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet
    
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenReadOnly ThisWorkbook.Path & "\DbDuckDb.duckdb"

    Set ws = ThisWorkbook.Worksheets(1)

    db.SelectToSheet _
        "SELECT ISIN, NumeroContrat, Prix, " & _
        "       strftime(ModifiedAt, '%Y-%m-%d %H:%M:%S') AS ModifiedAt " & _
        "FROM Instruments " & _
        "ORDER BY ModifiedAt DESC " & _
        "LIMIT 1000", _
        ws, "A1"

    MsgBox "Display OK"
    
CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit
    
End Sub

Public Sub Demo_DisplayFast() 'Demo Display Fast avec Méthode code C, chargement variable tableau (db.QueryFast)

    On Error GoTo Fail
    
    Dim db As New cDuck, ws As Worksheet, arr As Variant

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenReadOnly ThisWorkbook.Path & "\DbDuckDb.duckdb"

    Set ws = ThisWorkbook.Worksheets(1)

    arr = db.QueryFast("SELECT * FROM Instruments ORDER BY ISIN LIMIT 1000")

    Call ArrayToSheet(arr, ws, "A1")
    
    MsgBox "Display OK"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit
End Sub

Public Sub Demo_ImportCsv() 'Import CSV

    'Au préalable récupérer exemple data https://live.euronext.com/en/products/equities/list
    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, arr As Variant
    
    Dim t As New cHiPerfTimer, msQuery As Double
    
    t.Start

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenDuckDb ThisWorkbook.Path & "\DbDuckDb.duckdb"

    db.BeginTx
    db.ImportCsvReplace ThisWorkbook.Path & "\Euronext_Equities_2025-11-02.csv", "ImportedCsv"
    db.Commit
    
    Set ws = ThisWorkbook.Worksheets(1)
    arr = db.QueryFast("SELECT * FROM ImportedCsv;")
    Call ArrayToSheet(arr, ws, "A1")

    db.CloseDuckDb
    'MsgBox "Import CSV OK"
    
    msQuery = t.StopMilliseconds
    Debug.Print "QueryFast(SELECT * ...) : "; Format(msQuery, "0.000"); " ms"
    
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

Public Sub Demo_Brackets() 'Demo Bracker

    'Test pour code C (parsing)
    'pour [ Name ]
    'dll accepte WHERE [CODE ISIN]='FR0000987654' ou WHERE [CODE ISIN]="FR000098765"

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet
    
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"   ' ou ":memory:" pour un test RAM

    ' --- Table de test avec identifiants entre crochets ---
    db.Exec "CREATE TABLE IF NOT EXISTS [TestIsin](" & _
            "  [CODE ISIN] TEXT," & _
            "  [NumeroContrat] TEXT," & _
            "  [Prix] DOUBLE," & _
            "  [Modified At] TIMESTAMP" & _
            ");"
    db.Exec "DELETE FROM [TestIsin];"

    ' --- Inserts ---
    db.BeginTx
    db.Exec "INSERT INTO [TestIsin] ([CODE ISIN],[NumeroContrat],[Prix],[Modified At]) " & _
            "VALUES ('FR0000123456','C-001',101.25,NOW());"
    db.Exec "INSERT INTO [TestIsin] ([CODE ISIN],[NumeroContrat],[Prix],[Modified At]) " & _
            "VALUES ('FR0000987654','C-002', 99.80,NOW());"
    db.Commit

    ' --- Affichage sur la feuille ---
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear
    db.SelectToSheet _
        "SELECT [CODE ISIN], [NumeroContrat], [Prix], " & _
        "       strftime([Modified At], '%Y-%m-%d %H:%M:%S') AS [Modified At] " & _
        "FROM [TestIsin] " & _
        "WHERE [CODE ISIN]='FR0000987654' ;", _
        ws, "A1"
        
    MsgBox "Test Bracket & Parse"
CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.ExecOK "ROLLBACK;"
    db.ExecOK "CHECKPOINT;"
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

