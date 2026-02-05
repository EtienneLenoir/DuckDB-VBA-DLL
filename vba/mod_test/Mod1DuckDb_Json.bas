Attribute VB_Name = "Mod1DuckDb_Json"
Option Explicit

'===============================================================================
'   Démonstrations JSON <-> DuckDB depuis Excel/VBA via la classe cDuck
' Ce module couvre 3 usages principaux :
'   1) Exporter une requête DuckDB vers un fichier JSON / NDJSON
'   2) Importer un JSON dans DuckDB (read_json_auto) puis afficher dans Excel
'   3) Manipuler des données côté DuckDB (INSERT préparé, lecture glob *.json, etc.)
'
'-------------------------------------------------------------------------------
' Formats JSON : JSON vs NDJSON
'   - NDJSON (Newline Delimited JSON) : 1 objet JSON par ligne.
'       * Très rapide en streaming (écriture/lecture séquentielle).
'       * Idéal pour gros volumes.
'       * C’est souvent ce que produit DuckDB via COPY ... (FORMAT JSON).
'
'   - JSON "array" (un seul fichier JSON avec une liste d’objets) :
'       * Plus classique pour certains outils (API, front).
'       * Souvent moins streaming-friendly.
'       * Nécessite généralement une agrégation (json_group_array / row_to_json)
'         plutôt qu’un simple COPY.
'===============================================================================

Public Sub Demo_Json_Export()

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, outFile As String, p As String, BoolJson As Boolean

    outFile = ThisWorkbook.Path & "\data\Exndjson.ndjson"
    EnsureFolderExists outFile
    p = Replace(outFile, "\", "/")

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    
    BoolJson = True
    
    'Table + données
    db.Exec "CREATE TABLE Exndjson(userId UBIGINT, id UBIGINT, title VARCHAR, completed BOOLEAN);"
    db.Exec "INSERT INTO Exndjson VALUES (1,1,'buy milk', false), (1,2,'send email', true), (2,3,'ship v1', false);"

    If BoolJson = True Then
        'tableau JSON
        db.Exec "COPY (SELECT * FROM Exndjson ORDER BY userId, id) TO " & SqlQ(p) & " (ARRAY);"
    Else
        'newline-delimited JSON (NDJSON) :: 1 ligne = 1 objet JSON
        db.Exec "COPY (SELECT * FROM Exndjson ORDER BY userId, id) TO " & SqlQ(p) & " (FORMAT JSON);"
    End If

    'Aperçu Excel
    a = db.QueryFast("SELECT * FROM Exndjson ORDER BY userId, id;")
    If Not IsEmpty(a) Then ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

    db.CloseDuckDb
    MsgBox "NDJSON écrit : " & outFile, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur NDJSON: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub

'json.duckdb_extension --> read_json_auto / read_json
Public Sub Demo_Json_Display() 'Import Json fichier précédent + Display

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, p As String

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    
    'Fichier Exemple json
    p = Replace(ThisWorkbook.Path & "\instruments.json", "\", "/")

    'extension json (selon builds)
    On Error Resume Next
        db.LoadExt "json"
    On Error GoTo Fail

    'Chargement Array
    a = db.QueryFast("SELECT * FROM read_json_auto(" & SqlQ(p) & "); ")

    'Affichage
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

    db.CloseDuckDb
    MsgBox "Import NDJSON OK : Instruments_from_json (aperçu affiché)", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    
End Sub

'------------------------------------------------------------------------------
' Function : CopyToJson
' Exporte le résultat d'un SELECT DuckDB vers un fichier JSON.
'
' Arguments:
'   selectSql       : Requête SELECT (sans le point-virgule final idéalement).
'   outJson         : Chemin complet du fichier de sortie.
'   overwrite       : (Optionnel) True = supprime le fichier si déjà présent.
'                     Défaut = True.
'   boolJsonArray   : (Optionnel) True = écrit un tableau JSON unique
'                     ([{...},{...}])
'                     False = écrit du NDJSON (1 objet JSON par ligne).
'                     Défaut = False.
'------------------------------------------------------------------------------
Public Sub Demo_Json_FunctionExport() 'Export JSON Function CopyToJson

    On Error GoTo Fail

    Dim db  As New cDuck, ws As Worksheet, a As Variant, outFile As String, outNorm As String, sql As String
    Dim t   As New cHiPerfTimer, ms As Double

    outFile = ThisWorkbook.Path & "\data\instruments.json"
    Call EnsureFolderExists(outFile)
    outNorm = Replace(outFile, "\", "/")

    '--- Open DuckDB
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    '--- (best-effort) json extension (selon build)
    On Error Resume Next
        db.TryLoadExt "json"
    On Error GoTo Fail

    '--- Table + seed idempotent
    db.Exec "CREATE TABLE IF NOT EXISTS Instruments(" & _
            " ISIN TEXT, NumeroContrat TEXT, Prix DOUBLE, ModifiedAt TIMESTAMP);"

    db.Exec "INSERT INTO Instruments " & _
            "SELECT 'FR0000123456','C-001',103.10, NOW() " & _
            "WHERE NOT EXISTS (SELECT 1 FROM Instruments WHERE ISIN='FR0000123456');"

    '--- Export
    sql = "SELECT ISIN, NumeroContrat, Prix, ModifiedAt " & _
          "FROM Instruments " & _
          "ORDER BY ISIN, ModifiedAt;"

    On Error Resume Next
        'Kill outFile
    On Error GoTo Fail

    t.Start
    
    '------------------------------------------------------------------------------
    ' CopyToJson
    ' Exporte le résultat d'un SELECT DuckDB vers un fichier JSON.
    '------------------------------------------------------------------------------
    db.CopyToJson sql, outNorm, , True
    
    ms = t.StopMilliseconds

    '--- (optionnel) petit aperçu Excel (limité)
    a = db.QueryFast("SELECT * FROM Instruments ORDER BY ISIN, ModifiedAt LIMIT 200;")
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear
    If Not IsEmpty(a) Then
        ws.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
        ws.Columns.AutoFit
    End If

    db.CloseDuckDb

    Debug.Print "Export JSON | Temps : " & Format$(ms, "0.000") & " ms | " & outFile
    MsgBox "Export JSON OK : " & outFile & vbCrLf & _
           "Aperçu affiché en Feuil1." & vbCrLf & _
           "Temps : " & Format$(ms, "0.000") & " ms", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur export NDJSON : " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

Public Sub Demo_Json_Export2()

    On Error GoTo Fail

    Dim db As New cDuck, outFile As String, p As String, sql As String, a As Variant

    outFile = ThisWorkbook.Path & "\data\Exjson.json"
    EnsureFolderExists outFile
    p = Replace(outFile, "\", "/")

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'Extension json (souvent nécessaire pour to_json / read_json)
    On Error Resume Next
        db.TryLoadExt "json"
    On Error GoTo Fail

    'Table + données
    db.Exec "CREATE TABLE Exjson(userId UBIGINT, id UBIGINT, title VARCHAR, completed BOOLEAN);"
    db.Exec "INSERT INTO Exjson VALUES (1,1,'buy milk', false), (1,2,'send email', true), (2,3,'ship v1', false);"

    '1) Produit UNE seule valeur texte JSON = un array
    '2) Ecrit cette valeur dans un fichier
    sql = _
        "COPY (" & _
        "  SELECT to_json(list(struct_pack(" & _
        "    userId := userId," & _
        "    id := id," & _
        "    title := title," & _
        "    completed := completed" & _
        "  ) ORDER BY userId, id)) AS json_text" & _
        "  FROM Exjson" & _
        ") TO " & SqlQ(p) & " (" & _
        "  FORMAT CSV, HEADER 0, DELIMITER '', QUOTE '', ESCAPE ''" & _
        ");"

    db.Exec sql

    'Aperçu Excel
    a = db.QueryFast("SELECT * FROM Exjson ORDER BY userId, id;")
    If Not IsEmpty(a) Then ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

    db.CloseDuckDb
    MsgBox "JSON array écrit : " & outFile, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur JSON array: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub

'===============================================================================
' Demo_ReadJson_SubsetColumns
'   - écrit un JSON array (ExSubsetColumns.json) avec DuckDB
'   - lit le JSON en ne gardant qu'un sous-ensemble de colonnes via columns={...}
'   - affiche dans Excel
'===============================================================================
Public Sub Demo_ReadJson_SubsetColumns()

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, outFile As String, p As String, sqlWrite As String, sqlRead As String

    outFile = ThisWorkbook.Path & "\data\ExSubsetColumns.json"
    Call EnsureFolderExists(outFile)
    p = Replace(outFile, "\", "/")

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'JSON extension (souvent nécessaire pour read_json / to_json selon build)
    On Error Resume Next
        db.TryLoadExt "json"
    On Error GoTo Fail

    '1) Données source
    db.Exec "CREATE TABLE ExSubsetColumns(userId UBIGINT, id UBIGINT, title VARCHAR, completed BOOLEAN);"
    db.Exec "INSERT INTO ExSubsetColumns VALUES " & _
            "(1,1,'buy milk', false)," & _
            "(1,2,'send email', true)," & _
            "(2,3,'ship v1', false);"

    '2) Ecriture JSON array -> fichier (1 seule valeur JSON dans le fichier)
    '   (astuce : on écrit le texte JSON via COPY ... FORMAT CSV)
    sqlWrite = _
        "COPY (" & _
        "  SELECT to_json(list(struct_pack(" & _
        "    userId := userId," & _
        "    id := id," & _
        "    title := title," & _
        "    completed := completed" & _
        "  ) ORDER BY userId, id)) AS json_text" & _
        "  FROM ExSubsetColumns" & _
        ") TO " & SqlQ(p) & " (" & _
        "  FORMAT CSV, HEADER 0, DELIMITER '', QUOTE '', ESCAPE ''" & _
        ");"
    db.Exec sqlWrite

    '3) Lecture JSON : on ne spécifie que userId et completed
    '   => id + title sont EXCLUS du résultat
    sqlRead = _
        "SELECT * " & _
        "FROM read_json(" & SqlQ(p) & ", " & _
        "  format='array', " & _
        "  columns={userId: 'UBIGINT', completed: 'BOOLEAN'}" & _
        ") " & _
        "LIMIT 5;"

    a = db.QueryFast(sqlRead)
    If Not IsEmpty(a) Then
        ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"
    End If

    db.CloseDuckDb
    MsgBox "OK : JSON écrit puis lu (subset columns) -> affiché en Feuil1", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur Demo_ReadJson_SubsetColumns: " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

'===============================================================================
' Demo_JSON_ExportMemory
'
' Objectif
'   - Créer un petit jeu de données en DuckDB (session :memory:)
'   - Exporter en NDJSON via db.CopyToJson (1 ligne JSON = 1 enregistrement)
'   - Afficher un aperçu dans Excel (depuis DuckDB, pas en relisant le fichier)
'
' Notes
'   - Le support JSON peut être natif ou via extension selon le build DuckDB.
'     On fait un "best-effort" TryLoadExt("json") sans rendre la démo dépendante.
'   - Les chemins sont normalisés en "/" pour DuckDB.
'===============================================================================
Public Sub Demo_JSON_ExportMemory()

    On Error GoTo Fail

    Dim db As New cDuck, t As New cHiPerfTimer, ws As Worksheet, a As Variant, ms As Double, dataDir As String, outFile As String, p As String

    '--- dossier & sortie
    dataDir = ThisWorkbook.Path & "\data"
    outFile = dataDir & "\sample.json"
    p = Replace(outFile, "\", "/")

    '--- DuckDB : session mémoire
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'best-effort : JSON (selon builds)
    On Error Resume Next
    db.TryLoadExt "json"
    On Error GoTo Fail

    t.Start

    '--- Table & seed
    db.Exec "DROP TABLE IF EXISTS T;"
    db.Exec _
        "CREATE TABLE T(" & _
        "  ISIN          TEXT," & _
        "  NumeroContrat TEXT," & _
        "  Prix          DOUBLE," & _
        "  ModifiedAt    TIMESTAMP" & _
        ");"

    db.Exec _
        "INSERT INTO T " & _
        "SELECT " & _
        "  'FR' || lpad(CAST(i AS VARCHAR), 10, '0')              AS ISIN," & _
        "  'C-' || lpad(CAST(i % 1000 AS VARCHAR), 3, '0')        AS NumeroContrat," & _
        "  50 + (i % 1000) / 10.0                                 AS Prix," & _
        "  CURRENT_TIMESTAMP - i * INTERVAL '1 minute'            AS ModifiedAt " & _
        "FROM range(1, 201) AS t(i);"

    '--- Export NDJSON
    On Error Resume Next
        Kill outFile
    On Error GoTo Fail

    db.CopyToJson "SELECT * FROM T ORDER BY ModifiedAt DESC", p

    ms = t.StopMilliseconds

    '--- Aperçu Excel : relu depuis DuckDB (plus fiable/rapide que relire le fichier)
    a = db.QueryFast("SELECT * FROM T ORDER BY ModifiedAt DESC LIMIT 200;")

    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear
    If Not IsEmpty(a) Then
        ws.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
        ws.Columns.AutoFit
    End If

    db.CloseDuckDb

    Debug.Print "JSON export NDJSON | Rows=200 preview | Temps : " & Format$(ms, "0.000") & " ms | " & outFile
    MsgBox "Échantillon JSON créé : " & outFile & vbCrLf & _
           "Aperçu affiché en feuille 1." & vbCrLf & _
           "Temps : " & Format$(ms, "0.000") & " ms", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur Demo_JSON_ExportMemory : " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

Public Sub Demo_JSON_Read() 'READ JSON METHODE ReadToArray

    On Error GoTo Fail
    
    Dim db As New cDuck, v As Variant, p As String
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    db.TryLoadExt "json"

    p = Replace(ThisWorkbook.Path & "\data\*.json", "\", "/")
    v = db.ReadToArray(p, "ORDER BY 1 LIMIT 200")

    Call ArrayToSheet(v, ThisWorkbook.Worksheets(1), "A1")
    
Clean: db.CloseDuckDb: Exit Sub
Fail:  MsgBox Err.Description, vbExclamation: Resume Clean

End Sub

Public Sub Import_Json_Exemple()

    'https://raw.githubusercontent.com/prust/wikipedia-movie-data/master/movies.json
    
    On Error GoTo Fail

    Dim db As New cDuck, a As Variant
    Dim jsonPath As String, duckPath As String, duckTable As String, p As String, sql As String

    jsonPath = ThisWorkbook.Path & "\movies.json"  'ThisWorkbook.Path & "\movies.json"     ' <-- adapte
    duckPath = ThisWorkbook.Path & "\cache.duckdb"
    duckTable = "MoviesJson"

    p = Replace(jsonPath, "\", "/")

    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckPath

    'best-effort : selon build, JSON peut être extension
    On Error Resume Next
        db.TryLoadExt "json"
    On Error GoTo Fail

    db.Exec "DROP TABLE IF EXISTS " & db.QuoteIdent(duckTable) & ";"
    sql = "CREATE TABLE " & db.QuoteIdent(duckTable) & " AS " & _
          "SELECT * FROM read_json_auto(" & SqlQ(p) & ");"
    db.Exec sql


    db.CloseDuckDb
    MsgBox "OK: JSON importé (auto) -> " & duckTable, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur import JSON: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub

Public Sub Demo_JSON_Insert() 'JSON INSERT

    On Error GoTo Fail
    
    Dim db As New cDuck, ps As LongPtr, v As Variant
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    db.Exec "CREATE TABLE T(isin TEXT, px DOUBLE);"
    ps = db.Prepare("INSERT INTO T(isin, px) VALUES (?,?);")

    db.PS_BindText ps, 1, "FR0000000001"
    db.PS_BindDouble ps, 2, 101.25
    db.PS_Exec ps

    db.PS_BindText ps, 1, "FR0000000002"
    db.PS_BindDouble ps, 2, 99.8
    db.PS_Exec ps

    db.PS_CloseDuckDb ps

    v = db.QueryFast("SELECT * FROM T ORDER BY isin;")
    Call ArrayToSheet(v, ThisWorkbook.Worksheets(1), "A1")

Clean: db.CloseDuckDb: Exit Sub
Fail:  On Error Resume Next: db.PS_CloseDuckDb ps: db.CloseDuckDb
      MsgBox Err.Description, vbExclamation: Resume Clean
End Sub
