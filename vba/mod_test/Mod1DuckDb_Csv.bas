Attribute VB_Name = "Mod1DuckDb_Csv"
Option Explicit


'CREATE TABLE ontime AS
    'SELECT * FROM read_csv('flights.csv');
'DESCRIBE ontime;


'===============================================================================
'
' Objectif :
'   Centraliser les démonstrations CSV <-> DuckDB <-> Excel en VBA avec la classe cDuck (DLL bridge).

'     1) Importer un CSV dans DuckDB (rapide / reproductible) puis afficher un aperçu.
'     2) Produire un CSV en sortie (export) à partir d’une requête SQL DuckDB.
'     3) Diagnostiquer / auto-détecter la structure d’un CSV (sniff_csv, read_csv_auto).
'
' Points clés de performance
'   - Pour importer : privilégier COPY ... FROM (bulk load) et read_csv_auto(... LIMIT 0)
'     pour créer le schéma. Utiliser sample_size=-1 uniquement si nécessaire (full scan).
'   - Pour exporter : privilégier COPY (SELECT ...) TO ... ou db.SelectToCsv (wrapper).
'   - Pour afficher : utiliser QueryFast + ArrayToSheet sur un extrait (LIMIT) afin
'     d’éviter d’inonder Excel.
'
' Contenu du module (démos)
'   - TestMakeSampleCSV
'       Génère un CSV de test random (MakeSampleCsv) pour les essais.
'
'   - Demo_Euronext_Csv
'       Exemple d’import CSV réel (Euronext) via DuckDbReadImportCsv
'       (création de table + COPY massif + affichage Excel optionnel).
'
'   - Demo_ImportCsv
'       Variante d’import utilisant db.ImportCsvReplace (API wrapper) + transaction.
'
'   - Demo_WriteCSV / Demo_WriteCSV2
'       Export CSV depuis DuckDB :
'         * Demo_WriteCSV  : SQL natif COPY (SELECT ...) TO ...
'         * Demo_WriteCSV2 : wrapper db.SelectToCsv + affichage du résultat dans Excel
'
'   - Demo_CsvJoin
'       Exemple d’export CSV basé sur une requête JOIN, puis affichage Excel.
'
'   - Demo_CSV_AutoDetect
'       Diagnostic d’un CSV :
'         * read_csv_auto(sample_size=-1) pour un aperçu (ou filtrage)
'         * sniff_csv(sample_size=-1) pour récupérer les propriétés détectées
'
'===============================================================================

Sub TestMakeSampleCSV()
    
    'création d'un csv data random : ThisWorkbook.Path & "\data.csv"
    Call MakeSampleCsv

End Sub

Sub Demo_Euronext_Csv()

    '===============================================================================
    ' DuckDbReadImportCsv
    '
    ' Importe un fichier CSV dans DuckDB (création de table + chargement massif),
    ' puis (optionnel) affiche dans Excel.
    
    'DuckDbReadImportCsv(ByVal duckPath As String, ByVal csvPath As String, ByVal tableName As String, _
                Optional ByVal delim As String = "auto", Optional ByVal replaceAll As Boolean = True, Optional ByVal displayPreview As Boolean = True)
    '
    ' Paramètres
    '   duckPath       : Chemin du fichier .duckdb (ex: ThisWorkbook.Path & "\demo.duckdb")
    '                    Utiliser ":memory:" pour une session en RAM.
    '   csvPath        : Chemin du fichier CSV à importer.
    '   tableName      : Nom de la table DuckDB cible (créée ou remplacée).
    '
    '   delim          : Délimiteur CSV.
    '                    - "auto" : détection automatique
    '                    - ","    : virgule
    '                    - ";"    : point-virgule
    '                    - "tab" ou "\t" : tabulation
    '
    '   replaceAll     : True/ False  -> DROP TABLE IF EXISTS + recréation + import (False -> conserve la table, la vide (DELETE) puis réimporte)
    '   displayPreview : True /False -> affiche un extrait dans Excel (feuille 1, A1)
    '   displaySheet   : Feuille cible (numéro 1..N ou nom "Feuil1")
    '===============================================================================

    Call DuckDbReadImportCsv(ThisWorkbook.Path & "\demo.duckdb", _
                ThisWorkbook.Path & "\Euronext_Equities_2025-11-02.csv", _
                "ImportedCsv", ";", , True, 1)
        
End Sub

Private Sub Demo_ImportCsv() 'Import CSV

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

Private Sub Demo_WriteCSV() 'CREATE CSV (memory) avec SQL natif

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, tmp As String, p As String

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb ":memory:"

    db.Exec "CREATE TABLE T (ISIN TEXT, Prix DOUBLE, Volume BIGINT);"
    db.Exec "INSERT INTO T VALUES " & _
            "('FR0000987654', 101.25,  250000)," & _
            "('FR0000123456',  99.80, 1200000);"

    tmp = ThisWorkbook.Path & "\duck_out.csv"
    p = Replace(tmp, "\", "/")
    db.Exec "COPY (SELECT ISIN, Prix, Volume FROM T ORDER BY ISIN) " & _
            "TO " & SqlQ(p) & " (FORMAT CSV, HEADER, DELIMITER ',', QUOTE '""', OVERWRITE 1);"

    ' Vérif I/O rapide puis nettoyage (supprime le fichier après test)
    Dim f As Integer: f = FreeFile
    Open tmp For Input As #f: Close #f
    'Kill tmp   ' <-- retire cette ligne si tu veux garder le CSV

    db.CloseDuckDb
    MsgBox "OK : CSV ISIN/Prix/Volume écrit (test passé).", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub

Private Sub Demo_WriteCSV2() 'CREATE CSV avec méthode SelectToCsv + affichage résultat

    Dim db As New cDuck, v As Variant, duckPath As String, outCsv As String, sql As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\test.duckdb"
    outCsv = ThisWorkbook.Path & "\exports\result_simple.csv"

    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckPath

    db.Exec "CREATE TABLE IF NOT EXISTS demo(i INT, s VARCHAR);"
    db.Exec "DELETE FROM demo;"
    db.Exec "INSERT INTO demo VALUES (1,'a'),(2,'b'),(3,'c');"

    sql = "SELECT * FROM demo ORDER BY i;"

    Call EnsureFolderExists(outCsv)
    On Error Resume Next: Kill outCsv: On Error GoTo Fail
    db.SelectToCsv sql, outCsv

    '--- Afficher le résultat dans Excel (relit depuis DuckDB)
    v = db.QueryFast(sql)
    If Not IsEmpty(v) Then
        ArrayToSheet v, ThisWorkbook.Worksheets(1), "A1"
    End If

    db.CloseDuckDb
    MsgBox "CSV écrit : " & outCsv & vbCrLf & "Résultat affiché en Feuil1.", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur Demo_WriteCSV_Fct : " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

Public Sub Demo_CsvJoin()

    Dim db As New cDuck, v As Variant, duckPath As String, outCsv As String, sql As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\test.duckdb"
    outCsv = ThisWorkbook.Path & "\exports\result_join.csv"

    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckPath

    db.Exec "CREATE TABLE IF NOT EXISTS clients(id INT, nom VARCHAR);"
    db.Exec "CREATE TABLE IF NOT EXISTS ventes(id INT, montant DOUBLE, client_id INT);"
    db.Exec "DELETE FROM clients; DELETE FROM ventes;"
    db.Exec "INSERT INTO clients VALUES (1,'Alpha'),(2,'Beta'),(3,'Gamma');"
    db.Exec "INSERT INTO ventes  VALUES (10, 99.9, 1),(11, 50, 2),(12, 12.5, 1),(13, 200, 3);"

    sql = "SELECT v.id AS vente_id, c.nom AS client, v.montant " & _
          "FROM ventes v JOIN clients c ON c.id = v.client_id " & _
          "ORDER BY v.id;"

    EnsureFolderExists outCsv
    On Error Resume Next: Kill outCsv: On Error GoTo Fail
    db.SelectToCsv sql, outCsv

    '--- Afficher le résultat dans Excel (relit depuis DuckDB)
    v = db.QueryFast(sql)
    If Not IsEmpty(v) Then
        ArrayToSheet v, ThisWorkbook.Worksheets(1), "A1"
    End If

    db.CloseDuckDb
    MsgBox "CSV écrit : " & outCsv & vbCrLf & "Résultat affiché en Feuil1.", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur Demo_CsvJoin : " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub


' =============== DEMO CSV (auto_detect + sniff_csv) — style A ==================
Public Sub Demo_CSV_AutoDetect()

    'lancez d'abord TestMakeSampleCSV
    
    On Error GoTo Fail

    '0) Session DuckDB
    Dim db As New cDuck, ws As Worksheet, csvPath$, p$, v As Variant, y As Long, sqlAuto$, sqlSniff$
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"   ' ou ":memory:"

    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear

    '1) Chemin CSV à analyser
    csvPath = ThisWorkbook.Path & "\data.csv"     ' adapte si besoin
    p = Replace(csvPath, "\", "/")
    y = 1

    'Titre
    ws.Cells(y, 1).Value = "CSV auto-détection DuckDB (auto + sniffer)"
    ws.Cells(y, 1).Font.Bold = True
    y = y + 2
    
    '***************************
    '2) Aperçu direct : read_csv_auto + sample_size=-1 (full scan)
    sqlAuto = "SELECT * FROM read_csv_auto(" & SqlQ(p) & ", sample_size=-1);"
    
    'sqlAuto = "SELECT * FROM read_csv_auto(" & SqlQ(p) & ", sample_size=-1) WHERE ISIN='FR0000000005';"
    'SELECT * FROM read_csv_auto('ton.csv', delim=';');
    
    v = db.QueryFast(sqlAuto)
    ws.Cells(y, 1).Value = "Aperçu (read_csv_auto, sample_size=-1) — 50 premières lignes"
    ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2

    '3) Sniffer : propriétés détectées (delim, quote, header, columns, formats, etc.)
    sqlSniff = "SELECT * FROM sniff_csv(" & SqlQ(p) & ", sample_size=-1);"
        
    v = db.QueryFast(sqlSniff)
    ws.Cells(y, 1).Value = "sniff_csv() — propriétés détectées"
    ws.Cells(y, 1).Font.Bold = True: y = y + 1
    ws.Range("A" & y).Resize(UBound(v, 1), UBound(v, 2)).Value = v
    y = y + UBound(v, 1) + 2
    
    MsgBox "OK — Auto-détection CSV & sniffer exécutés. Résultats sur la feuille 1.", vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub
