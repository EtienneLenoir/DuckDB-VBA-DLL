Attribute VB_Name = "Mod1DuckDb_Parquet"
Option Explicit

'===============================================================================

' Objectif :
'   - Lire des fichiers Parquet (read_parquet / replacement scan FROM 'file.parquet')
'     et ramener un Variant(2D) dans Excel (QueryFast / ReadToArray / SelectToSheet).
'   - Exporter des tables / requetes vers Parquet (COPY ... (FORMAT parquet))
'     ou via db.CopyToParquet (wrapper DLL, compression ZSTD par defaut).

' Méthodes DuckDB "Parquet":
'   1) Replacement scan (lecture directe) :
'        CREATE TABLE test AS SELECT * FROM 'test.parquet';
'        -- ou sans créer de table :
'        SELECT * FROM 'test.parquet';
'
'   2) Fonction dédiée (lecture) :
'        SELECT * FROM read_parquet('test.parquet');
'        SELECT * FROM read_parquet(['file1.parquet','file2.parquet','file3.parquet']);
'
'   3) Vue (lecture réutilisable) :
'        CREATE VIEW people AS
'        SELECT * FROM read_parquet('test.parquet');
'
'   4) Export Parquet (COPY) :
'        COPY tbl TO 'output.parquet' (FORMAT parquet);
'        COPY (SELECT * FROM tbl) TO 'output.parquet' (FORMAT parquet);
'
'   5) Conversion CSV -> Parquet (COPY + options) :
'        COPY 'test.csv' TO 'result-uncompressed.parquet'
'        (FORMAT parquet, COMPRESSION uncompressed);
'        -- Variante plus universelle si besoin :
'        COPY (SELECT * FROM read_csv_auto('test.csv'))
'        TO 'result-uncompressed.parquet' (FORMAT parquet, COMPRESSION uncompressed);

'===============================================================================

Public Sub Demo_Parquet_To_Array_NoDB()

    On Error GoTo Fail
    
    Dim db As New cDuck, p As String, a As Variant

    'Init + base en mémoire
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'Extension parquet (best-effort)
    On Error Resume Next
    db.LoadExt "parquet"
    On Error GoTo Fail

    'Lecture du parquet -> Variant(2D)
    p = Replace(ThisWorkbook.Path & "\export.parquet", "\", "/")
    a = db.QueryFast( _
        "SELECT * FROM read_parquet(" & SqlQ(p) & ") " & _
        "WHERE a IS NOT NULL ORDER BY a LIMIT 1000")

    'Dump dans la feuille
    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        If Not IsEmpty(a) Then
            .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
        End If
    End With

    db.CloseDuckDb
    MsgBox "OK — Parquet lu directement en mémoire", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub

Public Sub Demo_TempList_FromParquet_Free()

    On Error GoTo Fail

    ' === OUVERTE EN MÉMOIRE (équivalent de h = Duck_Open(":memory:") ) ===
    Dim db As New cDuck, keys As Variant, vOut As Variant, src As String, outParquet As String, sql As String
    
    Call db.Init(ThisWorkbook.Path)
    db.OpenDuckDb ":memory:"


    '(best-effort) charge l’extension parquet
    On Error Resume Next
        db.LoadExt "parquet"
    On Error GoTo Fail

    ' Exporte Instruments -> Parquet (si tu veux partir du cache)
    src = ThisWorkbook.Path & "\cache.duckdb"
    outParquet = Replace(ThisWorkbook.Path & "\instruments.parquet", "\", "/")

    ' -- Optionnel : ré-export rapide (sinon laisse commenté)
    'Dim db2 As New cDuck
    'db2.Init ThisWorkbook.Path
    'db2.OpenDuckDb src
    'db2.CopyToParquet "SELECT * FROM Instruments", outParquet
    'db2.CloseDuckDb

    ' Clés à filtrer
    keys = Array("FR0000123459", "FR0000123460")

    ' SQL libre + JOIN sur la temp table que la DLL va créer
    sql = "SELECT q.* " & _
          "FROM read_parquet('" & outParquet & "') q " & _
          "JOIN __tmp_list t ON q.ISIN = t.v " & _
          "ORDER BY q.ISIN"

    'Création de la liste temporaire + exécution
    vOut = db.SelectWithTempList("__tmp_list", keys, "VARCHAR", sql, "", False)

    'Affichage
    ArrayToSheet vOut, ThisWorkbook.Worksheets("test"), "A1"
    MsgBox "OK (mode libre) — " & (UBound(vOut, 1) - 1) & " lignes.", vbInformation

Clean:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume Clean
    
End Sub

Sub Smoke_LoadParquet()

    Dim db As New cDuck, v As Variant, pattern As String
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    Call db.TryLoadExt("parquet")

    'OK : fichier unique connu
    v = db.ReadToArray(ThisWorkbook.Path & "\access_table.parquet") 'export.parquet
    ArrayToSheet v, ThisWorkbook.Worksheets(1), "A1"

    'Motif : vérifier qu’il y a au moins un match
    pattern = ThisWorkbook.Path & "\access_table.parquet"
    If Dir$(pattern, vbNormal + vbHidden + vbSystem) = "" Then
        MsgBox "Aucun .parquet ne correspond à : " & pattern, vbExclamation
        GoTo Clean
    End If
    v = db.ReadToArray(pattern, "WHERE Price > 100")
    ArrayToSheet v, ThisWorkbook.Worksheets(1), "A1"

Clean:
    db.CloseDuckDb
End Sub

Sub Rtt_B()
    On Error GoTo Fail

    Dim db As New cDuck, a1 As Variant, a2 As Variant, a3 As Variant, patParquet As String, pForDuck As String

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    
    'best-effort: extensions utiles
    Call db.TryLoadExt("parquet")
    Call db.TryLoadExt("json")

    'PARQUET (avec glob) ------------------------------------
    patParquet = ThisWorkbook.Path & "\data\*.parquet"
    ' évite l'IO Error de DuckDB si aucun fichier ne matche
    If Dir$(patParquet, vbNormal) <> "" Then
        pForDuck = Replace(patParquet, "\", "/")
        a1 = db.ReadToArray(pForDuck, "WHERE px > 100 ORDER BY isin")
    Else
        'tableau vide avec en-tête minimal pour éviter IsEmpty
        ReDim a1(1 To 1, 1 To 1): a1(1, 1) = "No parquet file found"
    End If

    'JSON (auto-schema) -------------------------------------
    a2 = db.ReadToArray(ThisWorkbook.Path & "\dump.json", "LIMIT 1000")

    'CSV (inférence étendue) --------------------------------
    a3 = db.ReadToArray(ThisWorkbook.Path & "\report.csv")

    'Affichage (a1 en A1 ; a2/a3 commentés si besoin)
    ArrayToSheet a1, ThisWorkbook.Worksheets(1), "A1"
    'ArrayToSheet a2, ThisWorkbook.Worksheets(1), "H1"
    'ArrayToSheet a3, ThisWorkbook.Worksheets(1), "N1"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Sub Make_Parquet_Sample()

    On Error GoTo Fail

    'Dossier data + chemin parquet
    
    Dim db As New cDuck, dataDir As String, outFile As String, p As String
    
    dataDir = ThisWorkbook.Path & "\data"
    If Dir$(dataDir, vbDirectory) = vbNullString Then MkDir dataDir
    outFile = dataDir & "\sample.parquet"
    p = Replace(outFile, "\", "/")

    '-- 2) Session DuckDB + (best-effort) extension parquet
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    On Error Resume Next
    db.TryLoadExt "parquet"
    On Error GoTo Fail

    ' -- 3) Table + données
    db.Exec "DROP TABLE IF EXISTS T;"
    db.Exec _
        "CREATE TABLE T(" & _
        "  ISIN TEXT," & _
        "  NumeroContrat TEXT," & _
        "  Prix DOUBLE," & _
        "  ModifiedAt TIMESTAMP" & _
        ");"

    ' NOTE: lpad(...), % (mod), et INTERVAL avec multiplication
    db.Exec _
        "INSERT INTO T " & _
        "SELECT " & _
        "  'FR' || lpad(CAST(i AS VARCHAR), 10, '0')              AS ISIN," & _
        "  'C-' || lpad(CAST(i % 100 AS VARCHAR), 3, '0')         AS NumeroContrat," & _
        "  50 + (i % 1000) / 10.0                                 AS Prix," & _
        "  CURRENT_TIMESTAMP - i * INTERVAL '1 minute'            AS ModifiedAt " & _
        "FROM range(1, 1001) AS t(i);"

    ' -- 4) Export Parquet
    db.CopyToParquet "SELECT * FROM T ORDER BY ModifiedAt DESC", p
    db.CloseDuckDb

    MsgBox "Échantillon Parquet créé : " & outFile, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur Make_Parquet_Sample_A : " & Err.Description, vbExclamation
End Sub


Sub Demo_Select_To_Sheet_FromParquet()

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, p As String, sql As String
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    
    On Error Resume Next
        db.TryLoadExt "parquet"
    On Error GoTo Fail

    Set ws = ThisWorkbook.Worksheets(1)
    p = Replace(ThisWorkbook.Path & "\data\*.parquet", "\", "/")  ' IMPORTANT: slashs avant

    sql = _
        "SELECT * " & vbCrLf & _
        "FROM read_parquet(" & SqlQ(p) & ") " & vbCrLf & _
        "WHERE Prix IS NOT NULL " & vbCrLf & _
        "ORDER BY ModifiedAt DESC " & vbCrLf & _
        "LIMIT 1000;"

    ' 3) Vers la feuille (via COPY UTF-8 dans ta méthode)
    db.SelectToSheet sql, ws

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur Parquet: " & Err.Description, vbExclamation
End Sub

'=== Auto-init en mémoire + import parquet + index ===
Public Sub AutoInit_DuckDB()

    On Error GoTo Fail

    ' 1) Session DuckDB en mémoire (reste vivante via le singleton m_singleton)
    Dim db As New cDuck, p As String
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    On Error Resume Next
        db.TryLoadExt "parquet"
    On Error GoTo Fail

    p = Replace(ThisWorkbook.Path & "\export.parquet", "\", "/")
    db.Exec "CREATE TABLE IF NOT EXISTS Quotes AS SELECT * FROM read_parquet(" & SqlQ(p) & ");"

    'Index pour accélérer les recherches insensibles à la casse
    'db.Exec "CREATE INDEX IF NOT EXISTS ix_quotes_isin ON Quotes (lower(ISIN));"
    'db.Exec "CREATE INDEX IF NOT EXISTS ix_quotes_name ON Quotes (lower(Name));"

    'Garder la session ouverte pour le reste de l’appli
    '    (le singleton du module garde la référence -> le Class_Terminate ne s’exécute pas)
    'Set m_singleton = db

    'Optionnel : exposer le handle natif comme avant
    ThisWorkbook.Names.Add name:="__DUCK_HANDLE", RefersTo:="=" & CStr(db.handle)

    MsgBox "DuckDB prêt (mémoire + index).", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    If Not db Is Nothing Then db.CloseDuckDb
    'Set m_singleton = Nothing
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

Public Sub Demo_Ext_Parquet()

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, outPath As String

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdbX"

    db.LoadExt "parquet"

    db.Exec "CREATE TABLE IF NOT EXISTS P(a INT, s TEXT);"
    db.Exec "DELETE FROM P;"
    db.Exec "INSERT INTO P VALUES (1,'x'),(2,'y');"

    'COPY TO PARQUET
    outPath = ThisWorkbook.Path & "\out.parquet"
    db.CopyToParquet "SELECT * FROM P ORDER BY a", outPath

    'Relecture directe et affichage
    a = db.QueryFast("SELECT * FROM read_parquet(" & q(Replace(outPath, "\", "/")) & ") ORDER BY a")
    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    End With

    db.CloseDuckDb
    MsgBox "OK Parquet -> " & outPath, vbInformation
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

Public Sub Parquet_Export_Example()

    On Error GoTo Fail

    Dim db As New cDuck, outPath As String

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    'Donées d’exemple (optionnel)
    db.Exec "CREATE TABLE IF NOT EXISTS P(a INT, s TEXT);"
    db.Exec "DELETE FROM P;"
    db.Exec "INSERT INTO P VALUES (1,'x'),(2,'y');"

    'Export Parquet (compression ZSTD côté DLL)
    outPath = ThisWorkbook.Path & "\export.parquet"
    db.CopyToParquet "SELECT * FROM P ORDER BY a", outPath

    db.CloseDuckDb
    MsgBox "Export Parquet OK ? " & outPath, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub


Public Sub Parquet_Import_CreateOrReplace()

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, a As Variant, p As String
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    '(best-effort) extension parquet
    On Error Resume Next
        db.LoadExt "parquet"
    On Error GoTo Fail

    'Chemin du parquet
    p = Replace(ThisWorkbook.Path & "\export.parquet", "\", "/")

    '(re)créer la table depuis le fichier parquet
    db.Exec "DROP TABLE IF EXISTS P_in;"
    db.Exec "CREATE TABLE P_in AS SELECT * FROM read_parquet(" & q(p) & ");"

    'Vérif en RAM -> feuille
    Set ws = ThisWorkbook.Worksheets(1)
    a = db.QueryFast("SELECT * FROM P_in ORDER BY a")
    ArrayToSheet a, ws, "A1"

    db.CloseDuckDb
    MsgBox "Import Parquet ? table OK", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

Public Sub Parquet_Import_Append()

    On Error GoTo Fail

    Dim db As New cDuck, p As String

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    On Error Resume Next
        db.LoadExt "parquet"
    On Error GoTo Fail

    'Chemin du parquet
    p = Replace(ThisWorkbook.Path & "\export.parquet", "\", "/")

    'Append dans une table existante (schéma compatible requis)
    db.Exec "INSERT INTO P_in SELECT * FROM read_parquet(" & q(p) & ");"

    db.CloseDuckDb
    MsgBox "Append Parquet ? table OK", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

Public Sub Ex_Parquet_CreateTable_From_File()

    On Error GoTo Fail

    Dim db As New cDuck, pParquet As String, a As Variant

    db.Init ThisWorkbook.Path
    db.ErrorMode = 0 '0=Raise, 1=MsgBox, 2=LogOnly
    db.OpenDuckDb ":memory:"

    'Parquet extension (lecture + écriture)
    db.LoadExt "parquet"

    'Chemin du parquet
    pParquet = Replace(ThisWorkbook.Path & "\test.parquet", "\", "/")

    'Si le fichier n'existe pas, on en crée un petit pour que la démo tourne
    If Len(Dir(ThisWorkbook.Path & "\test.parquet")) = 0 Then
        db.Exec "CREATE TABLE src_people(id INT, name TEXT);"
        db.Exec "INSERT INTO src_people VALUES (1,'Alice'),(2,'Bob'),(3,'Chloe');"
        db.Exec "COPY src_people TO " & SqlQ(pParquet) & " (FORMAT parquet);"
    End If

    'Exemple 1 : Replacement scan (DuckDB) -> FROM 'file.parquet'
    db.Exec "DROP TABLE IF EXISTS test;"
    db.Exec "CREATE TABLE test AS SELECT * FROM " & SqlQ(pParquet) & ";"

    'Alternative (équivalente) :
    'db.Exec "CREATE TABLE test AS SELECT * FROM read_parquet(" & SqlQ(pParquet) & ");"

    a = db.QueryFast("SELECT * FROM test ORDER BY id;")
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

    db.CloseDuckDb
    MsgBox "OK  CREATE TABLE ... FROM 'test.parquet'", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation

End Sub

Public Sub Ex_Parquet_Copy_Table_To_Parquet()

    On Error GoTo Fail

    Dim db As New cDuck, outPath As String

    db.Init ThisWorkbook.Path
    db.ErrorMode = 0
    db.OpenDuckDb ":memory:"
    db.LoadExt "parquet"

    'Table exemple
    db.Exec "CREATE TABLE tbl(id INT, s TEXT);"
    db.Exec "INSERT INTO tbl VALUES (1,'x'),(2,'y'),(3,'z');"

    'Exemple 2 : COPY tbl TO 'output.parquet' (FORMAT parquet);
    outPath = Replace(ThisWorkbook.Path & "\output_from_table.parquet", "\", "/")
    db.Exec "COPY tbl TO " & SqlQ(outPath) & " (FORMAT parquet);"

    db.CloseDuckDb
    MsgBox "OK  COPY tbl TO ... (FORMAT parquet)" & vbCrLf & outPath, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation

End Sub


Public Sub Ex_Parquet_Copy_Select_To_Parquet()

    On Error GoTo Fail

    Dim db As New cDuck, outPath As String

    db.Init ThisWorkbook.Path
    db.ErrorMode = 0
    db.OpenDuckDb ":memory:"
    db.LoadExt "parquet"

    'Table exemple
    db.Exec "CREATE TABLE tbl(id INT, s TEXT);"
    db.Exec "INSERT INTO tbl VALUES (2,'b'),(1,'a'),(3,'c');"

    'Exemple 3 : COPY (SELECT * FROM tbl) TO 'output.parquet' (FORMAT parquet);
    outPath = Replace(ThisWorkbook.Path & "\output_from_select.parquet", "\", "/")
    db.Exec "COPY (SELECT * FROM tbl ORDER BY id) TO " & SqlQ(outPath) & " (FORMAT parquet);"

    db.CloseDuckDb
    MsgBox "OK  COPY (SELECT ...) TO ... (FORMAT parquet)" & vbCrLf & outPath, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation

End Sub

Public Sub Ex_Parquet_Create_View_From_Parquet()

    On Error GoTo Fail

    Dim db As New cDuck, pParquet As String, a As Variant

    db.Init ThisWorkbook.Path
    db.ErrorMode = 0
    db.OpenDuckDb ":memory:"
    db.LoadExt "parquet"

    pParquet = Replace(ThisWorkbook.Path & "\test.parquet", "\", "/")

    'Si absent, on crée un petit parquet
    If Len(Dir(ThisWorkbook.Path & "\test.parquet")) = 0 Then
        db.Exec "COPY (SELECT 1 AS id, 'Alice' AS name UNION ALL SELECT 2,'Bob') TO " & SqlQ(pParquet) & " (FORMAT parquet);"
    End If

    'Exemple 5 : CREATE VIEW people AS SELECT * FROM read_parquet('test.parquet');
    db.Exec "DROP VIEW IF EXISTS people;"
    db.Exec "CREATE VIEW people AS SELECT * FROM read_parquet(" & SqlQ(pParquet) & ");"

    a = db.QueryFast("SELECT * FROM people ORDER BY id;")
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

    db.CloseDuckDb
    MsgBox "OK  CREATE VIEW people AS read_parquet(...)", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation

End Sub

Public Sub Ex_Csv_To_Parquet_Uncompressed()

    On Error GoTo Fail

    Dim db As New cDuck, csvPath As String, outPath As String, pCsv As String, pOut As String

    db.Init ThisWorkbook.Path
    db.ErrorMode = 0
    db.OpenDuckDb ":memory:"
    db.LoadExt "parquet"

    csvPath = ThisWorkbook.Path & "\titanic.csv"
    outPath = ThisWorkbook.Path & "\result-uncompressed.parquet"
    pCsv = Replace(csvPath, "\", "/")
    pOut = Replace(outPath, "\", "/")

    'Exemple 6 : conversion CSV -> Parquet (compression uncompressed)
    'Si ta version DuckDB n'accepte pas la forme COPY 'file.csv' TO ..., utilise la variante read_csv_auto juste en dessous.
    On Error Resume Next
        db.Exec "COPY " & SqlQ(pCsv) & " TO " & SqlQ(pOut) & " (FORMAT parquet, COMPRESSION uncompressed);"
    On Error GoTo Fail

    If db.LastError <> "" Then
        'Fallback plus universel
        db.Exec "COPY (SELECT * FROM read_csv_auto(" & SqlQ(pCsv) & ")) TO " & SqlQ(pOut) & " (FORMAT parquet, COMPRESSION uncompressed);"
    End If

    db.CloseDuckDb
    MsgBox "OK  CSV -> Parquet (uncompressed)" & vbCrLf & outPath, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation

End Sub


