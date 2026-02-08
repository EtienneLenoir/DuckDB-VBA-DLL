Attribute VB_Name = "Mod2DuckDb_PrepStat"
Option Explicit
'Handles des prepared statements
Private m_psPrefix   As LongPtr
Private m_psContains As LongPtr

'===============================================================================
'
'   DuckDB prend en charge les requêtes préparées, où les paramètres sont substitués lors de l'exécution de la requête
'
'   DuckDB prend en charge les requêtes préparées, où les paramètres sont substitués lors de l'exécution de la requête.
'   Cela améliore la lisibilité et contribue à prévenir les injections SQ

        'Il existe trois syntaxes pour indiquer les paramètres dans les requêtes préparées :
            'auto-incrémentés ( ?)
            'positionnels ( $1)
            'nommés ( $param).
'
'            https://duckdb.org/docs/stable/sql/query_syntax/prepared_statements
'
' Thèmes couverts
'   1) Prepared Statements (Prepare / Bind / Exec)
'      - Evite de construire des strings SQL énormes dans VBA.
'      - Plus propre, plus sûr (réduit le risque d’injection SQL), et souvent plus rapide
'        quand on exécute plusieurs fois la même requête.
'
'   2) Exécution "SELECT -> Variant(2D)"
'      - db.QueryFast(...) renvoie un tableau 2D prêt à coller sur une feuille Excel.
'      - db.PrepStatQueryArray(ps) exécute un prepared statement et renvoie directement
'        un Variant(2D) (même format que QueryFast).
'
'   3) Ingestion "Variant(2D) -> DuckDB"
'      - db.AppendArray(table, variant2D, hasHeader) charge un tableau VBA (avec en-têtes)
'        dans DuckDB sans passer par CSV.
'
'   4) Recherche / filtrage "marché" (LIKE contains)
'      - Exemple finance : filtrer la table Importée (ImportedCsv) par colonne Market.
'
' Remarques importantes
'   - Les handles de prepared statements (ps) sont liés à une connexion/handle DuckDB.
'     => si DB fermé et réouverture, ne pas réutiliser un ps ancien.
'     => éviter "Static ps" si db recrées à chaque appel.
'
'
' Liste des procédures / fonctions de test
'
'   Test_ExecPreparedToArray_Simple
'     - Démo minimale :
'         * crée une DB :memory:
'         * crée une table t et insère 3 lignes
'         * prépare un SELECT paramétré (id > ?)
'         * bind le paramètre
'         * exécute via db.PrepStatQueryArray(ps) -> Variant(2D)
'         * colle le résultat dans Feuil1
'
'   Demo_Prepared_A
'     - Démo "prepared INSERT" + lecture:
'         * ouvre une DB fichier (cache.duckdb)
'         * reset la table T(isin, px)
'         * prépare un INSERT (?, ?)
'         * exécute plusieurs inserts via Bind + Exec
'         * SELECT final -> feuille 1
'
'   Test_Prepared_Insert_A
'     - Variante de Demo_Prepared_A avec switch :
'         * useMemory = True  -> DB en RAM (:memory:)
'         * useMemory = False -> DB fichier (cache.duckdb)
'       Puis création table + prepared inserts + SELECT -> feuille.
'
'   Test_AppendArrayV_Order_A
'     - Démo ingestion rapide :
'         * crée table T(ISIN, Prix, ModifiedAt)
'         * construit un Variant(2D) avec en-têtes en ligne 1
'         * db.AppendArray "T", v, True
'         * SELECT + format timestamp -> feuille 1
'
'===============================================================================

Public Sub Test_ExecPreparedToArray_Simple()

    On Error GoTo Fail

    Dim db As New cDuck, ps As LongPtr, v As Variant

    '1) Session DuckDB en mémoire
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2          ' log only (optionnel)
    db.OpenDuckDb ":memory:"

    '2) Petit jeu de données
    db.Exec "CREATE TABLE t (id INT, name TEXT);"
    db.Exec "INSERT INTO t VALUES (1, 'Alice'), (2, 'Bob'), (3, 'Charlie');"

    '3) Préparé avec paramètre (tu peux aussi utiliser db.Prepare si tu l'as déjà "sans Raise")
    ps = DuckVba_PrepareW(db.handle, StrPtr("SELECT * FROM t WHERE id > ? ORDER BY id;"))
    If ps = 0 Then
        GoTo Bye
    End If

    '4) Bind param = 1
    db.PS_BindInt64 ps, 1, 1

    '5) Exécuter en renvoyant directement un Variant(2D) via ta méthode
    v = db.PrepStatQueryArray(ps)

    '6) Coller dans Feuil1!A1
    If Not IsEmpty(v) Then
        ArrayToSheet v, ThisWorkbook.Worksheets("Feuil1"), "A1"
    Else
        'optionnel : afficher l'erreur qui a été loggée
        MsgBox "KO: " & db.LastError, vbExclamation
    End If

Bye:
    If ps <> 0 Then DuckVba_Finalize ps
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume Bye

End Sub

Public Sub Demo_Prepared()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, ps As LongPtr, i As Long

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    'Schéma + reset
    db.Exec "DROP TABLE IF EXISTS T;"
    db.Exec "CREATE TABLE IF NOT EXISTS T(isin TEXT, px DOUBLE);"
    db.Exec "DELETE FROM T;"

    'Prepared INSERT
    ps = db.Prepare("INSERT INTO T VALUES (?, ?)")
    For i = 1 To 3
        db.PS_BindText ps, 1, "FR0000" & Format$(i, "000000")
        db.PS_BindDouble ps, 2, 100 + i          ' px est DOUBLE
        db.PS_Exec ps
    Next
    db.PS_CloseDuckDb ps

    'Lecture + affichage
    v = db.QueryFast("SELECT * FROM T ORDER BY isin")
    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(v, 1), UBound(v, 2)).Value = v
    End With

    db.CloseDuckDb
    MsgBox "OK Prepared (A)"
    Exit Sub

Fail:
    On Error Resume Next
    If ps <> 0 Then db.PS_CloseDuckDb ps
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub

Public Sub Test_Prepared_Insert()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, dbPath As String, ps As LongPtr, i As Long, useMemory As Boolean
    
    useMemory = False

    db.Init ThisWorkbook.Path
    If useMemory Then
        dbPath = ":memory:"                             'Option A (RAM)
    Else
        dbPath = ThisWorkbook.Path & "\cache.duckdb"    'Option B (fichier)
    End If
    db.OpenDuckDb dbPath

    'Schéma (propre à chaque run)
    db.Exec "DROP TABLE IF EXISTS T;"
    db.Exec "CREATE TABLE T(isin TEXT, px DOUBLE);"

    'Prepared INSERT
    ps = db.Prepare("INSERT INTO T(isin, px) VALUES (?, ?);")
    For i = 1 To 3
        db.PS_BindText ps, 1, "FR0000" & Format$(i, "000000")
        db.PS_BindDouble ps, 2, 100 + i
        db.PS_Exec ps
    Next
    db.PS_CloseDuckDb ps

    'Lecture
    v = db.QueryFast("SELECT * FROM T ORDER BY isin;")

    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(v, 1), UBound(v, 2)).Value = v
        .Columns("A:B").AutoFit
    End With

    db.CloseDuckDb
    MsgBox "OK Prepared/Insert/Select (A)", vbInformation
    Exit Sub
Fail:
    On Error Resume Next
    If ps <> 0 Then db.PS_CloseDuckDb ps
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub

'=== Zéro fichier intermédiaire : Variant(2D) -> DuckDB (classe cDuck) ===
Public Sub Test_AppendArrayV_Order()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, a As Variant, dbPath As String, useMemory As Boolean
    
    useMemory = False

    '1) Init + ouverture
    db.Init ThisWorkbook.Path
    dbPath = IIf(useMemory, ":memory:", ThisWorkbook.Path & "\cache.duckdb")
    db.OpenDuckDb dbPath

    '2) (Re)création de la table cible
    db.Exec "DROP TABLE IF EXISTS T;"
    db.Exec "CREATE TABLE T(ISIN TEXT, Prix DOUBLE, ModifiedAt TIMESTAMP);"

    '3) Prépare un Variant(2D) avec entêtes en ligne 1
    ReDim v(1 To 3, 1 To 3)
    v(1, 1) = "ISIN":     v(1, 2) = "Prix": v(1, 3) = "ModifiedAt"
    v(2, 1) = "FR0001":   v(2, 2) = 103.1:  v(2, 3) = DateSerial(2025, 9, 7) + TimeSerial(9, 43, 0)
    v(3, 1) = "FR0002":   v(3, 2) = CDbl(999): v(3, 3) = DateSerial(2025, 9, 7) + TimeSerial(10, 1, 0)

    '4) Append direct (hasHeader:=True)
    db.AppendArray "T", v, True

    '5) Lecture + affichage
    a = db.QueryFast("SELECT ISIN, Prix, strftime(ModifiedAt, '%Y-%m-%d %H:%M:%S') AS ModifiedAt FROM T ORDER BY ISIN;")
    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    End With

    db.CloseDuckDb
    MsgBox "OK si colonnes = ISIN / Prix / ModifiedAt (A)", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Sub

