Attribute VB_Name = "Mod2DuckDb_RapidFuzz_Ext"
Option Explicit

'===============================================================================
' SearchMarketContains / Fuzzy Search (Market data)
'
' Deux approches :
'   1) CONTAINS SQL (LIKE / lower + %term%) :
'      - rapide, simple, parfait quand le terme est clair.
'
'   2) FUZZY SEARCH (approximation) via l’extension DuckDB "rapidfuzz" :
'      - utile si l’utilisateur fait des fautes ou ne connaît pas l’orthographe.
'      - ex: "airbuss" -> "AIRBUS"
'
' Extension DuckDB : rapidfuzz (community extension)
'   - Docs : https://duckdb.org/community_extensions/extensions/rapidfuzz
'   - Cache extensions (Windows) :
'       %USERPROFILE%\.duckdb\extensions\v1.4.3\windows_amd64
'     (= C:\Users\<username>\.duckdb\extensions\v1.4.3\windows_amd64)
'
' Bonnes pratiques (après upgrade DuckDB) :
'   - Forcer la réinstallation pour éviter un cache d’ancienne version :
'       FORCE INSTALL rapidfuzz FROM community;
'       LOAD rapidfuzz;
'
' Exemple fuzzy typique (ratio) :
'   SELECT *, rapidfuzz_ratio(lower(Name), lower(?)) AS score
'   FROM ImportedCsv
'   WHERE score >= 70
'   ORDER BY score DESC
'   LIMIT 50;
'===============================================================================

' Levenshtein normalisée
' distance d’édition (insertions/suppressions/substitutions) normalisée
Sub Test_RapidFuzz_Simple()

    Dim db As New cDuck
    Dim a As Variant
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    
    'Charger l’extension
    db.LoadExt "rapidfuzz"
    
    'Petit jeu de données
    db.Exec "CREATE TABLE t(nom TEXT);" & _
            "INSERT INTO t VALUES " & _
            "('DuckDB'), ('DukDB'), ('Dock DB'), ('Oracle'), ('Dukddb');"
    
    'Calculer le score de similarité avec 'duckdb'
    a = db.QueryFast("SELECT nom," & _
        "       rapidfuzz_ratio(nom, 'duckdb') AS score " & _
        "FROM t " & _
        "ORDER BY score DESC;")
    
    DebugArray2DTable a, "rapidfuzz_ratio"
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"
    
    db.CloseDuckDb

End Sub

Public Sub Test_RapidFuzz_JaroWinkler_Prefix_Postfix_OSA()

    Dim db As New cDuck, a As Variant, ws As Worksheet, R As Long

    On Error GoTo Fail

    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells.Clear

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'Charger l’extension
    db.LoadExt "rapidfuzz"

    R = 1

    ' =========================================================================
    ' 1) Jaro-Winkler : utile pour noms / chaînes courtes / fautes de frappe
    ' =========================================================================
    db.Exec "CREATE TABLE t_jw(a TEXT, b TEXT);" & _
            "INSERT INTO t_jw VALUES " & _
            "('duck', 'duke')," & _
            "('martha', 'marhta')," & _
            "('dixon', 'dicksonx')," & _
            "('JONATHAN', 'JONNATHAN')," & _
            "('SMITH', 'SMYTH');"

    a = db.QueryFast( _
        "SELECT a, b, " & _
        "  rapidfuzz_jaro_winkler_distance(a,b)             AS dist, " & _
        "  rapidfuzz_jaro_winkler_similarity(a,b)           AS sim, " & _
        "  rapidfuzz_jaro_winkler_normalized_distance(a,b)  AS ndist, " & _
        "  rapidfuzz_jaro_winkler_normalized_similarity(a,b)AS nsim " & _
        "FROM t_jw " & _
        "ORDER BY nsim DESC;" _
    )

    ws.Cells(R, 1).Value = "Jaro-Winkler"
    R = R + 1
    Call ArrayToSheet(a, ws, "A" & R, True)
    R = R + UBound(a, 1) + 2

    ' =========================================================================
    ' 2) Prefix : ne compare que le préfixe
    ' =========================================================================
    db.Exec "CREATE TABLE t_pre(a TEXT, b TEXT);" & _
            "INSERT INTO t_pre VALUES " & _
            "('prefix', 'pretext')," & _
            "('prestation', 'presque')," & _
            "('invoice_2025_01', 'invoice_2025_02')," & _
            "('client_duckdb', 'client_duck');"

    a = db.QueryFast( _
        "SELECT a, b, " & _
        "  rapidfuzz_prefix_distance(a,b)             AS dist, " & _
        "  rapidfuzz_prefix_similarity(a,b)           AS sim, " & _
        "  rapidfuzz_prefix_normalized_distance(a,b)  AS ndist, " & _
        "  rapidfuzz_prefix_normalized_similarity(a,b)AS nsim " & _
        "FROM t_pre " & _
        "ORDER BY nsim DESC;" _
    )

    ws.Cells(R, 1).Value = "Prefix"
    R = R + 1
    Call ArrayToSheet(a, ws, "A" & R, True)
    R = R + UBound(a, 1) + 2

    ' =========================================================================
    ' 3) Postfix : ne compare que le suffixe
    ' =========================================================================
    db.Exec "CREATE TABLE t_post(a TEXT, b TEXT);" & _
            "INSERT INTO t_post VALUES " & _
            "('postfix', 'pretext')," & _
            "('file_2025_final', 'report_final')," & _
            "('backup_001.zip', 'archive_002.zip')," & _
            "('john.smith', 'jane.smith');"

    a = db.QueryFast( _
        "SELECT a, b, " & _
        "  rapidfuzz_postfix_distance(a,b)             AS dist, " & _
        "  rapidfuzz_postfix_similarity(a,b)           AS sim, " & _
        "  rapidfuzz_postfix_normalized_distance(a,b)  AS ndist, " & _
        "  rapidfuzz_postfix_normalized_similarity(a,b)AS nsim " & _
        "FROM t_post " & _
        "ORDER BY nsim DESC;" _
    )

    ws.Cells(R, 1).Value = "Postfix"
    R = R + 1
    Call ArrayToSheet(a, ws, "A" & R, True)
    R = R + UBound(a, 1) + 2

    ' =========================================================================
    ' 4) OSA : comme Levenshtein mais autorise transposition adjacente (1 seule fois)
    ' =========================================================================
    db.Exec "CREATE TABLE t_osa(a TEXT, b TEXT);" & _
            "INSERT INTO t_osa VALUES " & _
            "('abcdef', 'azced')," & _
            "('martha', 'marhta')," & _
            "('converse', 'convesre')," & _
            "('duckdb', 'dukcdb');"

    a = db.QueryFast( _
        "SELECT a, b, " & _
        "  rapidfuzz_osa_distance(a,b)             AS dist, " & _
        "  rapidfuzz_osa_similarity(a,b)           AS sim, " & _
        "  rapidfuzz_osa_normalized_distance(a,b)  AS ndist, " & _
        "  rapidfuzz_osa_normalized_similarity(a,b)AS nsim " & _
        "FROM t_osa " & _
        "ORDER BY nsim DESC;" _
    )

    ws.Cells(R, 1).Value = "OSA (Optimal String Alignment)"
    R = R + 1
    Call ArrayToSheet(a, ws, "A" & R, True)
    R = R + UBound(a, 1) + 2
    
    ' =========================================================
    ' 5) Partial matches (sous-chaînes) : rapidfuzz_partial_ratio
    ' =========================================================
    db.Exec "CREATE TABLE t_partial(texte TEXT, pattern TEXT);" & _
            "INSERT INTO t_partial VALUES " & _
            "('Facture client 2025-01 - DUCKDB', 'duckdb')," & _
            "('Commande #A-7781 : livraison express', 'livraison')," & _
            "('Adresse: 12 rue du Canard, 75000 Paris', 'canard')," & _
            "('ABC-123-XYZ', '123')," & _
            "('Lorem ipsum dolor sit amet', 'dolor');"

    a = db.QueryFast( _
        "SELECT texte, pattern, " & _
        "  rapidfuzz_partial_ratio(texte, pattern) AS partial_score " & _
        "FROM t_partial " & _
        "ORDER BY partial_score DESC;" _
    )
    
    ws.Cells(R, 1).Value = "Partial matches (rapidfuzz_partial_ratio)"
    R = R + 1
    Call ArrayToSheet(a, ws, "A" & R, True)
    R = R + UBound(a, 1) + 2
    
    ws.Columns.AutoFit
    db.CloseDuckDb
    MsgBox "OK : rapidfuzz tests (Jaro-Winkler / Prefix / Postfix / OSA/Partial)", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & db.LastError, vbExclamation

End Sub

Sub TestSearchLikeClassic()

    'Chargez avant sub "Demo_ImportCsv"
        'récupérer exemple data https://live.euronext.com/en/products/equities/list
    
    Dim v As Variant
    v = SearchMarketContains("Paris", 10, True)  ' affiche et renvoie le Variant(2D)
    
End Sub

Public Function SearchMarketContains(ByVal term As String, _
    Optional ByVal limitN As LongLong = 50, Optional ByVal showToSheet As Boolean = True) As Variant

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, arr As Variant
    Static psContains As LongPtr   ' cache local : préparé une seule fois sur ce handle

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\DbDuckDb.duckdb"
    
    If psContains = 0 Then
        psContains = db.Prepare( _
            "SELECT [Name], [ISIN], [Market], [Currency], [last Price], [Volume] " & _
            "FROM [ImportedCsv] " & _
            "WHERE lower([Market]) LIKE lower('%' || ? || '%') " & _
            "ORDER BY [Name];")
    End If

    db.PS_BindText psContains, 1, term
    db.PS_Exec psContains

    arr = db.QueryFast( _
        "SELECT [Name], [ISIN], [Market], [Currency], [last Price], [Volume] " & _
        "FROM [ImportedCsv] " & _
        "WHERE lower([Market]) LIKE lower(" & SqlQ("%" & LCase$(term) & "%") & ") " & _
        "ORDER BY [Name];")

    'Retour de la fonction
    SearchMarketContains = arr

    'Affichage si demandé
    If showToSheet Then
        Set ws = ThisWorkbook.Worksheets(1)
        ws.Cells.Clear
        If Not IsEmpty(arr) Then
            ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
        Else
            ws.Range("A1").Value = "Aucun résultat."
        End If
    End If
    
    db.CloseDuckDb
    Exit Function

Fail:
    On Error Resume Next
    db.CloseDuckDb
    If psContains <> 0 Then db.PS_CloseDuckDb psContains: psContains = 0
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
End Function

