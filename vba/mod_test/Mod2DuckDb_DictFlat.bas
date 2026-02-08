Attribute VB_Name = "Mod2DuckDb_DictFlat"
Option Explicit

'===============================================================================
' SelectToDictFlat
'   Exécute un SELECT et retourne un Scripting.Dictionary "plat" : Key -> Value.
'
'   Cas d’usage typiques (finance):
'     - ISIN -> Ticker / MIC
'     - InstrumentId -> LastPrice
'     - Date -> Rate / DF
'     - Code -> Libellé
'
'   Attendus sur le SELECT :
'     - Doit retourner 2 colonnes (ou plus, mais on n’en utilise qu’une valeur).
'     - keyCol = nom de la colonne clé (obligatoire).
'     - valCol :
'         * si renseigné => valeur prise dans cette colonne
'         * si ""        => auto : si le SELECT a exactement 2 colonnes, la DLL prend
'                           l’autre colonne comme valeur.
'
'   Paramètres :
'     - clearFirst : True => dict.RemoveAll avant remplissage
'     - onDupMode  : gestion doublons (selon ta DLL) :
'         0 = ignore les doublons
'         1 = remplace la valeur existante
'
'   Retour :
'     - Un Dictionary (late-binding) contenant toutes les paires clé/valeur.
'     - En cas d’erreur : la fonction log via HandleError (selon ErrorMode) et
'       retourne quand même un Dictionary (souvent vide) sans bloquer l’exécution.
'===============================================================================

'==================================================================================
' Exemple : dictionnaire Clé / Valeur de table DuckDb
'===================================================================================
Public Sub EX_ISIN_ToName_Dict()
    Dim db As cDuck, d As Object, k As Variant
    Set db = CurrentDuckDb
    db.OpenDuckDb ThisWorkbook.Path & "\market.duckdb"
    On Error GoTo FINALLY

    ' --- Jeu d’essai ---
    db.Exec "DROP TABLE IF EXISTS securities;"
    db.Exec "CREATE TABLE securities(isin TEXT PRIMARY KEY, name TEXT, sector TEXT);"
    db.Exec "INSERT INTO securities VALUES " & _
            "('FR0000133308','LVMH Moët Hennessy Louis Vuitton','Consumer Luxury')," & _
            "('US0378331005','Apple Inc','Technology')," & _
            "('US5949181045','Microsoft Corporation','Technology');"

    ' --- Création du dictionnaire clé=ISIN, valeur=nom ---
    Set d = db.SelectToDictFlat("SELECT isin, name FROM securities;", "isin", "name")

    ' --- Affichage ---
    Debug.Print "Nombre d’ISIN :"; d.Count
    For Each k In d.keys
        Debug.Print k, d(k)
    Next k

FINALLY:
    db.CloseDuckDb
End Sub

'===================================================================================
' Exemple : dictionnaire Clé / Valeur de table DuckDb
'===================================================================================
Public Sub EX_ISIN_ToLastClose_Dict()

    Dim db As cDuck, d As Object, k As Variant
    Set db = CurrentDuckDb
    db.OpenDuckDb ThisWorkbook.Path & "\market.duckdb"
    On Error GoTo FINALLY

    ' --- Jeu d’essai ---
    db.Exec "DROP TABLE IF EXISTS quotes;"
    db.Exec "CREATE TABLE quotes(isin TEXT, trade_date DATE, close DOUBLE);"
    db.Exec "INSERT INTO quotes VALUES " & _
            "('US0378331005','2025-10-31',230.5)," & _
            "('US0378331005','2025-11-04',231.8)," & _
            "('US5949181045','2025-11-01',395.0)," & _
            "('US5949181045','2025-11-04',399.3);"

    ' --- Création du dictionnaire ---
    Set d = db.SelectToDictFlat( _
              "SELECT isin, close FROM quotes q " & _
              "WHERE trade_date=(SELECT max(trade_date) FROM quotes qq WHERE qq.isin=q.isin)", _
              "isin", "close")

    ' --- Affichage ---
    Debug.Print "ISIN  ?  Dernier cours"
    For Each k In d.keys
        Debug.Print k, d(k)
    Next k

FINALLY:
    db.CloseDuckDb
End Sub

'===================================================================================
' Exemple : dictionnaire Clé /  plusieurs Valeurs avec concat de table DuckDb
'===================================================================================
Public Sub EX_ISIN_ToInfo_Dict()

    Dim db As cDuck, d As Object, k As Variant
    Set db = CurrentDuckDb
    db.OpenDuckDb ThisWorkbook.Path & "\market.duckdb"
    On Error GoTo FINALLY

    db.Exec "DROP TABLE IF EXISTS securities;"
    db.Exec "CREATE TABLE securities(isin TEXT PRIMARY KEY, name TEXT, sector TEXT, country TEXT);"
    db.Exec "INSERT INTO securities VALUES " & _
            "('FR0000133308','LVMH','Luxury','FR')," & _
            "('US0378331005','Apple','Tech','US')," & _
            "('JP0000000001','Toyota','Auto','JP');"

    ' Exemple : concaténer les champs pour avoir un mapping ISIN -> "Nom | Secteur | Pays"
    Set d = db.SelectToDictFlat( _
              "SELECT isin, name || ' | ' || sector || ' | ' || country AS info FROM securities;", _
              "isin", "info")

    Debug.Print "ISIN  ?  Infos"
    For Each k In d.keys
        Debug.Print k, d(k)
    Next k

FINALLY:
    db.CloseDuckDb
End Sub

'===================================================================================
' Exemple : dictionnaire Clé /  plusieurs Valeurs avec concat de table DuckDb
'===================================================================================
Public Sub EX_ISIN_Unique_PickDatePriceVol_OneDict1()

    Dim db As cDuck, d As Object, k As Variant
    Set db = CurrentDuckDb
    db.OpenDuckDb ThisWorkbook.Path & "\market.duckdb"
    On Error GoTo FINALLY

    ' 1) Schéma exact (on repart propre)
    db.Exec "DROP TABLE IF EXISTS px_eod;"
    db.Exec "DROP TABLE IF EXISTS px_rt;"
    db.Exec "DROP TABLE IF EXISTS px_vendor3;"

    db.Exec "CREATE TABLE px_eod(" & _
            "  isin TEXT PRIMARY KEY," & _
            "  trade_date DATE," & _
            "  close DOUBLE," & _
            "  volume BIGINT" & _
            ");"

    db.Exec "CREATE TABLE px_rt(" & _
            "  isin TEXT PRIMARY KEY," & _
            "  ts TIMESTAMP," & _
            "  last DOUBLE," & _
            "  vol BIGINT" & _
            ");"

    db.Exec "CREATE TABLE px_vendor3(" & _
            "  isin TEXT PRIMARY KEY," & _
            "  d DATE," & _
            "  p DOUBLE," & _
            "  v BIGINT" & _
            ");"

    ' 2) Données d’exemple
    db.Exec "INSERT INTO px_eod VALUES" & _
            "('FR0000133308','2024-12-31', 818.10,  950000)," & _
            "('US0378331005','2024-12-31', 194.60, 1250000);"

    db.Exec "INSERT INTO px_rt VALUES" & _
            "('FR0000133308','2024-12-31 16:59:59', 819.40, 980000);"

    db.Exec "INSERT INTO px_vendor3 VALUES" & _
            "('JP0000000001','2024-12-31', 2426.00, 510000);"

    ' 3) Requête (priorité RT > EOD > vendor3) ? pack "date|prix|volume"
    Dim q As String, qPacked As String
    q = _
      "WITH " & _
      "e AS (SELECT isin, trade_date AS d_e, close AS p_e,  volume AS v_e FROM px_eod)," & _
      "r AS (SELECT isin, ts::DATE   AS d_r, last  AS p_r,  vol    AS v_r FROM px_rt)," & _
      "v AS (SELECT isin, d          AS d_v, p     AS p_v,  v      AS v_v FROM px_vendor3) " & _
      "SELECT " & _
      "  coalesce(r.isin, e.isin, v.isin) AS isin," & _
      "  coalesce(d_r, d_e, d_v)          AS d," & _
      "  coalesce(p_r, p_e, p_v)          AS p," & _
      "  coalesce(v_r, v_e, v_v)          AS vol " & _
      "FROM r FULL OUTER JOIN e USING(isin) FULL OUTER JOIN v USING(isin)"

    qPacked = _
      "WITH base AS (" & q & ") " & _
      "SELECT isin, " & _
      "       coalesce(CAST(d   AS TEXT),'') || '|' || " & _
      "       coalesce(CAST(p   AS TEXT),'') || '|' || " & _
      "       coalesce(CAST(vol AS TEXT),'') AS packed " & _
      "FROM base;"

    Set d = db.SelectToDictFlat(qPacked, "isin", "packed")

    Debug.Print "ISIN            -> date | prix | volume"
    For Each k In d.keys
        Debug.Print k, d(k)
    Next k

FINALLY:
    db.CloseDuckDb
End Sub


'===================================================================================
' Exemple : dictionnaire Clé /  plusieurs Valeurs avec concat de table DuckDb
'===================================================================================
Public Sub EX_ISIN_Unique_PickDatePriceVol_OneDict2()

    Dim db As cDuck, d As Object, k As Variant, q As String, qPacked As String
    
    Set db = CurrentDuckDb
    db.OpenDuckDb ThisWorkbook.Path & "\market.duckdb"
    On Error GoTo FINALLY

    ' -- schéma & data (identiques à ta version) -------------------
    db.Exec "DROP TABLE IF EXISTS px_eod;"
    db.Exec "DROP TABLE IF EXISTS px_rt;"
    db.Exec "DROP TABLE IF EXISTS px_vendor3;"

    db.Exec "CREATE TABLE px_eod(" & _
            "  isin TEXT PRIMARY KEY," & _
            "  trade_date DATE," & _
            "  close DOUBLE," & _
            "  volume BIGINT" & _
            ");"

    db.Exec "CREATE TABLE px_rt(" & _
            "  isin TEXT PRIMARY KEY," & _
            "  ts TIMESTAMP," & _
            "  last DOUBLE," & _
            "  vol BIGINT" & _
            ");"

    db.Exec "CREATE TABLE px_vendor3(" & _
            "  isin TEXT PRIMARY KEY," & _
            "  d DATE," & _
            "  p DOUBLE," & _
            "  v BIGINT" & _
            ");"

    db.Exec "INSERT INTO px_eod(isin,trade_date,close,volume) VALUES" & _
            "('FR0000133308','2024-12-31', 818.10,  950000)," & _
            "('US0378331005','2024-12-31', 194.60, 1250000);"

    db.Exec "INSERT INTO px_rt(isin,ts,last,vol) VALUES" & _
            "('FR0000133308','2024-12-31 16:59:59', 819.40, 980000);"

    db.Exec "INSERT INTO px_vendor3(isin,d,p,v) VALUES" & _
            "('JP0000000001','2024-12-31', 2426.00, 510000);"

    ' -- requête rapide : UNION ALL + fenêtre ----------------------
    q = _
      "WITH src AS (" & _
      "  SELECT 1 AS prio, isin, date(ts) AS d, last AS p, vol AS vol FROM px_rt " & _
      "  UNION ALL " & _
      "  SELECT 2 AS prio, isin, trade_date, close, volume FROM px_eod " & _
      "  UNION ALL " & _
      "  SELECT 3 AS prio, isin, d, p, v FROM px_vendor3 " & _
      "), ranked AS (" & _
      "  SELECT *, row_number() OVER (PARTITION BY isin ORDER BY prio) AS rn " & _
      "  FROM src " & _
      ") " & _
      "SELECT isin, d, p, vol FROM ranked WHERE rn = 1"

    qPacked = _
      "WITH base AS (" & q & ") " & _
      "SELECT isin, " & _
      "       coalesce(CAST(d   AS TEXT),'') || '|' || " & _
      "       coalesce(CAST(p   AS TEXT),'') || '|' || " & _
      "       coalesce(CAST(vol AS TEXT),'') AS packed " & _
      "FROM base;"

    Set d = db.SelectToDictFlat(qPacked, "isin", "packed")

    Debug.Print "ISIN            -> date | prix | volume"
    For Each k In d.keys
        Debug.Print k, d(k)
    Next k

FINALLY:
    db.CloseDuckDb
End Sub

'===================================================================================
' ISIN (panier) -> Dictionnaire  "date|prix|volume"
' - Crée une temp list via CreateTempList "__basket" (colonne = v)
' - Ne lit que les ISIN du panier (WHERE isin IN (SELECT v FROM __basket))
' - Priorité des sources : RT (1) > EOD (2) > Vendor3 (3)
' - Emballage texte pour SelectToDictFlat
'===================================================================================
Public Sub EX_ISIN_LastPxVol_WITH_TEMP_LIST()

    Dim db As cDuck, d As Object, k As Variant, keys As Variant, q As String, qPacked As String

    Set db = CurrentDuckDb
    db.OpenDuckDb ThisWorkbook.Path & "\market.duckdb"
    On Error GoTo FINALLY

    'Panier d’ISIN (peut venir d’une plage Excel)
    keys = Array("FR0000133308", "US0378331005", "JP0000000001")

    '1) Temp list (rapide via DLL) -> table __basket(v TEXT)
    db.CreateTempList "__basket", keys, "TEXT"

    '2) Requête rapide filtrée sur le panier
    '    UNION ALL + ROW_NUMBER pour choisir la meilleure source (1 = RT, 2 = EOD, 3 = Vendor3)
    q = _
      "WITH src AS (" & _
      "  SELECT 1 AS prio, isin, date(ts) AS d, last AS p, vol AS vol" & _
      "  FROM px_rt" & _
      "  WHERE isin IN (SELECT v FROM __basket)" & _
      "  UNION ALL " & _
      "  SELECT 2 AS prio, isin, trade_date AS d, close AS p, volume AS vol" & _
      "  FROM px_eod" & _
      "  WHERE isin IN (SELECT v FROM __basket)" & _
      "  UNION ALL " & _
      "  SELECT 3 AS prio, isin, d AS d, p AS p, v AS vol" & _
      "  FROM px_vendor3" & _
      "  WHERE isin IN (SELECT v FROM __basket)" & _
      "), ranked AS (" & _
      "  SELECT *, row_number() OVER (PARTITION BY isin ORDER BY prio) AS rn" & _
      "  FROM src" & _
      ") " & _
      "SELECT isin, d, p, vol FROM ranked WHERE rn = 1"

    '3) Emballer en "date|prix|volume" pour un dict plat
    qPacked = _
      "WITH base AS (" & q & ") " & _
      "SELECT isin," & _
      "       coalesce(CAST(d   AS TEXT),'') || '|' || " & _
      "       coalesce(CAST(p   AS TEXT),'') || '|' || " & _
      "       coalesce(CAST(vol AS TEXT),'') AS packed " & _
      "FROM base;"

    ' 4) Récupération en Dictionary (clé = isin, valeur = "date|prix|volume")
    Set d = db.SelectToDictFlat(qPacked, "isin", "packed")

    ' 5) Consommation simple (Debug.Print)
    Debug.Print "ISIN            -> date | prix | volume"
    For Each k In d.keys
        Debug.Print k, d(k)
    Next k

FINALLY:
    db.CloseDuckDb
End Sub

'Exemple : gestion des doublons onDupMode : 0=ignore / 1=remplace
Public Sub Test_DictFlat_Duplicates()

    Dim db As cDuck, d As Object, k As Variant
    Set db = CurrentDuckDb
    db.OpenDuckDb ThisWorkbook.Path & "\demo.duckdb"

    On Error GoTo FINALLY

    db.Exec "DROP TABLE IF EXISTS dup;"
    db.Exec "CREATE TABLE dup(k INTEGER, v TEXT);"
    db.Exec "INSERT INTO dup VALUES (1,'A'),(1,'A2'),(2,'B');"

    ' (a) ignore les doublons (garde la 1ère valeur)
    Set d = CreateObject("Scripting.Dictionary")
    db.FillDictFlat "SELECT k, v FROM dup ORDER BY k, v;", "k", "v", d, True, 0
    Debug.Print "--- onDupMode=0 (ignore) ---"
    For Each k In d.keys: Debug.Print k, d(k): Next

    ' (b) remplace par la dernière occurrence
    Set d = CreateObject("Scripting.Dictionary")
    db.FillDictFlat "SELECT k, v FROM dup ORDER BY k, v;", "k", "v", d, True, 1
    Debug.Print "--- onDupMode=1 (replace) ---"
    For Each k In d.keys: Debug.Print k, d(k): Next

FINALLY:
    db.CloseDuckDb
    
End Sub


