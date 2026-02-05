Attribute VB_Name = "Mod2DuckDb_DictRow1D"
Option Explicit

'------------------------------------------------------------------------------
' SelectToDictRow1D / FillDictVals
'
' But :
'   Remplit un Dictionary à partir d’un SELECT :
'       dict(key) = Variant(1D) contenant UNIQUEMENT les valeurs des colonnes choisies.
'
' Signature logique
'   - keyCol        : colonne utilisée comme clé du Dictionary
'   - valueColsCsv  : liste CSV des colonnes à mettre dans le Variant(1D)
'                     (ex: "px,name,ccy"). Si vide => toutes les colonnes sauf keyCol.
'   - onDupMode     : 0 ignore doublon, 1 remplace.
'
' PLUS RAPIDE que SelectToDictRow2D :
'   SelectToDictRow2D (ex FillDict / SelectToDictW) construit pour CHAQUE ligne :
'     - un SAFEARRAY 2D (2 x N) + allocations BSTR des noms de colonnes (labels)
'     - => beaucoup d’allocations COM (BSTR), plus de copies mémoire, plus de pression GC/heap.
'
'   Ici (Row1D) on ne stocke PAS les labels par ligne :
'     - un simple Variant(1D) de valeurs par clé
'     - => moins d’allocations, moins de BSTR, moins de trafic mémoire
'     - => gain très visible quand il y a beaucoup de lignes (et/ou plusieurs colonnes).
'
' Trade-off :
'   - Row1D est plus rapide et léger, MAIS tu récupères les valeurs "par position"
'     (ex: vals(0)=px, vals(1)=name, etc). Si tu veux accès par nom de colonne,
'     utilise Row2D, ou crée une map "colName -> index" UNE SEULE FOIS côté VBA.
'
' Exemple :
'   Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
'   db.SelectToDictRow1D "SELECT isin, px, name FROM T", "isin", d, "px,name"
'   vals = d("FR0001")   ' vals(0)=px, vals(1)=name
'------------------------------------------------------------------------------

Public Sub TestDict_SelectToDictRow1D()

    On Error GoTo Fail

    Dim db As New cDuck, d As New Dictionary, arr As Variant, j As Long

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    ' Données de démo
    db.Exec "CREATE TABLE T(k TEXT, Name TEXT, ISIN TEXT, a INT, b TEXT);"
    db.Exec "INSERT INTO T VALUES" & _
            "('A','Alpha','FR000A',10,'x')," & _
            "('B','Beta' ,'FR000B',20,'y');"

    ' ------------------------------------------------------------------
    ' 1) DEFAULT : toutes colonnes sauf la clé -> SAFEARRAY 1D
    ' ------------------------------------------------------------------
    db.SelectToDictRow1D "SELECT * FROM T", "k", d, vbNullString, True, 1

    arr = d("B")
    Debug.Print "B (default, all cols except key):"
    For j = LBound(arr) To UBound(arr)
        Debug.Print "  [" & j & "] = "; arr(j)
    Next

    ' ------------------------------------------------------------------
    ' 2) SELECTION : seulement Name + ISIN -> SAFEARRAY 1D
    ' ------------------------------------------------------------------
    d.RemoveAll
    db.SelectToDictRow1D "SELECT * FROM T", "k", d, "Name,ISIN", True, 1

    arr = d("B")
    Debug.Print "B (selected cols: Name, ISIN):"
    For j = LBound(arr) To UBound(arr)
        Debug.Print "  [" & j & "] = "; arr(j)
    Next

Clean:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & "Duck: " & Native_LastErrorText(), vbExclamation
    Resume Clean

End Sub

'------------------------------------------------------------------------------
' Idée :
'   - L’agrégation est faite par SQL (GROUP BY + SUM/MIN/MAX/COUNT/AVG…)
'   - SelectToDictRow1D charge le résultat dans un Dictionary :
'       dict(key) = Variant(1D) contenant uniquement les colonnes agrégées
'
' Exemple : statistiques par ISIN
'
'   sql = "SELECT isin," & _
'         "       SUM(price) AS sum_price," & _
'         "       MIN(price) AS min_price," & _
'         "       MAX(price) AS max_price" & _
'         "  FROM trades" & _
'         " GROUP BY isin"
'
'   db.SelectToDictRow1D sql, "isin", d, "sum_price,min_price,max_price", True, 1
'
' Résultat :
'   vals = d("FR000A")    ' vals(0)=sum_price, vals(1)=min_price, vals(2)=max_price
'
' Notes / bonnes pratiques :
'   - Une seule colonne demandée (ex: "sum_price") => array 1D de taille 1 : vals(0)
'   - Pour éviter l’accès “par position”, crée une table d’index UNE FOIS :
'         idx("sum_price")=0 : idx("min_price")=1 : idx("max_price")=2
'     puis réutilise : vals(idx("min_price"))
'
' Quand une fonction DLL dédiée devient vraiment utile :
'   - groupRows / list : retourner toutes les lignes d’un groupe (pack 2D par clé)
'   - nested dict : construire directement desk -> isin -> values (sans surcoût VBA)
'   - topN per key : top N lignes par clé, tri/filtre côté natif (perf + mémoire)
'------------------------------------------------------------------------------

Public Sub TestDict_SelectToDictRow1D_GroupByStats()

    On Error GoTo Fail

    Dim db As New cDuck, d As New Dictionary, arr As Variant, sql As String, j As Long

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'Données de démo
    db.Exec "CREATE TABLE trades(isin TEXT, price DOUBLE);"
    db.Exec "INSERT INTO trades VALUES " & _
            "('FR000A', 10.0)," & _
            "('FR000A', 12.5)," & _
            "('FR000A',  9.5)," & _
            "('FR000B', 20.0)," & _
            "('FR000B', 22.0);"

    'GROUP BY en SQL + stockage 1D dans le dict
    sql = "SELECT isin," & _
          "       SUM(price) AS sum_price," & _
          "       MIN(price) AS min_price," & _
          "       MAX(price) AS max_price" & _
          "  FROM trades" & _
          " GROUP BY isin;"

    d.RemoveAll
    db.SelectToDictRow1D sql, "isin", d, "sum_price,min_price,max_price", True, 1

    ' Vérification / affichage
    Debug.Print "FR000A (sum, min, max):"
    arr = d("FR000A")
    For j = LBound(arr) To UBound(arr)
        Debug.Print "  [" & j & "] = "; arr(j)
    Next

    Debug.Print "FR000B (sum, min, max):"
    arr = d("FR000B")
    For j = LBound(arr) To UBound(arr)
        Debug.Print "  [" & j & "] = "; arr(j)
    Next

Clean:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & "Duck: " & Native_LastErrorText(), vbExclamation
    Resume Clean

End Sub
