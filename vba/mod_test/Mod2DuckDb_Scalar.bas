Attribute VB_Name = "Mod2DuckDb_Scalar"
Option Explicit

'===============================================================================
'           TEST : Scalar_* (1 valeur)
'
'Avantage (général) :
'  - Évite de rapatrier un Variant(2D) complet (headers + lignes) juste pour un KPI.
'  - Moins d’allocation mémoire / conversions COM -> Variant, donc souvent plus rapide.
'
'Exemples typiques :
'  - Compter :       nb = db.ScalarLong("SELECT COUNT(*) FROM trades;")
'  - Dernière date : dt = db.ScalarDate("SELECT MAX(ts) FROM ticks;")
'  - Dernier prix :  px = db.ScalarDbl("SELECT last(price) FROM ticks WHERE isin='FR...';")
'  - Un libellé :    nm = db.ScalarText("SELECT name FROM instruments WHERE isin='FR...';")
'  - Sanity check :  ok = (db.ScalarLong("SELECT COUNT(*) FROM t WHERE px IS NULL;") = 0)
'
'Notes :
'  - Si la requête renvoie NULL/Empty, wrappers renvoient une valeur par défaut (0, "", 0# selon le type)
'===============================================================================

Sub Test_Scalar_SuperSimple()

    Dim db As New cDuck, nRows  As Long, lastId As Long, lastNom As String, lastDt As Date, AvgId As Double

    '1) Ouvre DuckDB en mémoire
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    '2) Petite table de test
    db.Exec "CREATE TABLE t(" & _
            "  Id   INTEGER," & _
            "  Nom  TEXT," & _
            "  Dt   DATE" & _
            ");"

    db.Exec "INSERT INTO t VALUES " & _
            "(1, 'Alice', DATE '2024-01-01')," & _
            "(2, 'Bob'  , DATE '2024-02-10');"

    '3) Appels scalaires
    nRows = db.ScalarLong("SELECT COUNT(*) FROM t;")
    Debug.Print "Nb lignes = "; nRows

    lastId = db.ScalarLong("SELECT MAX(Id) FROM t;")
    Debug.Print "Max(Id) = "; lastId

    lastNom = db.ScalarText("SELECT Nom FROM t ORDER BY Id DESC LIMIT 1;")
    Debug.Print "Dernier Nom = "; lastNom

    lastDt = db.ScalarDate("SELECT MAX(Dt) FROM t;")
    Debug.Print "Dernière date = "; IIf(lastDt = 0, "(NULL/Empty)", CStr(lastDt))

    AvgId = db.ScalarDbl("SELECT AVG(Id) FROM t;")
    Debug.Print "SUM(Id) = "; AvgId
    
    '4) Fermeture
    db.CloseDuckDb

End Sub
