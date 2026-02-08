Attribute VB_Name = "Mod2DuckDb_DictRow2D"
Option Explicit

'===============================================================================
' Remplit un Scripting.Dictionary depuis un SELECT.
'
' Pour chaque ligne du résultat :
'   - la clé du Dictionary = valeur de la colonne `keyCol`
'   - la valeur du Dictionary = un petit tableau 2D (2 x (nCol-1)) contenant :
'         Ligne 1 : les noms des colonnes (hors keyCol)
'         Ligne 2 : les valeurs correspondantes (hors keyCol)
'
' Exemple (SELECT id, bid, ask FROM quotes) avec keyCol="id" :
'   dict(123) = [ ["bid","ask"];
'                 [ 1.01,  1.02] ]
'
' clearFirst : True => dict.RemoveAll avant remplissage
' onDupMode  : 0 = ignore si clé déjà présente ; 1 = remplace la valeur existante
'===============================================================================

Public Sub TestDict_ISIN_50k()

    On Error GoTo Fail

    Dim db  As New cDuck, d As New Dictionary, t As New cHiPerfTimer, sql As String
    Dim v   As Variant, k As Variant, arr As Variant, msDict As Double, n As Long, i As Long

    'Nombre de lignes de test (et d'ISIN uniques)
    n = 50000

    '1) Session DuckDB en mémoire
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    '2) Création de la table T avec N lignes
    '    ISIN : FR000001 .. FR050000 (N uniques)
    '    Prix : 100 + (i+1)
    '    ModifiedAt : constante
    sql = _
        "CREATE TABLE T AS " & _
        "SELECT " & _
        "  'FR' || lpad(CAST(i+1 AS VARCHAR), 6, '0') AS ISIN, " & _
        "  100.0 + (i+1) AS Prix, " & _
        "  TIMESTAMP '2025-09-07 09:00:00' AS ModifiedAt " & _
        "FROM range(" & n & ") tbl(i);"

    db.Exec sql

    '3) Afficher un échantillon de la table T dans Feuil1 (contrôle visuel)
    'v = db.QueryFast("SELECT * FROM T ORDER BY ISIN ;")
    'If Not IsEmpty(v) Then
        'ArrayToSheet v, ThisWorkbook.Worksheets("Feuil1"), "A1"
    'End If

    '4) Remplir le Dictionary :
    '    - clé = ISIN (50000 clés uniques)
    '    - valeur = SAFEARRAY(2 x nbColonnes) (noms / valeurs)
    '    - onDupMode=1 => ici pas d'effet, il n'y a pas de doublons
    t.Start
    db.SelectToDictRow2D _
        "SELECT * FROM T ORDER BY ISIN", _
        "ISIN", _
        d, _
        True, _
        1
    msDict = t.StopMilliseconds

    '5) Résultats du bench
    Debug.Print "===== TestDict_ISIN_50k ====="
    Debug.Print "Lignes générées         : "; n
    Debug.Print "Temps SelectToDictRow2D (Dict)   : "; Format$(msDict, "0.000"); " ms"
    Debug.Print "Nombre de clés dans Dict: "; d.Count
    
    
    Dim BoolDp As Boolean
    '6) Afficher un exemple pour 1 ISIN (si dispo)
    If d.Count And BoolDp = False Then
        k = d.keys()(0)   ' première clé
        arr = d(k)

        Debug.Print "----- Exemple pour clé : "; CStr(k)

        ' Cas SelectToDictRow2D : 2 lignes x nbColonnes (noms / valeurs)
        If UBound(arr, 1) = 2 Then
            For i = LBound(arr, 2) To UBound(arr, 2)
                Debug.Print "  "; arr(1, i) & " = "; arr(2, i)
            Next i
        ElseIf UBound(arr, 2) = 2 Then
            ' Si jamais c'est N x 2
            For i = LBound(arr, 1) To UBound(arr, 1)
                Debug.Print "  "; arr(i, 1) & " = "; arr(i, 2)
            Next i
        Else
            Debug.Print "  (shape inattendu pour arr : UB1=" & _
                        UBound(arr, 1) & ", UB2=" & UBound(arr, 2) & ")"
        End If
    End If

Clean:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    Debug.Print "Erreur TestDict_ISIN_50k: "; Err.Description
    Resume Clean

End Sub

Public Sub TestDict_A()

    On Error GoTo Fail

    Dim db As New cDuck, d As New Dictionary, arr As Variant, j As Long
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'Données de démo
    db.Exec "CREATE TABLE T(k TEXT, a INT, b TEXT);"
    db.Exec "INSERT INTO T VALUES('A',10,'x'),('B',20,'y');"

    'Dictionnaire (clé = k ; valeur = SAFEARRAY 2xJ titres/valeurs)
    db.SelectToDictRow2D "SELECT * FROM T", "k", d, True, 1   ' onDupMode=1 : remplace en cas de doublon

    'Test lecture clé "B"
    arr = d("B")
    For j = LBound(arr, 2) To UBound(arr, 2)
        Debug.Print arr(1, j) & " = "; arr(2, j)
    Next

Clean:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume Clean
    
End Sub

Public Sub TestDict_ISIN_A()

    On Error GoTo Fail

    'Session cDuck en mémoire
    Dim db As New cDuck, d As New Dictionary, arr As Variant, k As Variant, j As Long
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'Table + données (doublon sur FR0002)
    db.Exec _
        "CREATE TABLE T (ISIN TEXT, Prix DOUBLE, ModifiedAt TIMESTAMP);" & _
        "INSERT INTO T VALUES " & _
        "('FR0001', 103.10, TIMESTAMP '2025-09-07 09:43:00')," & _
        "('FR0002', 999.00,  TIMESTAMP '2025-09-07 10:01:00')," & _
        "('FR0003',  47.50,  TIMESTAMP '2025-09-07 11:15:00')," & _
        "('FR0002', 1001.0,  TIMESTAMP '2025-09-07 12:00:00');"

    'Dictionnaire (clé = ISIN ; valeur = SAFEARRAY(2 x J) titres/valeurs)
    'onDupMode=1 : en cas de doublon de clé, garde la DERNIÈRE ligne rencontrée (FR0002 -> 1001.0)
    db.SelectToDictRow2D "SELECT * FROM T ORDER BY ISIN, ModifiedAt", "ISIN", d, True, 1

    'Tester une clé précise
    If d.Exists("FR0002") Then
        arr = d("FR0002")
        Debug.Print "== FR0002 =="
        For j = LBound(arr, 2) To UBound(arr, 2)
            Debug.Print arr(1, j) & " = "; arr(2, j)
        Next
    End If

    'Boucler sur toutes les clés
    For Each k In d.keys
        arr = d(k)
        Debug.Print "== " & CStr(k) & " =="
        For j = LBound(arr, 2) To UBound(arr, 2)
            Debug.Print arr(1, j) & " = "; arr(2, j)
        Next
    Next

Clean:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub
Fail:
    On Error Resume Next
    If Not db Is Nothing Then db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description, vbExclamation
    
End Sub

