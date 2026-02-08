Attribute VB_Name = "Mod1DuckDb_Parquet_Research"
Option Explicit

'===============================================================================
' Objectif :
'   Rechercher rapidement des lignes dans un fichier Parquet
'   à partir d’une clé  (ex: ISIN), ou liste de clé par méthode join+ table temp
'
' Contexte :
'   - DuckDB sait lire un Parquet directement via read_parquet(...).
'   - L’enjeu performance vient surtout de “combien de fois” on scanne le Parquet :
'       * 1 clé  => idéalement 1 requête (1 scan du Parquet)
'       * N clés => surtout PAS N requêtes (sinon N scans du Parquet)
'
' Fonctions / démos incluses :
'   1) Test_ParquetRowByKey
'      - Cas simple : une seule valeur de clé (1 ISIN)
'      - Appelle ParquetRowByKey(parquetPath, keyCol, keyValue)
'      - Retour attendu : tableau 2D (headers + 1 ligne)
'
'   2) Test_ParquetRowsByKeyDict :
'      - Cas perf : une liste de clés dans un Dictionary (N ISIN)
'      - Appelle ParquetRowsByKeyDict(parquetPath, keyCol, dictKeys, keepOrder)
'      - Stratégie : on injecte la liste des clés dans une table TEMP DuckDB
'        (FrameFromValue), puis on fait un JOIN sur read_parquet(...).
'      - Résultat : 1 seule requête, 1 seul scan du Parquet (beaucoup plus rapide).
'===============================================================================

Public Sub Test_ParquetRowByKey()

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, parquetName As String, parquetWin As String, parquetDuck As String, keyCol As String, isin As String

    parquetName = "TestParquetSearch.parquet"   ' <- change ici si besoin
    parquetWin = ThisWorkbook.Path & "\" & parquetName
    parquetDuck = Replace(parquetWin, "\", "/")

    keyCol = "ISIN"
    isin = "MDJ0875X95NQ"

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    Call db.TryLoadExt("parquet")

    ' --- Création du parquet si absent ---
    Ensure_TestParquet db, parquetWin

    ' --- Lecture ciblée ---
    a = ParquetRowByKey(db, parquetDuck, keyCol, isin)

    If IsEmpty(a) Then
        MsgBox "Aucune ligne trouvée pour " & keyCol & "=" & isin, vbInformation
        GoTo CleanExit
    End If

    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"
    MsgBox "OK : ligne trouvée et affichée.", vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur Test_ParquetRowByKey: " & Err.Description & _
           IIf(Len(Native_LastErrorText()) > 0, vbCrLf & Native_LastErrorText(), ""), vbExclamation
    Resume CleanExit

End Sub
Private Sub Ensure_TestParquet(db As cDuck, parquetWin As String)

    Dim parquetDuck As String
    parquetDuck = Replace(parquetWin, "\", "/")

    ' Si le fichier existe déjà ? rien à faire
    If Dir$(parquetWin, vbNormal) <> "" Then Exit Sub

    ' Table de démo (avec ISIN)
    db.Exec "CREATE OR REPLACE TABLE demo_parquet AS " & _
            "SELECT * FROM (VALUES " & _
            " ('MDJ0875X95NQ','Instrument A',120.0,NOW())," & _
            " ('FR0000123456','Instrument B', 95.5,NOW())," & _
            " ('US0000000001','Instrument C',180.2,NOW())" & _
            ") AS t(ISIN, Name, Price, ModifiedAt);"

    ' Export Parquet vers le fichier demandé
    db.Exec "COPY (SELECT * FROM demo_parquet) TO " & _
            SqlQ(parquetDuck) & " (FORMAT PARQUET);"

End Sub

'--- Démo : Recherche ISIN à partir d'un dict d’ISIN dans Ficher Parquet (Méthode join + table temp)
Public Sub Test_ParquetRowsByKeyDict()

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, d As Object, pWin As String, p As String

    Set d = CreateObject("Scripting.Dictionary")

    pWin = ThisWorkbook.Path & "\TestParquetSearch.parquet"
    p = Replace(pWin, "\", "/")

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    Call db.TryLoadExt("parquet")

    ' Crée le parquet si absent (fonction helper que tu as déjà/peux ajouter)
    Ensure_TestParquet db, pWin

    ' Exemple liste
    d("MDJ0875X95NQ") = True
    d("FR0000123456") = True
    d("W3J0INZILTBB") = True
    d("UGDIP1CJ2V4M") = True

    a = ParquetRowsByKeyDict(db, p, "ISIN", d, True)

    If IsEmpty(a) Then
        MsgBox "Aucune ligne.", vbInformation
        GoTo CleanExit
    End If

    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit

End Sub

'--- Démo : Lookup Parquet ultra rapide (1..N colonnes) via ParquetRowByKey_SelectCols
Public Sub Test_ParquetRowByKey_SelectCols()

    On Error GoTo Fail

    Dim db As New cDuck, price As Variant, arr As Variant, parquetName As String, parquetWin As String, parquetDuck As String, keyCol As String, isin As String
    
    parquetName = "TestParquetSearch.parquet"
    parquetWin = ThisWorkbook.Path & "\" & parquetName
    parquetDuck = Replace(parquetWin, "\", "/")

    keyCol = "ISIN"
    isin = "MDJ0875X95NQ"

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    Call db.TryLoadExt("parquet")

    ' --- Création du parquet si absent ---
    Ensure_TestParquet db, parquetWin

    ' --- 1 colonne => scalaire (ultra pratique pour du lookup VBA)
    price = ParquetRowByKey_SelectCols(db, parquetDuck, keyCol, isin, "Price")

    ' --- N colonnes => Variant(2D) (headers + 1 ligne)
    arr = ParquetRowByKey_SelectCols(db, parquetDuck, keyCol, isin, "ISIN", "Name", "ModifiedAt")

    With ThisWorkbook.Worksheets(1)
        ArrayToSheet arr, ThisWorkbook.Worksheets(1), "A1"
        .Columns.AutoFit
    End With

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur Test_ParquetRowByKey_SelectCols: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit

End Sub

'==============================================================================
' Test_ParquetReadFiltersToArray
' Démo:
'   - Ouvre DuckDB en mémoire (:memory:)
'   - Charge l'extension parquet (best-effort)
'   - Crée un Parquet de test si absent (Ensure_TestParquet)
'   - Applique plusieurs filtres (ParamArray) + ORDER BY
'   - Affiche le résultat dans Excel
'==============================================================================
Public Sub Test_ParquetReadFiltersToArray()

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, parquetWin As String, parquetDuck As String

    parquetWin = ThisWorkbook.Path & "\TestParquetSearch.parquet"
    parquetDuck = Replace(parquetWin, "\", "/")

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    Call db.TryLoadExt("parquet")

    ' Crée le parquet si absent (helper de ton module)
    Ensure_TestParquet db, parquetWin

    'N filtres -> AND ( ... ) + ORDER BY
    a = ParquetReadFiltersToArray(db, parquetDuck, "Price DESC", _
        "Price > 100", _
        "ISIN LIKE 'FR%' OR ISIN LIKE 'MDJ%'" _
    )

    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"
    MsgBox "OK — filtres Parquet appliqués (aperçu affiché)", vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur Test_ParquetReadFiltersToArray: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit

End Sub

