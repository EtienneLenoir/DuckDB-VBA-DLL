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

'--- Démo :  Get row parquet : récupère la ligne ISIN dans Fichier Parquet
Public Sub Test_ParquetRowByKey()

    On Error GoTo Fail

    Dim db As New cDuck, a As Variant, parquetPath As String, keyCol As String, isin As String

    parquetPath = ThisWorkbook.Path & "\access_table.parquet"
    keyCol = "ISIN"
    isin = "MDJ0875X95NQ"
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb parquetPath

    a = ParquetRowByKey(db, parquetPath, keyCol, isin)

    If IsEmpty(a) Then
        MsgBox "Aucune ligne trouvée pour " & keyCol & "=" & isin, vbInformation
        Exit Sub
    End If

    'Affichage Excel (ligne 1 = headers, ligne 2 = data)
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"
    'ou : ShowArrayOnSheet a
    MsgBox "OK : ligne trouvée et affichée.", vbInformation
    Exit Sub

Fail:
    MsgBox "Erreur Test_ParquetRowByKey_ToExcel: " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

'--- Démo : Recherche ISIN à partir d'un dict d’ISIN dans Ficher Parquet (Méthode join + table temp)
Public Sub Test_ParquetRowsByKeyDict()

    Dim db As New cDuck, a As Variant, d As Object, p As String
    Set d = CreateObject("Scripting.Dictionary")

    p = ThisWorkbook.Path & "\access_table.parquet"
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    
    'Exemple liste
    d("YE434WYRUWNP") = True
    d("Y8T2SU37NNJE") = True
    d("W3J0INZILTBB") = True
    d("UGDIP1CJ2V4M") = True

    a = ParquetRowsByKeyDict(db, p, "ISIN", d, True)

    If IsEmpty(a) Then
        MsgBox "Aucune ligne.", vbInformation
        Exit Sub
    End If
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"
    
End Sub
