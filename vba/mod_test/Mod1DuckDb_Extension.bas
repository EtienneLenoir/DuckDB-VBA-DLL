Attribute VB_Name = "Mod1DuckDb_Extension"
Option Explicit

'===============================================================================
'  DuckDB Extensions : https://duckdb.org/docs/stable/extensions/overview
'
'  DuckDB télécharge les extensions (INSTALL) dans un dossier cache utilisateur.
'  Sur Windows, le chemin par défaut ressemble à :
'
'     %USERPROFILE%\.duckdb\extensions\v1.4.3\windows_amd64\
'     (= C:\Users\<username>\.duckdb\extensions\v1.4.3\windows_amd64\ )
'
'  - %USERPROFILE% est une variable Windows (ex: C:\Users\Etien)
'  - v1.4.3 dépend de la version DuckDB utilisée
'  - windows_amd64 = Windows 64-bit
'
'  IMPORTANT (environnement pro / réseau verrouillé) :
'  - Si le poste n’a PAS accès internet / proxy / droits, les commandes :
'        INSTALL <ext> FROM community;
'        LOAD <ext>;
'    peuvent échouer.
'
'  Solution offline :
'    1) Sur une machine autorisée, exécuter une fois :
'         FORCE INSTALL rapidfuzz FROM community;
'         FORCE INSTALL nanodbc  FROM community;
'         FORCE INSTALL miniplot FROM community;
'         ...
'       => cela télécharge les fichiers *.duckdb_extension dans le cache.
'
'    2) Copier manuellement les fichiers *.duckdb_extension téléchargés depuis :
'         %USERPROFILE%\.duckdb\extensions\v1.4.3\windows_amd64\
'       vers le même chemin sur les postes “bloqués”.
'
'    3) Ensuite sur le poste offline, tu peux faire uniquement :
'         LOAD rapidfuzz;
'         LOAD nanodbc;
'         ...
'       (plus besoin de INSTALL si les fichiers sont déjà présents au bon endroit)
'
'  CONSEIL :
'  - Pour éviter les conflits entre versions (1.3.x -> 1.4.x), garde un dossier
'    d’extensions par version. Après upgrade DuckDB, utilise :
'       FORCE INSTALL <ext> ...
'    pour re-télécharger les binaires compatibles avec la nouvelle version.
'===============================================================================

'Procédure globale : appelle ça au début de tes demos (ou dans db.OpenDuckDb si tu veux)
Public Sub DuckEnsureDefaultExtensions()

    Dim db As New cDuck, forceInstall As Boolean

    On Error GoTo Fail

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb ":memory:"

    forceInstall = True
     
    'Community
    'Call EnsureExt(db, "nanodbc", True, forceInstall, True)
    'Call EnsureExt(db, "rapidfuzz", True, forceInstall, True)
    'Call EnsureExt(db, "miniplot", True, forceInstall, True)
    'Call EnsureExt(db, "stochastic", True, forceInstall, True)

    ' Core/officielles (à activer seulement si tu les utilises)
    Call EnsureExt(db, "postgres_scanner", False, forceInstall, True)
    'Call EnsureExt(db, "json", False, forceInstall, True)
    'Call EnsureExt(db, "parquet", False, forceInstall, True)

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

'Installe + charge une extension (community ou core)
Private Sub EnsureExt(db As cDuck, ByVal extName As String, _
    Optional ByVal isCommunity As Boolean = True, Optional ByVal forceInstall As Boolean = True, Optional ByVal doLoad As Boolean = True)

    Dim sql As String

    ' Important: DuckDB prend ce directory pour installer/charger les extensions
    'db.Exec "SET extension_directory = " & SqlQ(DuckExtDir()) & ";"

    If forceInstall Then
        If isCommunity Then
            sql = "FORCE INSTALL " & extName & " FROM community;"
        Else
            sql = "FORCE INSTALL " & extName & ";"
        End If
    Else
        If isCommunity Then
            sql = "INSTALL " & extName & " FROM community;"
        Else
            sql = "INSTALL " & extName & ";"
        End If
    End If

    db.Exec sql

    If doLoad Then
        db.Exec "LOAD " & extName & ";"
    End If
End Sub

'Configuration : où mettre les extensions
Private Function DuckExtDir() As String
    'Option 1: par utilisateur (cache DuckDB)
    'DuckExtDir = Environ$("USERPROFILE") & "\.duckdb\extensions"

    'Option 2: dossier dédié (recommandé pour éviter les conflits de versions)
    DuckExtDir = Environ$("TEMP") & "\duck_ext"
End Function


