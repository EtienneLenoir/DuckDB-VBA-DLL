Attribute VB_Name = "Mod2DuckDb_CI_Ext"
Option Explicit
Public gUiDuck As cDuck

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'===============================================================================
                            ' CLI DuckDB & Extension UI
' Prérequis :
'   - DuckDB CLI (optionnel, pour tester en ligne de commande) :
'       winget install DuckDB.cli
'       Docs : https://duckdb.org/docs/stable/clients/cli/overview
'       Install : https://duckdb.org/install/?platform=windows&environment=cli
'
'   - Côté VBA :
'       * duckdb.dll + duckdb_vba_bridge.dll accessibles (gUiDuck.Init ThisWorkbook.Path)
'       * Fichier de base : cache.duckdb (ou adapter le chemin)
'       * Extension "ui" disponible : INSTALL ui; LOAD ui;
'
' Fonctions / commandes DuckDB utilisées :
'   - CALL start_ui();          -> lance le serveur UI (par défaut : http://localhost:4213/)
'   - CALL stop_ui_server();    -> stoppe le serveur UI
'   - (à activer si besoin) :
'       INSTALL ui;
'       LOAD ui;
'
' Démos :
'   - DuckUI_Open_KeepAlive : ouvre cache.duckdb en lecture seule et démarre l’UI
'   - DuckUI_Stop           : arrête l’UI et libère la session (Sleep pour laisser le temps)
'
' Notes :
'   - gUiDuck est global pour garder la connexion vivante (sinon l’UI s’arrête).
'   - Sleep n’existe qu’en VBA7 (PtrSafe) : adapter si Excel 32-bit.
'===============================================================================

Sub DuckUI_Open_KeepAlive()

    Set gUiDuck = New cDuck
    gUiDuck.Init ThisWorkbook.Path

    gUiDuck.OpenReadOnly ThisWorkbook.Path & "\cache.duckdb"
    
    '1) Installer (best-effort) puis charger
    On Error Resume Next
        'gUiDuck.Exec "LOAD ui;"
    On Error GoTo 0

    '2) Lancer l'UI
    gUiDuck.Exec "CALL start_ui();"

    '3) Ouvre : ouvre http://localhost:4213/
 
End Sub

Public Sub DuckUI_Stop()

    If Not gUiDuck Is Nothing Then

        On Error Resume Next
            gUiDuck.Exec "CALL stop_ui_server();"
            Sleep 400
            '2) ferme aussi le singleton si tu l’utilises
            If Not m_singleton Is Nothing Then
                m_singleton.CloseDuckDb
                Set m_singleton = Nothing
            End If
        On Error GoTo 0
        gUiDuck.CloseDuckDb
        Set gUiDuck = Nothing
    End If
    
End Sub
