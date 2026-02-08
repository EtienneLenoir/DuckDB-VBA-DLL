VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHiPerfTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    ' Office 64/32 bits récents
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
        ByRef lpPerformanceCount As LongLong) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
        ByRef lpFrequency As LongLong) As Long

#End If

Private m_useQPC        As Boolean   'True = on utilise QPC, False = fallback Timer
Private m_freq          As Double    'fréquence des ticks QPC
Private m_startCounter  As Double    'valeur de départ (ticks ou Timer)
Private m_lastSeconds   As Double    'dernier Stop() en secondes

'----------------------------------------------
' Init : détecte si QueryPerformanceCounter est dispo
'----------------------------------------------
Private Sub Class_Initialize()
#If VBA7 Then
    Dim f As LongLong
    If QueryPerformanceFrequency(f) <> 0 And f > 0 Then
        m_useQPC = True
        m_freq = CDbl(f)
    Else
        m_useQPC = False
    End If
#Else
    Dim fCur As Currency
    If QueryPerformanceFrequency(fCur) <> 0 Then
        m_useQPC = True
        m_freq = CDbl(fCur)
    Else
        m_useQPC = False
    End If
#End If
End Sub

'----------------------------------------------
' Démarre le chrono
'----------------------------------------------
Public Sub Start()
    m_lastSeconds = 0#
    If m_useQPC Then
        ' Haute précision
    #If VBA7 Then
        Dim C As LongLong
        QueryPerformanceCounter C
        m_startCounter = CDbl(C)
    #End If
    Else
        ' Fallback : Timer (secondes depuis minuit)
        m_startCounter = VBA.Timer
    End If
End Sub

'----------------------------------------------
' Temps écoulé en secondes (ne reset pas le chrono)
'----------------------------------------------
Public Function ElapsedSeconds() As Double
    If m_useQPC Then
    #If VBA7 Then
        Dim C As LongLong
        QueryPerformanceCounter C
        ElapsedSeconds = (CDbl(C) - m_startCounter) / m_freq
    #End If
    Else
        Dim t As Double
        t = VBA.Timer
        ' Gestion du passage de minuit : Timer repart de 0
        If t < m_startCounter Then
            t = t + 86400#
        End If
        ElapsedSeconds = t - m_startCounter
    End If
End Function

'----------------------------------------------
' Temps écoulé en millisecondes
'----------------------------------------------
Public Function ElapsedMilliseconds() As Double
    ElapsedMilliseconds = ElapsedSeconds() * 1000#
End Function

'----------------------------------------------
' Stop : retourne le temps en secondes et mémorise le résultat
'----------------------------------------------
Public Function StopSeconds() As Double
    m_lastSeconds = ElapsedSeconds()
    StopSeconds = m_lastSeconds
End Function

'----------------------------------------------
' Stop en millisecondes
'----------------------------------------------
Public Function StopMilliseconds() As Double
    StopMilliseconds = StopSeconds() * 1000#
End Function

'----------------------------------------------
' Dernier résultat (en secondes) après Stop*
'----------------------------------------------
Public Property Get LastSeconds() As Double
    LastSeconds = m_lastSeconds
End Property


