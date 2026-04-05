Attribute VB_Name = "modConfig"
Option Explicit

' Configuração base do encaixe.
Public Const MIN_GAP_MM As Double = 0.5
Public Const GRID_STEP_COARSE_MM As Double = 1#
Public Const GRID_STEP_FINE_MM As Double = 0.2
Public Const GRID_STEP_MICRO_MM As Double = 0.1

' Pesos de score para compactação.
Public Const SCORE_WEIGHT_CONTACT As Double = 2.2
Public Const SCORE_WEIGHT_VOID As Double = 0.08
Public Const SCORE_WEIGHT_SPREAD As Double = 0.35
Public Const SCORE_WEIGHT_BOUNDARY As Double = 1.1
Public Const SCORE_WEIGHT_TOUCH As Double = 0.6
