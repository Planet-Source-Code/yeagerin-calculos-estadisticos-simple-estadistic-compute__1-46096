VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ppal 
   AutoRedraw      =   -1  'True
   Caption         =   "Estadística - Juan Belón Pérez, 1º Sistemas"
   ClientHeight    =   6915
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Ppal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2715
      Left            =   0
      TabIndex        =   1
      Top             =   4230
      Width           =   7575
      Begin VB.Frame Frame8 
         Caption         =   "Varianzas marginales"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   2070
         TabIndex        =   28
         Top             =   1410
         Width           =   1935
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   23
            Left            =   840
            TabIndex        =   32
            Top             =   1020
            Width           =   60
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Segundo Parcial"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   22
            Left            =   60
            TabIndex        =   31
            Top             =   780
            Width           =   1140
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            Index           =   3
            X1              =   60
            X2              =   1860
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   3
            X1              =   1890
            X2              =   60
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   21
            Left            =   840
            TabIndex        =   30
            Top             =   480
            Width           =   60
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Primer Parcial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   20
            Left            =   60
            TabIndex        =   29
            Top             =   210
            Width           =   960
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Varianzas condicionadas"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   30
         TabIndex        =   23
         Top             =   1410
         Width           =   2025
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Suspensos"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   19
            Left            =   60
            TabIndex        =   27
            Top             =   1020
            Width           =   765
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Aprobados"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   18
            Left            =   60
            TabIndex        =   26
            Top             =   810
            Width           =   780
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   1920
            X2              =   60
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            Index           =   2
            X1              =   60
            X2              =   1920
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Aprobados"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   17
            Left            =   60
            TabIndex        =   25
            Top             =   480
            Width           =   780
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Suspensos"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   16
            Left            =   60
            TabIndex        =   24
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.CommandButton B_Calcular 
         Caption         =   "&Calcular"
         Default         =   -1  'True
         Height          =   1095
         Left            =   5970
         TabIndex        =   22
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox txt_Adorno 
         Enabled         =   0   'False
         Height          =   1155
         Left            =   5940
         TabIndex        =   21
         Top             =   210
         Width           =   1485
      End
      Begin VB.Frame Frame6 
         Caption         =   "Medias marginales"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   5790
         TabIndex        =   18
         Top             =   1410
         Width           =   1755
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   15
            Left            =   630
            TabIndex        =   34
            Top             =   1020
            Width           =   60
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   14
            Left            =   630
            TabIndex        =   33
            Top             =   450
            Width           =   60
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   1620
            X2              =   30
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Segundo Parcial"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   13
            Left            =   60
            TabIndex        =   20
            Top             =   780
            Width           =   1140
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Primer Parcial:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   12
            Left            =   60
            TabIndex        =   19
            Top             =   240
            Width           =   1020
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            Index           =   1
            X1              =   30
            X2              =   1620
            Y1              =   720
            Y2              =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Medias condicionadas"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   4020
         TabIndex        =   13
         Top             =   1410
         Width           =   1755
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Aprobados"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   17
            Top             =   1020
            Width           =   780
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Suspensos"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   10
            Left            =   60
            TabIndex        =   16
            Top             =   780
            Width           =   765
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   1680
            X2              =   30
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            Index           =   0
            X1              =   30
            X2              =   1680
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Aprobados"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   15
            Top             =   480
            Width           =   780
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Suspensos"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   14
            Top             =   210
            Width           =   765
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "2º Parcial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1275
         Left            =   2100
         TabIndex        =   8
         Top             =   120
         Width           =   2025
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Sobresalientes"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   12
            Top             =   990
            Width           =   1050
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Aprobados"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   11
            Top             =   450
            Width           =   780
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Notables"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   10
            Top             =   720
            Width           =   630
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Suspensos"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   9
            Top             =   210
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "1er Parcial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1275
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   2025
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Sobresalientes"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   7
            Top             =   990
            Width           =   1050
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Notables"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   6
            Top             =   720
            Width           =   630
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Aprobados"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   5
            Top             =   450
            Width           =   780
         End
         Begin VB.Label E_Notas 
            AutoSize        =   -1  'True
            Caption         =   "Suspensos"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   180
            Width           =   765
         End
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   7605
      Begin MSFlexGridLib.MSFlexGrid Tabla 
         Height          =   4125
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   7276
         _Version        =   393216
         Rows            =   20
         Cols            =   5
         SelectionMode   =   1
         MergeCells      =   2
         AllowUserResizing=   3
         FormatString    =   "<Clase|Alumno|Asignatura|>Nota1|Nota2"
      End
   End
   Begin VB.Menu mnu_Ppal_Arhivo 
      Caption         =   "&Archivo"
      Begin VB.Menu menu_Archivo 
         Caption         =   "&Abrir base de datos(sin implementar)"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu menu_Archivo 
         Caption         =   "&Generar Datos"
         Index           =   1
         Shortcut        =   ^G
      End
      Begin VB.Menu menu_Archivo 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu menu_Archivo 
         Caption         =   "&Salir"
         Index           =   3
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frm_Ppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Notas entre 0 y 5 y entre 5 y 10 , 1er y 2º parcial
Dim mas_5_1 As Integer, mas_5_2 As Integer, menos_5_1  As Integer, menos_5_2 As Integer
Private Sub B_Calcular_Click()
'Comprobar que la tabla no esté vacía
If (Tabla.TextArray(Indice(1, 0))) = "" Then
 MsgBox "No hay datos", vbCritical + vbSystemModal, "Atención"
 Exit Sub
End If
Dim i As Integer 'índice de la tabla de datos
Dim a_1 As Integer, a_2 As Integer 'Aprobados
Dim n_1 As Integer, n_2 As Integer 'Notables
Dim so_1 As Integer, so_2 As Integer 'Sobresalientes
LimpiarVariables 'poner a cero los contadores de notas
For i = Tabla.FixedRows To Tabla.Rows - 1
'NOTAS DEL 1er PARCIAL-----------------
 If (Tabla.TextArray(Indice(i, 3))) < 5 Then
   menos_5_1 = menos_5_1 + 1
 ElseIf (Tabla.TextArray(Indice(i, 3))) < 7 Then
   a_1 = a_1 + 1 'Aprobados
   mas_5_1 = mas_5_1 + 1
 ElseIf (Tabla.TextArray(Indice(i, 3))) < 9 Then
   n_1 = n_1 + 1 'Notables
   mas_5_1 = mas_5_1 + 1
 Else
   so_1 = so_1 + 1 'Sobresalientes
   mas_5_1 = mas_5_1 + 1
 End If
 'NOTAS DEL 2º PARCIAL------------------
 If (Tabla.TextArray(Indice(i, 4))) < 5 Then
    menos_5_2 = menos_5_2 + 1
 ElseIf (Tabla.TextArray(Indice(i, 4))) < 7 Then
   a_2 = a_2 + 1 'Aprobados
   mas_5_2 = mas_5_2 + 1
 ElseIf (Tabla.TextArray(Indice(i, 4))) < 9 Then
  n_2 = n_2 + 1 'Notables
  mas_5_2 = mas_5_2 + 1
 Else
  so_2 = so_2 + 1 'Sobresalientes
  mas_5_2 = mas_5_2 + 1
 End If
Next i
E_Notas(0).Caption = "Aprobados: " & a_1
E_Notas(1).Caption = "Suspensos: " & menos_5_1
E_Notas(2).Caption = "Notables: " & n_1
E_Notas(3).Caption = "Sobresalientes:" & so_1
E_Notas(4).Caption = "Aprobados: " & a_2
E_Notas(5).Caption = "Suspensos: " & menos_5_2
E_Notas(6).Caption = "Notables: " & n_2
E_Notas(7).Caption = "Sobresalientes:" & so_2
'Calculando medias condicionadas:
E_Notas(8).Caption = "Suspensos:" & _
 Format((((2.5 * menos_5_1) + (mas_5_1 * 7.5)) / (menos_5_1 + mas_5_1)), "#.00")
E_Notas(9).Caption = "Aprobados:" & _
 Format((((2.5 * menos_5_2) + (7.5 * mas_5_2)) / (menos_5_2 + mas_5_2)), "#.00")
E_Notas(10).Caption = "Suspensos:" & _
 Format((((2.5 * menos_5_1) + (7.5 * menos_5_2)) / (menos_5_1 + menos_5_2)), "#.00")
E_Notas(11).Caption = "Aprobados:" & _
 Format((((2.5 * mas_5_1) + (7.5 * mas_5_2)) / (mas_5_1 + mas_5_2)), "#.00")
'Calculando medias marginales de los parciales
E_Notas(14).Caption = Format((((2.5 * (menos_5_1 + menos_5_2)) + (7.5 * (mas_5_1 + mas_5_2))) / (menos_5_1 + menos_5_2 + mas_5_1 + mas_5_2)), "#.00")
E_Notas(15).Caption = Format((((2.5 * (menos_5_1 + mas_5_1)) + (7.5 * (menos_5_2 + mas_5_2))) / (menos_5_1 + menos_5_2 + mas_5_1 + mas_5_2)), "#.00")
'Varianzas condicionadas
E_Notas(16).Caption = "Suspensos:" & _
 Format((((2.5 * menos_5_1 * menos_5_1) + (7.5 * mas_5_1 * mas_5_1)) / (menos_5_1 + mas_5_1)) - ((((2.5 * menos_5_1) + (mas_5_1 * 7.5)) / (menos_5_1 + mas_5_1)) * (((2.5 * menos_5_1) + (mas_5_1 * 7.5)) / (menos_5_1 + mas_5_1))), "#.00")
E_Notas(17).Caption = "Aprobados:" & _
 Format((((2.5 * mas_5_1) * mas_5_1 + (7.5 * mas_5_2) * mas_5_2) / mas_5_1 + mas_5_2) - (((2.5 * mas_5_1) * mas_5_1) * ((7.5 * mas_5_2) * mas_5_2) * ((2.5 * mas_5_1) * mas_5_1) * ((7.5 * mas_5_2) * mas_5_2)), "#.00")
E_Notas(18).Caption = "Suspensos:" & _
 Format(((((2.5 * menos_5_1 * menos_5_1) + (7.5 * menos_5_2 * menos_5_2)) / (menos_5_1 + menos_5_2)) - ((((2.5 * menos_5_1) + (7.5 * menos_5_2)) / (menos_5_1 + menos_5_2)) * (((2.5 * menos_5_1) + (7.5 * menos_5_2)) / (menos_5_1 + menos_5_2)))), "#.00")
E_Notas(19).Caption = "Aprobados:" & _
 Format(((((2.5 * mas_5_1 * mas_5_1) + (7.5 * mas_5_2 * mas_5_2)) / (mas_5_1 + mas_5_2)) - ((((2.5 * mas_5_1) + (7.5 * mas_5_2)) / (mas_5_1 + mas_5_2)) * (((2.5 * mas_5_1) + (7.5 * mas_5_2)) / (mas_5_1 + mas_5_2)))), "#.00")
'Varianzas Marginales
E_Notas(21).Caption = Format((((((menos_5_1 + menos_5_2) * (2.5 * (menos_5_1 + menos_5_2))) + (mas_5_1 + mas_5_2 * ((mas_5_1 + mas_5_2) * 7.5))) / (menos_5_1 + menos_5_2 + mas_5_1 + mas_5_2)) - ((((2.5 * (menos_5_1 + menos_5_2)) + (7.5 * (mas_5_1 + mas_5_2))) / (menos_5_1 + menos_5_2 + mas_5_1 + mas_5_2)) * (((2.5 * (menos_5_1 + menos_5_2)) + (7.5 * (mas_5_1 + mas_5_2))) / (menos_5_1 + menos_5_2 + mas_5_1 + mas_5_2)))), "#.00")
E_Notas(23).Caption = Format(((((menos_5_1 + menos_5_2) * (2.5 * (menos_5_1 + menos_5_2))) / ((menos_5_1 + mas_5_1 + menos_5_2 + mas_5_2))) - ((((2.5 * (menos_5_1 + mas_5_1)) + (7.5 * (menos_5_2 + mas_5_2))) / (menos_5_1 + menos_5_2 + mas_5_1 + mas_5_2)) * (((2.5 * (menos_5_1 + mas_5_1)) + (7.5 * (menos_5_2 + mas_5_2))) / (menos_5_1 + menos_5_2 + mas_5_1 + mas_5_2)))), "#.00")
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Width = Me.Width - 105
    Frame1.Height = Me.Height - 3330
    Frame2.Top = Frame1.Height - 45
    Frame2.Width = Frame1.Width - 50
    Tabla.Width = Frame1.Width - 110
    Tabla.Height = Frame1.Height - 150
    B_Calcular.Left = Frame2.Width - B_Calcular.Width - 100
    txt_Adorno.Left = B_Calcular.Left - 30
End Sub

Private Sub menu_Archivo_Click(Index As Integer)
Select Case Index
 Case 0: 'Cargar base de datos
 Case 1: 'Generar datos
   Dim filas 'STRING
    filas = InputBox("Introduce el número de filas de datos a generar:", "Número de filas", 20)
   If filas = "" Then Exit Sub
   filas = Val(filas) 'CASTING A DOUBLE
   If filas <= 0 Then Exit Sub
   Tabla.Rows = filas + 1
   Dim i As Integer
   ' Crear la matriz
   filas = 0
   For i = Tabla.FixedRows To Tabla.Rows - 1
    'El nombre de la clase, hay 3, el numero de filas/3
    If filas < (Tabla.Rows / 3) Then
      Tabla.TextArray(Indice(i, 0)) = "1ºIngeniería Sistemas"
    ElseIf filas < (Tabla.Rows / 3) * 2 Then
      Tabla.TextArray(Indice(i, 0)) = "1ºIngeniería Gestión"
    Else
      Tabla.TextArray(Indice(i, 0)) = "1º Ingeniería Superior"
    End If
    filas = filas + 1
    'Nombre del alumno
    Tabla.TextArray(Indice(i, 1)) = RandomString(1)
    'Nombre de la asignatura
    Tabla.TextArray(Indice(i, 2)) = RandomString(2)
    'Nota del alumno del 1er parcial
    Tabla.TextArray(Indice(i, 3)) = Format(Rnd * 10, "#.00")
    'Nota del alumno del 2º parcial
    Tabla.TextArray(Indice(i, 4)) = Format(Rnd * 10, "#.00")
   Next i
   
   'Configurar la combinación
   Tabla.MergeCol(0) = True
   Tabla.MergeCol(1) = True
   Tabla.MergeCol(2) = True
   
   'Ordenar para ver el efecto
   Ordenar
   
   'Formato de la cuadrícula
   Tabla.ColWidth(0) = 2000
   Tabla.ColWidth(1) = 1500
   Tabla.ColWidth(2) = 1400
   Tabla.ColWidth(3) = 1200
   Tabla.ColWidth(4) = 1200
 Case 3: End 'Salir
End Select
End Sub
Function Indice(r As Integer, c As Integer) As Integer
 Indice = c + Tabla.Cols * r
End Function

Function RandomString(kind As Integer) As String
Dim s As String
Select Case kind
Case 0 'Clase
 Select Case (Rnd * 1000) Mod 3
  Case 0: s = "1ºIngeniería Sistemas"
  Case 1: s = "1ºIngeniería Gestión"
  Case Else: s = "1ºIngeniería Superior"
 End Select
Case 1 'Alumno
 Select Case (Rnd * 1000) Mod 12
  Case 0: s = "Juan"
  Case 1: s = "Waldo"
  Case 2: s = "Natalia"
  Case 3: s = "Tiffany"
  Case 4: s = "Chatín"
  Case 5: s = "Belcebú"
  Case 6: s = "Alex"
  Case 7: s = "Paco"
  Case 8: s = "Berto"
  Case 9: s = "Carlos"
  Case 10: s = "David"
  Case Else: s = "Lenna"
 End Select
Case 2 'Asignatura
 Select Case (Rnd * 1000) Mod 7
  Case 0: s = "MP II"
  Case 1: s = "Estadística"
  Case 2: s = "E.D."
  Case 3: s = "F.L.P."
  Case 4: s = "T.C"
  Case 5: s = "Dibujo"
  Case Else: s = "F.T.C."
 End Select

End Select
RandomString = s
End Function
Sub Ordenar()
    Tabla.Col = 0
    Tabla.ColSel = Tabla.Cols - 1
    Tabla.Sort = 0 'Orden Ascendente
End Sub
Sub LimpiarVariables()
    mas_5_1 = 0: mas_5_2 = 0: menos_5_1 = 0: menos_5_2 = 0
End Sub
