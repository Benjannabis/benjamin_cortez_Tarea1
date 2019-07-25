VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton FIN 
      Caption         =   "FIN"
      Height          =   615
      Left            =   7560
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "Form1.frx":0000
      Left            =   480
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   2
   End
   Begin VB.CommandButton btnGenerar 
      Caption         =   "Generar CrystalReport"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Data DataComunas 
      Caption         =   "DataComunas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data DataResultado 
      Caption         =   "DataResultado"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data DataCenso 
      Caption         =   "DataCenso"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnCalcular 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Data DataProvincias 
      Caption         =   "DataProvincias"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data DataRegiones 
      Caption         =   "DataRegiones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   1  'ODBCCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox ListaFechas 
      Height          =   1425
      Left            =   5040
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox ListaProvincias 
      Height          =   1425
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ListBox ListaRegiones 
      Height          =   1425
      ItemData        =   "Form1.frx":001C
      Left            =   360
      List            =   "Form1.frx":001E
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Busqueda en CENSO por AÑO y PROVINCIA"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblProvincias 
      Caption         =   "Provincias"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblRegiones 
      Caption         =   "Regiones"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ruta As String, ContRegistroTV As Integer, ContRegistro As Integer, i As Integer, FechaSeleccionada As Integer
Dim ContadorTV, ContadorPC, ContadorHorno As Integer, lblProvincia, lblRegion As String, RutaResultado As String



Private Sub btnCalcular_Click()

RegionSeleccionada = ListaRegiones.ItemData(ListaRegiones.ListIndex)
ProvinciaSeleccionada = ListaProvincias.ItemData(ListaProvincias.ListIndex)
FechaSeleccionada = ListaFechas.Text


If ListaFechas = "" Then
MsgBox ("INGRESE UNA AÑO")
Else

Do While Not DataComunas.Recordset.EOF

    If DataComunas.Recordset.Fields("COD_PROV") = ProvinciaSeleccionada Then
    ComunaAnalizada = DataComunas.Recordset.Fields("NOMBRE_COMUNA")
       
       
       Do While Not DataCenso.Recordset.EOF
        LeerFecha = DataCenso.Recordset.Fields("FECHA")
            If LeerFecha = FechaSeleccionada Then
                If DataCenso.Recordset.Fields("COD_COM") = DataComunas.Recordset.Fields("COD_COMUNA") Then
                    ContadorTV = ContadorTV + DataCenso.Recordset.Fields("TV")
                    ContadorPC = ContadorPC + DataCenso.Recordset.Fields("COMP")
                    ContadorHorno = ContadorHorno + DataCenso.Recordset.Fields("HORNO")
                    ContadorGrupoFam = ContadorGrupoFam + DataCenso.Recordset.Fields("GRUPO_FAM")
                End If
            End If
        DataCenso.Recordset.MoveNext
        Loop


        
        

respuesta = MsgBox("Está apunto de guardar estos datos en la base de datos" & vbCrLf & "¿Está seguro que desea guardarlos?" & vbCrLf & vbCrLf & vbCrLf & "CENSO del AÑO " & FechaSeleccionada & vbCrLf & vbCrLf & "Región: " & ListaRegiones.Text & vbCrLf & "Provincia: " & ListaProvincias.Text & vbCrLf & "Comuna: " & ComunaAnalizada & vbCrLf & vbCrLf & "Grupo Familiar: " & ContadorGrupoFam & vbCrLf & "Televisores: " & ContadorTV & vbCrLf & "Computadores: " & ContadorPC & vbCrLf & "Hornos: " & ContadorHorno, vbOKCancel, "CENSO: Confirmar Guardado de Datos - Benjamin Cortez")
    If respuesta = vbOK Then
    
        'Agregamos un nuevo registro
        DataResultado.Recordset.AddNew
        DataResultado.Recordset.Fields("NOMBRE_REG") = ListaRegiones.Text
        DataResultado.Recordset.Fields("NOMBRE_PROV") = ListaProvincias.Text
        DataResultado.Recordset.Fields("Nom_comu") = ComunaAnalizada
        DataResultado.Recordset.Fields("Total_Hab") = ContadorGrupoFam
        DataResultado.Recordset.Fields("Total_TV") = ContadorTV
        DataResultado.Recordset.Fields("Total_Comp") = ContadorPC
        DataResultado.Recordset.Fields("Total_Hornos") = ContadorHorno
        DataResultado.Recordset.Update
    End If
        
        
       
'MsgBox ("Está apunto de guardar estos datos en la base de datos" & vbCrLf & "¿Está seguro que desea guardarlos?" & vbCrLf & vbCrLf & vbCrLf & "CENSO AÑO " & FechaSeleccionada & vbCrLf & "Provincia: " & lblProvincia & vbCrLf & "Comuna: " & ComunaAnalizada & vbCrLf & vbCrLf & "Grupo Familiar: " & ContadorGrupoFam & vbCrLf & "Televisores: " & ContadorTV & vbCrLf & "Computadores: " & ContadorPC & vbCrLf & "Hornos: " & ContadorHorno)
End If

DataComunas.Recordset.MoveNext
DataCenso.Recordset.MoveFirst
ContadorTV = 0
ContadorPC = 0
ContadorHorno = 0
ContadorGrupoFam = 0
Loop
DataComunas.Recordset.MoveFirst

End If
End Sub

Private Sub btnGenerar_Click()
'Dim Formula As String


'With CrystalReport1
  '  .ReportFileName = Ruta + "/listado.rpt"
   ' .DataFiles(0) = RutaResultado
   ' .ReportSource = DataResultados
    
   ' .Action = 1
    
'End With

CrystalReport1.ReportFileName = Ruta & "/listado.rpt"
CrystalReport1.Action = 1


End Sub

Private Sub FIN_Click()
End
End Sub

Private Sub Form_Load()
Ruta = App.Path


ListaProvincias.Clear


DataRegiones.DatabaseName = Ruta & "/POE_CENSO.MDB"
DataRegiones.RecordSource = "REGIONES"

DataProvincias.DatabaseName = Ruta & "/POE_CENSO.MDB"
DataProvincias.RecordSource = "PROVINCIAS"

DataCenso.DatabaseName = Ruta & "/POE_CENSO.MDB"
DataCenso.RecordSource = "CENSO"

DataResultado.DatabaseName = Ruta & "/POE_CENSO.MDB"
RutaResultado = DataResultado.DatabaseName
DataResultado.RecordSource = "RESULTADO"

DataComunas.DatabaseName = Ruta & "/POE_CENSO.MDB"
DataComunas.RecordSource = "COMUNAS"


End Sub


Private Sub form_activate()
Cont = 0
DataRegiones.Recordset.MoveFirst


Do While Not DataRegiones.Recordset.EOF
ListaRegiones.AddItem DataRegiones.Recordset.Fields("NOMBRE_REGION")
ListaRegiones.ItemData(Cont) = DataRegiones.Recordset.Fields("COD_REGION")
Cont = Cont + 1

DataRegiones.Recordset.MoveNext
Loop






End Sub
Private Sub ListaProvincias_Click()
ListaFechas.Clear
For Fechas = 2010 To 2012
ListaFechas.AddItem Fechas
Next
End Sub

Private Sub ListaRegiones_Click()

COD_REGION = ListaRegiones.ItemData(ListaRegiones.ListIndex)
ListaFechas.Clear
ListaProvincias.Clear
Cont = 0
DataProvincias.Recordset.MoveFirst

Do While Not DataProvincias.Recordset.EOF
If COD_REGION = DataProvincias.Recordset.Fields("COD_REGION") Then

ListaProvincias.AddItem DataProvincias.Recordset.Fields("NOMBRE_PROV")
ListaProvincias.ItemData(Cont) = DataProvincias.Recordset.Fields("COD_PROV")
Cont = Cont + 1
End If
DataProvincias.Recordset.MoveNext

Loop



End Sub



