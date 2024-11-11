VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconectado"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComando 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   10095
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox cboProg 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton optSketch 
         Caption         =   "Sketch.hex"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optBootloader 
         Caption         =   "Bootloader"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optLock 
         Caption         =   "Lock"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtFile 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   5895
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   7335
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   6240
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cboBoard 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
      Begin VB.ComboBox cboPort 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdPort 
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   6240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   4080
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   120
      Picture         =   "Form1.frx":169B2
      Top             =   2880
      Width           =   10080
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Shell
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Variável global
Dim port As String
Dim board As String
Dim prog As String
Dim lockbit As String
Dim fileBootloader As String
Dim scan As Boolean
Dim config(2) As String

Private Sub Form_Load()
    ' Barra de Titulo
    Me.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
       
    ' Lista de Board
    cboBoard.AddItem "Atmega8"
    cboBoard.AddItem "Atmega328P"
    cboBoard.AddItem "ATtiny13"
    cboBoard.AddItem "ATtiny85"
    cboBoard.Text = "Atmega8" ' Board inicial
    
    ' Lista de Programador
    cboProg.AddItem "Arduino"
    cboProg.AddItem "UsbAsp"
    cboProg.Text = "Arduino" ' Programador inicial
    
    optBootloader.Value = True ' Opção inicial
    txtFile.Locked = True
    Call cboBoard_Click
    
    ' Detecta portas disponíveis
    Call cmdPort_Click
   
End Sub

Private Sub scanPort()
   cboPort.Clear
   Dim i As Integer
   For i = 1 To 16 'Procura portas COM de 1 a 16
      MSComm1.CommPort = i
      On Error Resume Next 'ignora o tratamento de erro
      MSComm1.PortOpen = True 'tenta abrir a porta
      If Err.Number = 0 Then 'a porta está disponível
         cboPort.AddItem "COM" & i
         cboPort.ListIndex = 1
         MSComm1.PortOpen = False 'fecha a porta
      End If
      On Error GoTo 0 'ativa o tratamento de erro novamente
   Next i
   
   If cboPort.List(0) <> Empty Then cboPort.Text = cboPort.List(0)
   
   ' Scan finalizado...
   cmdPort.Caption = "Port"
   Beep
   
End Sub

Private Sub cmdPort_Click()
   ' Scaneando...
   DoEvents
   cmdPort.Caption = "Scanning..."
   Sleep (1000)
   Call scanPort
   
End Sub

Private Sub cboPort_Click()
   port = cboPort.Text
   
End Sub

Private Sub cboBoard_Click()
   If cboBoard.ListIndex = 0 Then
        board = "m8" ' Atmega8
        lockbit = "0x0C"
        optLock.ToolTipText = "LB:0x0C"
        fileBootloader = verificarArquivo("bootloader_atmega8.hex")
   ElseIf cboBoard.ListIndex = 1 Then
        lockbit = "0x0C"
        board = "m328p" ' Atmega328P
        optLock.ToolTipText = "LB:0x0C"
        fileBootloader = verificarArquivo("bootloader_atmega328.hex")
   ElseIf cboBoard.ListIndex = 2 Then
        lockbit = "0x3C"
        board = "t13" ' ATtiny13
        optLock.ToolTipText = "LB:0x3C"
        fileBootloader = verificarArquivo("bootloader_attiny13.hex")
   ElseIf cboBoard.ListIndex = 3 Then
        lockbit = "0x3C"
        board = "t85" ' ATtiny85
        optLock.ToolTipText = "LB:0x3C"
        fileBootloader = verificarArquivo("bootloader_attiny85.hex")
   End If
  
End Sub

Private Sub cboProg_Click()
   If cboProg.ListIndex = 0 Then
        prog = "Arduino" ' Arduino
   ElseIf cboProg.ListIndex = 1 Then
        prog = "usbasp" ' UsbAsp
   End If
   
End Sub

Private Sub msg_Box(mensagem As String)
   MsgBox mensagem, , "DALÇÓQUIO AUTOMAÇÃO"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'DoEvents
   'Sleep (1000)
   'End

End Sub

Private Sub cmdfile_Click()
    ' Define o filtro para exibir tipo de arquivos
    If optSketch.Value = True Then
        CommonDialog1.Filter = "Arquivos (*.hex)|*.hex"
    End If
    
    ' Abre o diálogo de seleção de arquivo
    CommonDialog1.ShowOpen
    
    ' Obtém o caminho completo do arquivo selecionado
    Dim filePath As String
    filePath = CommonDialog1.FileName
    
    ' Exibe o caminho do arquivo no TextBox
    txtFile.Text = filePath
    txtFile.ToolTipText = txtFile.Text
    
End Sub

Private Sub cmdUpload_Click()
    Dim uploadCmd As String
    Dim filePath As String
    ' Carregar as variáveis de dependências
    Call cboBoard_Click ' board e lockbit
    Call cboPort_Click ' porta
    Call cboProg_Click ' programador
    
    ' BOOTLOADER
    If optBootloader.Value = True Then
        ' Especialmente para attiny13
        If cboBoard.ListIndex = 2 Then
             ' Define o comando para fazer o upload do arquivo compilado
            uploadCmd = "cmd.exe /k avrdude -u -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U lock:w:0x3F:m -U hfuse:w:0b11111011:m -U lfuse:w:0x7A:m"
            ' Executa o comando de upload e abre a janela do prompt de comando
            Shell uploadCmd, vbNormalFocus
            txtComando.Text = Mid(uploadCmd, 12, Len(uploadCmd))
            Exit Sub
        End If
        ' Carrega arquivo para upload
        filePath = fileBootloader ' Bootloader.hex
        ' Define o comando para fazer o upload do arquivo compilado
        uploadCmd = "cmd.exe /k avrdude -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U flash:w:" & filePath & ":a"
        ' Executa o comando de upload e abre a janela do prompt de comando
        Shell uploadCmd, vbNormalFocus
        txtComando.Text = Mid(uploadCmd, 12, Len(uploadCmd))
    End If

    ' SKETCH.HEX
    If optSketch.Value = True Then
        ' Carrega arquivo para upload
        filePath = txtFile.Text ' Sketch.hex
        ' Define o comando para fazer o upload do arquivo compilado
        uploadCmd = "cmd.exe /k avrdude -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U flash:w:" & filePath & ":a"
        ' Executa o comando de upload e abre a janela do prompt de comando
        Shell uploadCmd, vbNormalFocus
        txtComando.Text = Mid(uploadCmd, 12, Len(uploadCmd))
    End If
    
     ' -------------------------------------------------------------------------------------------------------------------------------------------
    ' CHATGPT: atemga8/328p: lock:0x0C unlock:0x3F  -  attiny13/85: lock:0x03C unlock:0x3F
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    ' LOCK
    If optLock.Value = True Then
        ' Define o comando para fazer o upload do lock bits
        uploadCmd = "cmd.exe /k avrdude -u -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U lock:w:" & lockbit & ":m"
        ' Executa o comando de upload e abre a janela do prompt de comando
        Shell uploadCmd, vbNormalFocus
        txtComando.Text = Mid(uploadCmd, 12, Len(uploadCmd))
    End If
    
End Sub

Private Sub Timer1_Timer()
    ' Bootloader selecionado
    If optBootloader.Value = True Then
        cmdUpload.Enabled = True
        cmdFile.Enabled = False
        txtFile.BackColor = vbWhite
    End If
    
    ' Sketch selecionado
    If optSketch.Value = True Then
        cmdFile.Enabled = True
        ' Verifica se arquivo em branco
        If txtFile.Text = Empty Then
            cmdUpload.Enabled = False
            txtFile.BackColor = vbYellow
        ' Verifica se porta em branco
        ElseIf cboPort.Text = Empty Then
            cmdUpload.Enabled = False
        Else
            cmdUpload.Enabled = True
            txtFile.BackColor = vbWhite
        End If
    End If
    
    ' Lock selecionado
    If optLock.Value = True Then
        cmdUpload.Enabled = True
        cmdFile.Enabled = False
        txtFile.BackColor = vbWhite
        optLock.ForeColor = vbRed
        optLock.Font.Bold = True
    Else
        optLock.ForeColor = vbBlack
        optLock.Font.Bold = False
    End If
   
End Sub

Private Function verificarArquivo(nomeArquivo As String) As String
    Dim caminhoArquivo As String
    Dim resultado As String
    
    caminhoArquivo = App.Path & "\" & nomeArquivo
    
    resultado = Dir(caminhoArquivo)
    
    If resultado <> Empty Then
        verificarArquivo = caminhoArquivo
    Else
        MsgBox "O arquivo " & nomeArquivo & " não foi encontrado", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        verificarArquivo = ""
    End If

End Function

Private Sub txtComando_Change()
    txtComando.ToolTipText = txtComando.Text
    
End Sub

Private Sub txtComando_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtComando <> Empty Then
        Shell "cmd.exe /k " & txtComando.Text, vbNormalFocus
    End If
    
    If KeyAscii = 1 Then
        txtComando.SelStart = 0
        txtComando.SelLength = Len(txtComando.Text)
    End If
    
End Sub
