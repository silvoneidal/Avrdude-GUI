VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconectado"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
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
         TabIndex        =   12
         Top             =   1920
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
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optUnlock 
         Caption         =   "UnLock"
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
         Top             =   600
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
         Top             =   960
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
         Top             =   240
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
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   3840
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
      Top             =   3960
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
    cboProg.AddItem "UsbAsp"
    cboProg.AddItem "Arduino"
    cboProg.Text = "UsbAsp" ' Programador inicial
    
    ' Configuração inicial
    config(0) = cboBoard.Text
    config(1) = optSketch.Caption
    config(2) = cboProg.Text
    
    optSketch.Value = True
    txtFile.Locked = True
    
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
        optUnlock.ToolTipText = "LB:0xCF"
   ElseIf cboBoard.ListIndex = 1 Then
        lockbit = "0x0C"
        board = "m328p" ' Atmega328P
        optLock.ToolTipText = "LB:0x0C"
        optUnlock.ToolTipText = "LB:0xCF"
   ElseIf cboBoard.ListIndex = 2 Then
        lockbit = "0x3C"
        board = "t13" ' ATtiny13
        optLock.ToolTipText = "LB:0x3C"
        optUnlock.ToolTipText = "LB:0xCF"
   ElseIf cboBoard.ListIndex = 3 Then
        lockbit = "0x3C"
        board = "t85" ' ATtiny85
        optLock.ToolTipText = "LB:0x3C"
        optUnlock.ToolTipText = "LB:0xCF"
   End If
   
   config(0) = cboBoard.Text
   Call setArquivo
  
End Sub

Private Sub cboProg_Click()
   If cboProg.ListIndex = 0 Then
        prog = "usbasp" ' UsbAsp
   ElseIf cboProg.ListIndex = 1 Then
        prog = "Arduino" ' Arduino
   End If
   
   config(2) = cboProg.Text
   
End Sub

Private Sub optLock_Click()
    txtFile.Locked = True
    cmdFile.Enabled = False
    cmdUpload.Enabled = True
    txtFile.Text = Empty
    txtFile.ToolTipText = Empty
    config(1) = optLock.Caption
    
End Sub

Private Sub optUnlock_Click()
    txtFile.Locked = True
    cmdFile.Enabled = False
    cmdUpload.Enabled = True
    txtFile.Text = Empty
    txtFile.ToolTipText = Empty
    config(1) = optUnlock.Caption
    
End Sub


Private Sub optBootloader_Click()
    txtFile.Locked = True
    cmdFile.Enabled = False
    cmdUpload.Enabled = True
    Call setArquivo
    config(1) = optBootloader.Caption
    
End Sub

Private Sub optSketch_Click()
    txtFile.Locked = False
    cmdFile.Enabled = True
    cmdUpload.Enabled = True
    txtFile.Text = Empty
    txtFile.ToolTipText = Empty
    config(1) = optSketch.Caption
    
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
    ' Aguarda confirmação do usuário
    Beep
    Dim resposta As Integer
    resposta = MsgBox("Fazer upload nessas configurações: " & vbCrLf & _
                      "Board: " & config(0) & vbCrLf & _
                      "Comando: " & config(1) & vbCrLf & _
                      "Programador:  " & config(2), vbYesNo + vbQuestion, "DALÇÓQUIO AUTOMAÇÃO")
    If resposta = vbNo Then Exit Sub
    
    Dim uploadCmd As String
    Dim filePath As String
    ' Carregar as variáveis de dependências
    Call cboBoard_Click ' board e lockbit
    Call cboPort_Click ' porta
    Call cboProg_Click ' programador
    filePath = txtFile.Text ' sketch.hex
    
    Call cboBoard_Click ' Busca a board selecionada
        
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    ' CHATGPT: atemga8/328p: lock:0x0C unlock:0x3F  -  attiny13/85: lock:0x03C unlock:0x3F
    ' -------------------------------------------------------------------------------------------------------------------------------------------
    ' LOCK
    If optLock.Value = True Then
        Beep
        resposta = MsgBox("Você tem certeza que deseja usar a opção Lock (Bloqueo)", vbYesNo + vbExclamation, "DALÇÓQUIO AUTOMAÇÃO")
        If resposta = vbYes Then
            ' Define o comando para fazer o upload do lock bits
            uploadCmd = "cmd.exe /k avrdude -u -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U lock:w:" & lockbit & ":m"
            ' Executa o comando de upload e abre a janela do prompt de comando
            Shell uploadCmd, vbNormalFocus
        End If
    End If
    
    ' UNLOCK
    If optUnlock.Value = True Then
        ' Define o comando para fazer o upload do lock bits
        uploadCmd = "cmd.exe /k avrdude -u -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U lock:w:0xCF:m"
        ' Executa o comando de upload e abre a janela do prompt de comando
        Shell uploadCmd, vbNormalFocus
    End If
    
    ' BOOTLOADER
    If optBootloader.Value = True Then
        ' Define o comando para fazer o upload do arquivo compilado
        uploadCmd = "cmd.exe /k avrdude -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U flash:w:" & filePath & ":a"
        ' Executa o comando de upload e abre a janela do prompt de comando
        Shell uploadCmd, vbNormalFocus
    End If

    ' SKETCH.HEX
    If optSketch.Value = True Then
        ' Define o comando para fazer o upload do arquivo compilado
        uploadCmd = "cmd.exe /k avrdude -c " & prog & " -p " & board & " -P " & port & " -b 19200 -F -v -v -U flash:w:" & filePath & ":a"
        ' Executa o comando de upload e abre a janela do prompt de comando
        Shell uploadCmd, vbNormalFocus
    End If
    
End Sub

Private Sub Timer1_Timer()
    ' Se opçao carregar selecionada
    If optSketch.Value = True Then
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
    Else
        txtFile.BackColor = vbWhite
    End If
    
    ' Se opção lock (bloqueo) selecionada
    If optLock.Value = True Then
        optLock.ForeColor = vbRed
        optLock.Font.Bold = True
    Else
        optLock.ForeColor = vbBlack
        optLock.Font.Bold = False
    End If
   
End Sub

Private Sub setArquivo()
    ' Seleciona arquivos hexadecimais
    If optBootloader.Value = True Then
        Dim nomeArquivo As String
        Select Case cboBoard.ListIndex
            Case 0
                txtFile.Text = verificarArquivo("bootloader_atmega8.hex")
                txtFile.ToolTipText = txtFile.Text
            Case 1
                txtFile.Text = verificarArquivo("bootloader_atmega328.hex")
                txtFile.ToolTipText = txtFile.Text
            Case 2
                txtFile.Text = verificarArquivo("bootloader_attiny13.hex")
                txtFile.ToolTipText = txtFile.Text
            Case 3
                txtFile.Text = verificarArquivo("bootloader_attiny85.hex")
                txtFile.ToolTipText = txtFile.Text
        End Select
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


