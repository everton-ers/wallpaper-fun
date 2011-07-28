VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   90
      Top             =   135
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
            
    'winzcv.dll   Imagem camuflada
    'msmask32.bat Bat para renomear a imagem
                
    Dim lobjFile     As FileSystemObject
    Dim lobjText     As TextStream
    Dim lsCaminhoBat As String
    Dim lsNumero     As String
    
    Randomize
                      
    lsNumero = CStr(CInt(Rnd * 2))
    
    Randomize
    
    lsNumero = lsNumero & CStr(CInt(Rnd * 3))
    
    Randomize
    
    lsNumero = lsNumero & CStr(CInt(Rnd * 4))
    
    Randomize
    
    lsNumero = lsNumero & CStr(CInt(Rnd * 5))
    
    Randomize
    
    lsNumero = lsNumero & CStr(CInt(Rnd * 6))
    
    msNomeArquivo = lsNumero
    
    'Define o arquivo.
    lsCaminhoBat = App.Path & "\msmask32.bat"
    
    'Instância de um novo objeto de arquivo.
    Set lobjFile = New FileSystemObject
    
    'Se o bat existir, exclui ele.
    If (lobjFile.FileExists(lsCaminhoBat) = True) Then lobjFile.DeleteFile (lsCaminhoBat)
    
    'Cria o arquivo de impressão de etiqueta.
    Call lobjFile.CreateTextFile(lsCaminhoBat, True, False)
    
    'Abre o arquivo para escrita.
    Set lobjText = lobjFile.OpenTextFile(lsCaminhoBat, ForWriting, True, TristateFalse)
     
    'Grava os dados no arquivo.
    lobjText.WriteLine "C:"
    lobjText.WriteLine "cd\"
    lobjText.WriteLine "copy c:\winnt\system32\winzcv.dll c:\winnt"
    lobjText.WriteLine "cd winnt"
    lobjText.WriteLine "ren winzcv.dll " & msNomeArquivo & ".bmp"
            
    'Fecha o arquivo.
    lobjText.Close
        
    'Executa o bat.
    Shell lsCaminhoBat, vbHide
       
End Sub

Private Sub Timer1_Timer()
    
    'P/ chamar:
    SetWallpaper "C:\WINNT\" & Trim(msNomeArquivo) & ".bmp"
    
    'Finaliza a aplicação.
    End
    
End Sub
