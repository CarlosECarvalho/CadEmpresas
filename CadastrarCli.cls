VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CadastrarCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub CadastrarCli_Initialize(Cadastro)
'========== CADASTRO DE EMPRESA
    Public Sub CNPJMkBox_LostFocus()
    Cadastro.Add (CNPJMkBox), "CNPJ"
    End Sub
    
    Private Sub IEBox_LostFocus()
    Cadastro.Add (IEBox), "IE"
    End Sub
    
    Private Sub DTCadMkBox_LostFocus()
    Cadastro.Add (DTCadMkBox), "Data"
    End Sub
    
    Private Sub RazBox_LostFocus()
    Cadastro.Add (RazBox), "Razao"
    End Sub
    
    Private Sub FanBox_LostFocus()
    Cadastro.Add (FanBox), "Nome"
    End Sub
    
    Private Sub EndBox_LostFocus()
    Cadastro.Add (EndBox), "Endereco"
    End Sub
    
    Private Sub NrBox_LostFocus()
    Cadastro.Add (NrBox), "Nro"
    End Sub
    
    Private Sub BairroBox_LostFocus()
    Cadastro.Add (BairroBox), "Bairro"
    End Sub
    
    Private Sub CidBox_LostFocus()
    Cadastro.Add (CidBox), "Cidade"
    End Sub
    
    Private Sub UFCombo_LostFocus()
    Cadastro.Add (UFCombo), "UF"
    End Sub
    
    Private Sub TelMkBox_LostFocus()
    Cadastro.Add (TelMkBox), "Tel"
    End Sub
    
    Private Sub EmailBox_LostFocus()
    Cadastro.Add (EmailBox), "Email"
    End Sub
'=================================================FIM DOS INPUTS==================================================

End Sub
