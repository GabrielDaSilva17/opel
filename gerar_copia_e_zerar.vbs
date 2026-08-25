Option Explicit

Dim fso, pastaAtual, caminhoDocPrincipal, pastaCopias
Dim nomeDocOriginal, caminhoDocCopia, timestamp
Dim wordApp, doc, shp
Dim shapesParaZerar, nomeShape

' 1. Configuração dos caminhos e arquivos
Set fso = CreateObject("Scripting.FileSystemObject")
pastaAtual = fso.GetParentFolderName(WScript.ScriptFullName)
nomeDocOriginal = "doc ausbildund.docm"
caminhoDocPrincipal = pastaAtual & "\" & nomeDocOriginal
pastaCopias = pastaAtual & "\copias"

' Verifica se o documento original existe
If Not fso.FileExists(caminhoDocPrincipal) Then
    MsgBox "Erro: O arquivo '" & nomeDocOriginal & "' nao foi encontrado na pasta!", vbCritical, "VBScript"
    WScript.Quit
End If

' Cria a pasta de cópias se não existir
If Not fso.FolderExists(pastaCopias) Then
    fso.CreateFolder(pastaCopias)
End If

' 2. Gera o nome da cópia com data/hora (ex: doc_ausbildund_copia_2026-08-19_111500.docm)
timestamp = Year(Now) & "-" & _
            Right("0" & Month(Now), 2) & "-" & _
            Right("0" & Day(Now), 2) & "_" & _
            Right("0" & Hour(Now), 2) & _
            Right("0" & Minute(Now), 2) & _
            Right("0" & Second(Now), 2)

caminhoDocCopia = pastaCopias & "\doc_ausbildund_copia_" & timestamp & ".docm"

' 3. Cria a cópia do documento preenchido na pasta de cópias
fso.CopyFile caminhoDocPrincipal, caminhoDocCopia, True

' 4. Abre o documento principal no Word em segundo plano para zerar os campos
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
wordApp.DisplayAlerts = 0 ' wdAlertsNone

Set doc = wordApp.Documents.Open(caminhoDocPrincipal)

' Lista com os IDs exatos das caixas que serão limpas
shapesParaZerar = Array( _
    "bm_vom", _
    "bm_bis", _
    "bm_nr", _
    "bm_montag", _
    "bm_dienstag", _
    "bm_mittwoch", _
    "bm_donnerstag", _
    "bm_freitag" _
)

' Zera o texto de cada Shape
On Error Resume Next
For Each nomeShape In shapesParaZerar
    Set shp = doc.Shapes(nomeShape)
    If Not shp Is Nothing Then
        shp.TextFrame.TextRange.Text = ""
    End If
    Set shp = Nothing
Next
On Error GoTo 0

' Salva o documento principal limpo e fecha o Word
doc.Save
doc.Close
wordApp.Quit

Set doc = Nothing
Set wordApp = Nothing
Set fso = Nothing

MsgBox "Processo concluido com sucesso!" & vbCrLf & vbCrLf & _
       "1. Copia salva em: \copias\" & vbCrLf & _
       "2. Documento principal foi zerado.", vbInformation, "Automação Concluída"
