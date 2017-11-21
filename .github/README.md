# excel-modules
Excel Modules


---
 
 ### Padrão de Desenvolvimento
 
 #### 1. Nomes dos módulos
 
 - "m_" - todas os módulos iniciam-se com "m" de módulo
 - "f_" - os módulos de função utilizam o "f" para determinar que são módulos de funções
 - "nome" - porteriormente o nome do módulo que deve ter vínculo com o fim das funções do módulo 
 
 Exemplo: 
 - m_f_Tabelas_Excel - Módulo de Funções para Tabelas de Excel
 
  #### 2. O nome das funções
 
 O nome das funções devem iniciar com letra maiúscula descrevendo literalmente o que a função irá retornar/executar.
 
 Exemplo:
 - Ordenar_Tabela_Por_Campo - Função para ordernar uma tabela do Excel por determinado campo
 
 #### 3. Comentários 
 
 Todas as funções e linhas de execução em uma função deverá ser comentada.
 
 ```
Function Ordenar_Tabela_Por_Campo(ByRef NOME_TABELA As String, _
                                  ByRef NOME_CAMPO As String)
'
' Ordena Tabela por determinada Coluna (NOME_CAMPO)
'
Dim TABELA As ListObject

    ' Remove visualização em tela
    Application.ScreenUpdating = False
    On Error GoTo TrataErro
        ' Busca a tabela conforme nome informado
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    
    ' Limpa todos os filtros atuais na tabela
    TABELA.Sort.SortFields.Clear
    
    ' Insere o filtro na tabela
    TABELA.Sort.SortFields.Add Key:=Range(NOME_TABELA & "[[#All],[" & NOME_CAMPO & "]]") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    ' Efetiva (visualmente) o filtro na tabela
    With TABELA.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Exit Function
TrataErro:
    MsgBox "Erro ao ordendar a tabela: '" & NOME_TABELA & "'" & vbNewLine & _
        "Provavelmente esta tabela foi excluída indevidamente" & vbNewLine & vbNewLine & _
        "Módulo: " & Application.VBE.ActiveCodePane.CodeModule.Name, vbCritical, _
        "Erro - " & "Ordenar_Tabela_Por_Campo"

End Function

```
 
 
 
 
 
