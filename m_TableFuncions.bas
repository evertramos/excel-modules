' -------------------------------------------------------
'
' Módulo Table Functions
'
' Possui diversas funções para trabalhar com tabelas
' pré-formatadas no Excel
'
' Versão: 0.6.2
'
' -------------------------------------------------------
'
' Lista de funções:
'
' limparFiltroTabela
' filtrarTabelaPorCampo
' retornarNumeroColuna
' retornarQtdLinhasTabela
' retornarQtdLinhasColuna
' Retornar_Nome_Todas_Tabelas
' selecionarDadosDaColuna
' selecionarDadosDaTabela
' selecionarLinhaDaTabela
' converterTextoEmNumero
' deletarDadosTabela
' inserirNovaColuna
' inserirFormulaColuna
' inserirDadoTabela
' retornaNomePlanilha
' activateTable
' deactivateTable
' verificaTabelaExiste
'
'
' Funções Privadas:
'
' addSortFieldToTable
' inserirNovaLinha
' checkColumnExists

' -------------------------------------------------------

Option Explicit

' -------------------------------------------------------
' Funções de Ordenação de Dados
' -------------------------------------------------------

Function ordenarTabelaPorCampo(ByRef NOME_TABELA As String, _
                                Optional ByRef NOME_CAMPO_S As Variant)
'
' Ordena Tabela por coluna
'
Dim TABELA As ListObject
Dim FIELD_NAME As Variant

    Application.ScreenUpdating = False
    On Error GoTo ErroOrdenar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    
    TABELA.Sort.SortFields.Clear
    If IsArray(NOME_CAMPO_S) Then
        For Each FIELD_NAME In NOME_CAMPO_S
            ' Função privada dentro desse módulo
            addSortFieldToTable TABELA, NOME_TABELA, CStr(FIELD_NAME)
        Next
    Else
        ' Função privada dentro desse módulo
        addSortFieldToTable TABELA, NOME_TABELA, CStr(NOME_CAMPO_S)
    End If
    With TABELA.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Exit Function
ErroOrdenar:
    MsgBox "Erro ao ordendar a tabela: '" & NOME_TABELA & "'" & vbNewLine & _
        "Provavelmente foi excluída indevidamente", vbCritical, "Erro - ordenarTabelaPorCampo - Módulo: m_TableFunctions"

End Function

' -------------------------------------------------------
' Funções de Filtros nas Tabelas
' -------------------------------------------------------

Function filtrarTabelaPorCampo(ByRef NOME_TABELA As String, _
                                ByRef NOME_CAMPO As String, _
                                ByRef TEXTO_FILTRO As String, _
                                Optional ByRef LIMPA_FILTRO As Boolean)
'
' Filtra Tabela por coluna
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    
    Application.ScreenUpdating = False
    On Error GoTo ErroFiltrar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_CAMPO, TABELA.HeaderRowRange, 0)
    
    If LIMPA_FILTRO Then
        limparFiltroTabela NOME_TABELA
    End If
    
    If ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible <> xlVisible Then
        activateTable NOME_TABELA
    End If
    
    TABELA.ShowAutoFilter = True
    TABELA.Range.AutoFilter Field:=COLUMN_NUMBER, Criteria1:=TEXTO_FILTRO
    
    
    
Exit Function
ErroFiltrar:
    MsgBox "Erro ao filtrar a tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - filtrarTabelaPorCampo - Módulo: m_TableFunctions"

End Function

Function Filtrar_Tabela_Texto_Inicia_Com(ByRef NOME_TABELA As String, _
                                         ByRef NOME_CAMPO As String, _
                                         ByRef TEXTO_FILTRO As String, _
                                         Optional ByRef LIMPA_FILTRO As Boolean = True)
'
' Filtra Tabela em coluna específica onde inicia-se com o texto TEXTO_FILTRO
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    
    Application.ScreenUpdating = False
    On Error GoTo ErroFiltrar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_CAMPO, TABELA.HeaderRowRange, 0)
    
    If LIMPA_FILTRO Then
        limparFiltroTabela NOME_TABELA
    End If
    
    If ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible <> xlVisible Then
        activateTable NOME_TABELA
    End If
    
    TABELA.ShowAutoFilter = True
    TABELA.Range.AutoFilter Field:=COLUMN_NUMBER, Criteria1:=TEXTO_FILTRO & "*", Operator:=xlAnd
    
Exit Function
ErroFiltrar:
    MsgBox "Erro ao filtrar a tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - Filtrar_Tabela_Texto_Inicia_Com - Módulo: m_TableFunctions"

End Function

Function Filtrar_Tabela_Numero(ByRef NOME_TABELA As String, _
                               ByRef NOME_CAMPO As String, _
                               ByRef NUMERO_FILTRO As String, _
                               Optional OPERACAO As String = ">=", _
                               Optional ByRef LIMPA_FILTRO As Boolean = True)
'
' Filtra Tabela em coluna específica onde inicia-se com o texto TEXTO_FILTRO
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    
    Application.ScreenUpdating = False
    On Error GoTo ErroFiltrar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_CAMPO, TABELA.HeaderRowRange, 0)
    
    If LIMPA_FILTRO Then
        limparFiltroTabela NOME_TABELA
    End If
    
    If ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible <> xlVisible Then
        activateTable NOME_TABELA
    End If
    
    TABELA.ShowAutoFilter = True
    TABELA.Range.AutoFilter Field:=COLUMN_NUMBER, Criteria1:=OPERACAO & NUMERO_FILTRO, Operator:=xlAnd
    
Exit Function
ErroFiltrar:
    MsgBox "Erro ao filtrar a tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - Filtrar_Tabela_Numero - Módulo: m_TableFunctions"

End Function


Function limparFiltroTabela(ByRef NOME_TABELA As String)
'
' Limpa todos os filtros de uma Tabela
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo TabelaInexistente
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    TABELA.ShowAutoFilter = True
    If TABELA.AutoFilter.FilterMode Then
        TABELA.AutoFilter.ShowAllData
    End If
    
Exit Function
TabelaInexistente:
    MsgBox "Erro ao limpar filtro da tabela: '" & NOME_TABELA & "'" & vbNewLine & _
        "Provavelmente foi excluída indevidamente", vbCritical, "Erro - limparFiltroTabela - Módulo: m_TableFunctions"

End Function

' -------------------------------------------------------
' Funções de Retorno de Dados
' -------------------------------------------------------

Function retornaNomePlanilha(ByVal NOME_TABELA As String) As String
'
' Retorna Nome da planilha com determinada Tabela
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo Erro
        retornaNomePlanilha = Range(NOME_TABELA).Parent.Name

Exit Function
Erro:
    retornaNomePlanilha = ""
    Exit Function
End Function

Function retornarNumeroColuna(ByVal NOME_TABELA As String, _
                                    ByVal NOME_COLUNA As String) As Integer
'
' Retorna o número da coluna de um determinado campo (cabeçalho) de uma tabela
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroRetornar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)
    retornarNumeroColuna = TABELA.ListColumns(COLUMN_NUMBER).Range.Column
        
Exit Function
ErroRetornar:
    MsgBox "Erro buscar número da coluna na tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - retornarNumeroColuna - Módulo: m_TableFunctions"

End Function

Function retornarQtdLinhasTabela(ByVal NOME_TABELA As String, _
                                 Optional ByVal CONSIDERA_FILTRO As Boolean = True) As Integer
'
' Retorna a quantidade de linhas de uma tabela, levando em consideração o filtro atual aplicado
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo ErroRetornar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    
    If CONSIDERA_FILTRO Then
        retornarQtdLinhasTabela = TABELA.DataBodyRange.Columns(1).SpecialCells(xlCellTypeVisible).count
    Else
        retornarQtdLinhasTabela = TABELA.DataBodyRange.Rows.count
    End If
        
Exit Function
ErroRetornar:
    MsgBox "Erro ao buscar a quantidade de linhas da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - retornarQtdLinhasTabela - Módulo: m_TableFunctions"

End Function

Function retornarQtdLinhasColuna(ByVal NOME_TABELA As String, _
                                 ByVal NOME_COLUNA As String, _
                                 Optional ByVal SOMENTE_NAO_VAZIAS As Boolean = True) As Integer
'
' Retorna a quantidade de linhas de uma tabela, levando em consideração o filtro atual aplicado
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroRetornar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)
    
    If SOMENTE_NAO_VAZIAS Then
        retornarQtdLinhasColuna = TABELA.DataBodyRange.Columns(COLUMN_NUMBER).SpecialCells(xlCellTypeConstants).count
    Else
        'retornarQtdLinhasColuna = TABELA.DataBodyRange.Columns(COLUMN_NUMBER).Rows.COUNT
        retornarQtdLinhasColuna = TABELA.DataBodyRange.Rows.count
    End If
        
Exit Function
ErroRetornar:
    MsgBox "Erro ao buscar a quantidade de linhas da coluna: '" & NOME_COLUNA & "da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - retornarQtdLinhasColuna - Módulo: m_TableFunctions"

End Function

Function Retornar_Nome_Todas_Tabelas() As Variant
'
' Retorna o número da tabelas da planilha
'
Dim PLANILHA As Worksheet
Dim TABELA As ListObject
Dim NOMES As Variant
Dim i As Integer
    
    Application.ScreenUpdating = False
    
    ReDim NOMES(1 To Retornar_Qtd_Tabelas)
    
    On Error GoTo ErroRetornar
    i = 1
    'Loop em todas as planilhas
    For Each PLANILHA In ThisWorkbook.Worksheets
        For Each TABELA In PLANILHA.ListObjects
            NOMES(i) = TABELA.Name
            i = i + 1
        Next
    Next
    Retornar_Nome_Todas_Tabelas = NOMES
    
Exit Function
ErroRetornar:
    MsgBox "Erro buscar o nome das tabelas da planilha.", vbCritical, "Erro -  Retornar_Nome_Todas_Tabelas - Módulo: m_TableFunctions"

End Function

Function Retornar_Qtd_Tabelas() As Integer
'
' Retorna o número da tabelas da planilha
'
Dim PLANILHA As Worksheet
Dim TABELA As ListObject
Dim i As Integer
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErroRetornar
    i = 0
    'Loop em todas as planilhas
    For Each PLANILHA In ThisWorkbook.Worksheets
        For Each TABELA In PLANILHA.ListObjects
            i = i + 1
        Next
    Next
    Retornar_Qtd_Tabelas = i
    
Exit Function
ErroRetornar:
    MsgBox "Erro buscar a quantidade de tabelas da planilha.", vbCritical, "Erro -  Retornar_Qtd_Tabelas - Módulo: m_TableFunctions"

End Function

' -------------------------------------------------------
' Funções de Seleção de Dados
' -------------------------------------------------------

Function selecionarDadosDaColuna(ByVal NOME_TABELA As String, _
                                    ByVal NOME_COLUNA As String, _
                                    Optional ByVal TODOS_OS_DADOS As Boolean = False) As Boolean
'
' Seleciona dados de uma coluna
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroSelecionar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)

    activateTable NOME_TABELA
    
    If TODOS_OS_DADOS Then
        TABELA.ListColumns(COLUMN_NUMBER).DataBodyRange.Select
    Else
        On Error GoTo SemDados
            TABELA.ListColumns(COLUMN_NUMBER).DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    End If
    selecionarDadosDaColuna = True
    
Exit Function
SemDados:
    selecionarDadosDaColuna = False
    Exit Function
ErroSelecionar:
    MsgBox "Erro ao selecionar os dados da coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - selecionarDadosDaColuna - Módulo: m_TableFunctions"
    selecionarDadosDaColuna = False
    Exit Function
End Function

Function selecionarDadosDaTabela(ByVal NOME_TABELA As String, _
                                 Optional ByVal TODOS_OS_DADOS As Boolean = False) As Boolean
'
' Seleciona todos os dados de uma coluna
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErroSelecionar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
   
    On Error GoTo SemDados
        activateTable NOME_TABELA
    
    If TODOS_OS_DADOS Then
        TABELA.DataBodyRange.Select
    Else
        TABELA.DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    End If
    selecionarDadosDaTabela = True
    
Exit Function
SemDados:
    selecionarDadosDaTabela = False
    Exit Function
ErroSelecionar:
    MsgBox "Erro ao selecionar os dados da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - selecionarDadosDaTabela - Módulo: m_TableFunctions"
    selecionarDadosDaTabela = False
    Exit Function
End Function

Function selecionarLinhaDaTabela(ByVal NOME_TABELA As String, _
                                 ByVal LINHA_TABELA As Integer, _
                                 Optional ByVal DIF_LINHA As Integer = 0, _
                                 Optional ByVal LINHA_INTEIRA As Boolean = False) As Boolean
'
' Seleciona dados de uma coluna
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroSelecionar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
   
    On Error GoTo SemDados
        activateTable NOME_TABELA
        
        If LINHA_INTEIRA Then
            TABELA.ListRows(LINHA_TABELA + DIF_LINHA).Range.Select
        Else
            TABELA.ListRows(LINHA_TABELA + DIF_LINHA).Range(, 1).Select
        End If
        selecionarLinhaDaTabela = True
    
Exit Function
SemDados:
    selecionarLinhaDaTabela = False
    Exit Function
ErroSelecionar:
    MsgBox "Erro ao selecionar a linha: '" & LINHA_TABELA & "'" & " da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - selecionarLinhaDaTabela - Módulo: m_TableFunctions"
    selecionarLinhaDaTabela = False
    Exit Function
End Function

' -------------------------------------------------------
' Funções de Formatação
' -------------------------------------------------------

Function formataAutoFit(ByVal NOME_TABELA As String, _
                 ByVal NOME_COLUNA As String) As Boolean
'
' Formata a coluna com autofit
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroFormatar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)

    'activateTable NOME_TABELA
    
    TABELA.ListColumns(COLUMN_NUMBER).Range.Columns.autoFit
    formataAutoFit = True
    
Exit Function
ErroFormatar:
    MsgBox "Erro ao formatar os dados da coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - formataAutoFit - mod_TableFunctions"
    formataAutoFit = False
    Exit Function
End Function

Function formataNumero(ByVal NOME_TABELA As String, _
                       ByVal NOME_COLUNA As String, _
                       Optional ByRef FORMATO As String = "#,##0.00") As Boolean
'
' Formata a coluna como número
'
    Application.ScreenUpdating = False
    On Error GoTo ErroFormatar
    If selecionarDadosDaColuna(NOME_TABELA, NOME_COLUNA) = False Then
        formataNumero = False
        Exit Function
    End If
    Selection.NumberFormat = FORMATO
    formataNumero = True
    Exit Function

ErroFormatar:
    formataNumero = False
    MsgBox "Erro ao formatar a coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - formataNumero - mod_TableFunctions"
End Function

Function converterTextoEmNumero(ByVal NOME_TABELA As String, _
                                    ByVal NOME_COLUNA As String)
'
' Converte os dados de determinada Coluna em determinada Tabela em número
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroConverter
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)
    TABELA.ListColumns(COLUMN_NUMBER).DataBodyRange = TABELA.ListColumns(COLUMN_NUMBER).DataBodyRange.Value
    
    ' Modo simples
    'Range(NOME_TABELA & "[" & NOME_COLUNA & "]") = Range(NOME_TABELA & "[" & NOME_COLUNA & "]").Value
        
Exit Function
ErroConverter:
    MsgBox "Erro buscar converter os dados da coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - converterTextoEmNumero - Módulo: m_TableFunctions"

End Function

' -------------------------------------------------------
' Funções de Inserção, Edição e Exclusão de Dados
' -------------------------------------------------------

Function Excluir_Linhas_em_Branco(ByVal NOME_TABELA As String)
'
' Deleta todos os dados de uma Tabela
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo ErroDeletar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    limparFiltroTabela NOME_TABELA
    If TABELA.ListRows.count > 0 Then
        On Error GoTo SemLinhaEmBranco
        TABELA.DataBodyRange.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    End If
       
Exit Function
SemLinhaEmBranco:
    Exit Function
ErroDeletar:
    MsgBox "Erro ao deletar todos os dados da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - deletarDadosTabela - Módulo: m_TableFunctions"

End Function

Function deletarDadosTabela(ByVal NOME_TABELA As String)
'
' Deleta todos os dados de uma Tabela
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo ErroDeletar
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    limparFiltroTabela NOME_TABELA
    If TABELA.ListRows.count > 0 Then
        TABELA.DataBodyRange.Delete
    End If
       
Exit Function
ErroDeletar:
    MsgBox "Erro ao deletar todos os dados da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - deletarDadosTabela - Módulo: m_TableFunctions"

End Function

Function inserirNovaColuna(ByVal NOME_TABELA As String, _
                           ByVal NOME_COLUNA As String) As Integer
'
' Insere uma coluna na tabela
'
Dim TABELA As ListObject

    Application.ScreenUpdating = False
    On Error GoTo ErroInserir
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    If checkColumnExists(NOME_TABELA, NOME_COLUNA) = False Then
        TABELA.ListColumns.Add.Name = NOME_COLUNA
    End If
    inserirNovaColuna = retornarNumeroColuna(NOME_TABELA, NOME_COLUNA)
    
Exit Function
ErroInserir:
    inserirNovaColuna = 0
    MsgBox "Erro ao inserir a coluna: '" & NOME_COLUNA & "' na tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirColuna - Módulo: m_TableFunctions"
End Function

Function inserirFormulaColuna(ByVal NOME_TABELA As String, _
                                    ByVal NOME_COLUNA As String, _
                                    ByVal FORMULA_ As String) As Boolean
'
' Insere uma fórmula (string) em um determinada coluna
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroInserir
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)
    limparFiltroTabela NOME_TABELA
    TABELA.ListColumns(COLUMN_NUMBER).DataBodyRange.Value = FORMULA_
    
    inserirFormulaColuna = True
    ' Modo simples
    'Range(NOME_TABELA & "[" & NOME_COLUNA & "]").value = FORMULA_
        
Exit Function
ErroInserir:
    inserirFormulaColuna = False
    MsgBox "Erro ao inserir fórmula na coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirFormulaColuna - Módulo: m_TableFunctions"
End Function

Function inserirDadoTabela(ByVal NOME_TABELA As String, _
                           ByVal NOME_COLUNA As String, _
                           ByVal DADO As Variant, _
                           Optional LINHA_DESTINO As Integer = 0) As Integer
'
' Insere dados em uma tabela
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
Dim NEW_ROW As Integer
Dim FIELD_NAME As Variant

    Application.ScreenUpdating = False
    On Error GoTo ErroInserir
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    If NOME_COLUNA <> "" Then
        COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)
    Else
        COLUMN_NUMBER = 1
    End If
    
    If Not LINHA_DESTINO > 0 Then
        NEW_ROW = inserirNovaLinha(NOME_TABELA)
    Else
        NEW_ROW = LINHA_DESTINO
    End If
    
    ' Retorna a primeira linha inserida
    inserirDadoTabela = NEW_ROW
    
    If IsArray(DADO) Then
        For Each FIELD_NAME In DADO
            TABELA.Range.Cells(NEW_ROW, COLUMN_NUMBER).Value = FIELD_NAME
            'NEW_ROW = inserirNovaLinha(NOME_TABELA)
            NEW_ROW = NEW_ROW + 1
        Next
    Else
        TABELA.Range.Cells(NEW_ROW, COLUMN_NUMBER).Value = DADO
    End If
       
Exit Function
ErroInserir:
    MsgBox "Erro ao inserir dados na coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirDadoTabela - Módulo: m_TableFunctions"
End Function

Function activateTable(ByVal NOME_TABELA As String) As Boolean
'
' Ativa a planilha da tabela para seleção
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo ErroActive
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)

    If ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible <> xlSheetVisible Then
        ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible = xlSheetVisible
    End If
    ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Activate
    
    activateTable = True
    Exit Function

ErroActive:
    activateTable = False
    
End Function

Function deactivateTable(ByVal NOME_TABELA As String) As Boolean
'
' Desativa a planilha da tabela para seleção
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo ErroActive
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)

    If ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible = xlSheetVisible Then
        ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible = xlVeryHidden
    End If
    
    deactivateTable = True
    Exit Function

ErroActive:
    deactivateTable = False
    
End Function

' -------------------------------------------------------
' Funções de Verificação
' -------------------------------------------------------

Function verificaTabelaExiste(ByVal NOME_TABELA As String) As Boolean
'
' Inserir nova linha em tabela
'
Dim TABELA As ListObject
    Application.ScreenUpdating = False
    On Error GoTo Erro
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    verificaTabelaExiste = True
Exit Function

Erro:
    verificaTabelaExiste = True
End Function


' -------------------------------------------------------
' Funções Privadas
' -------------------------------------------------------
'
' @TODO - Funções privadas devem retornar TRUE ou FALSE

Private Function inserirNovaLinha(ByVal NOME_TABELA As String, _
                                Optional POSITION_ As Integer, _
                                Optional ALWAYS_INSERT_ As Boolean = True) As Integer
'
' Inserir nova linha em tabela
'
Dim TABELA As ListObject
Dim NEW_ROW As Object
    Application.ScreenUpdating = False
    On Error GoTo ErroInserir
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
        On Error GoTo SemDados
            ' Ativa Planilhapara usar o Select
            activateTable NOME_TABELA
            
        TABELA.DataBodyRange.Select
        
        If POSITION_ Then
            Set NEW_ROW = TABELA.ListRows.Add(POSITION:=POSITION_, AlwaysInsert:=ALWAYS_INSERT_)
        Else
            Set NEW_ROW = TABELA.ListRows.Add(AlwaysInsert:=ALWAYS_INSERT_)
        End If
        inserirNovaLinha = NEW_ROW.Range.Row
Exit Function

SemDados:
    inserirNovaLinha = TABELA.Range.Rows.count
    Exit Function
ErroInserir:
    MsgBox "Erro ao inserir nova linha na tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirNovaLInha - Módulo: m_TableFunctions"
End Function

Private Function addSortFieldToTable(ByRef TABELA As ListObject, _
                                     ByRef NOME_TABELA As String, _
                                     ByRef NOME_CAMPO As String)
'
' Adiciona Filtro à uma tabela
'
    TABELA.Sort.SortFields.Add Key:=Range(NOME_TABELA & "[[#All],[" & NOME_CAMPO & "]]") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
End Function

Private Function checkColumnExists(ByVal NOME_TABELA As String, _
                                   ByVal NOME_COLUNA As String) As Boolean
'
' Verifica se determinada columa existe em uma tabela
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
    Application.ScreenUpdating = False
    On Error GoTo ErroInserir
        Set TABELA = ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).ListObjects(NOME_TABELA)
    
    On Error GoTo ColunaInexistente
        COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)
    
    checkColumnExists = True
        
Exit Function
ErroInserir:
    checkColumnExists = False
    Exit Function
ColunaInexistente:
    Err.Clear
    checkColumnExists = False
    Exit Function
End Function


