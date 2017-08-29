Attribute VB_Name = "mod_TableFuncions"
' -------------------------------------------------------
'
' M�dulo Table Functions
'
' Possui diversas fun��es para trabalhar com tabelas
' pr�-formatadas no Excel
'
'
' Vers�o: 0.4
'
' �ltima atualiza��o: 29/08/2017
'
' -------------------------------------------------------
'
' Lista de fun��es:
'
' limparFiltroTabela
' filtrarTabelaPorCampo
' retornarNumeroColuna
' retornarQtdLinhasTabela
' retornarQtdLinhasColuna
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
'
'
' Fun��es Privadas:
'
' addSortFieldToTable
' inserirNovaLinha
' checkColumnExists

' -------------------------------------------------------

Option Explicit

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
        "Provavelmente foi exclu�da indevidamente", vbCritical, "Erro - limparFiltroTabela - mod_TableFunctions"

End Function

Function ordenarTabelaPorCampo(ByRef NOME_TABELA As String, _
                                ByRef NOME_CAMPO_S As Variant)
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
            ' Fun��o privada dentro desse m�dulo
            addSortFieldToTable TABELA, NOME_TABELA, CStr(FIELD_NAME)
        Next
    Else
        ' Fun��o privada dentro desse m�dulo
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
        "Provavelmente foi exclu�da indevidamente", vbCritical, "Erro - ordenarTabelaPorCampo - mod_TableFunctions"

End Function

Function filtrarTabelaPorCampo(ByRef NOME_TABELA As String, _
                                ByRef NOME_CAMPO As String, _
                                ByRef TEXTO_FILTRO As String, _
                                Optional ByRef LIMPA_FILTRO As Boolean)
'
' Ordena Tabela por coluna
'
Dim TABELA As ListObject
Dim COLUMN_NUMBER As Integer
Dim PRETECTED As Boolean
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
    MsgBox "Erro ao filtrar a tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - filtrarTabelaPorCampo - mod_TableFunctions"

End Function

Function retornarNumeroColuna(ByVal NOME_TABELA As String, _
                                    ByVal NOME_COLUNA As String) As Integer
'
' Retorna o n�mero da coluna de um determinado campo (cabe�alho) de uma tabela
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
    MsgBox "Erro buscar n�mero da coluna na tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - retornarNumeroColuna - mod_TableFunctions"

End Function

Function retornarQtdLinhasTabela(ByVal NOME_TABELA As String, _
                                 Optional ByVal CONSIDERA_FILTRO As Boolean = True) As Integer
'
' Retorna a quantidade de linhas de uma tabela, levando em considera��o o filtro atual aplicado
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
    MsgBox "Erro ao buscar a quantidade de linhas da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - retornarQtdLinhasTabela - mod_TableFunctions"

End Function

Function retornarQtdLinhasColuna(ByVal NOME_TABELA As String, _
                                 ByVal NOME_COLUNA As String, _
                                 Optional ByVal SOMENTE_NAO_VAZIAS As Boolean = True) As Integer
'
' Retorna a quantidade de linhas de uma tabela, levando em considera��o o filtro atual aplicado
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
    MsgBox "Erro ao buscar a quantidade de linhas da coluna: '" & NOME_COLUNA & "da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - retornarQtdLinhasColuna - mod_TableFunctions"

End Function

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

'    If ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible <> xlSheetVisible Then
'        ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Visible = xlSheetVisible
'    End If
'    ActiveWorkbook.Worksheets(Range(NOME_TABELA).Parent.Name).Activate
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
    MsgBox "Erro ao selecionar os dados da coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - selecionarDadosDaColuna - mod_TableFunctions"
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
    MsgBox "Erro ao selecionar os dados da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - selecionarDadosDaTabela - mod_TableFunctions"
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
    MsgBox "Erro ao selecionar a linha: '" & LINHA_TABELA & "'" & " da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - selecionarLinhaDaTabela - mod_TableFunctions"
    selecionarLinhaDaTabela = False
    Exit Function
End Function

Function converterTextoEmNumero(ByVal NOME_TABELA As String, _
                                    ByVal NOME_COLUNA As String)
'
' Converte os dados de determinada Coluna em determinada Tabela em n�mero
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
    MsgBox "Erro buscar converter os dados da coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - converterTextoEmNumero - mod_TableFunctions"

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
    MsgBox "Erro ao deletar todos os dados da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - deletarDadosTabela - mod_TableFunctions"

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
    MsgBox "Erro ao inserir a coluna: '" & NOME_COLUNA & "' na tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirColuna - mod_TableFunctions"
End Function

Function inserirFormulaColuna(ByVal NOME_TABELA As String, _
                                    ByVal NOME_COLUNA As String, _
                                    ByVal FORMULA_ As String) As Boolean
'
' Insere uma f�rmula (string) em um determinada coluna
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
    MsgBox "Erro ao inserir f�rmula na coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirFormulaColuna - mod_TableFunctions"
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
    COLUMN_NUMBER = Application.WorksheetFunction.Match(NOME_COLUNA, TABELA.HeaderRowRange, 0)
    
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
    MsgBox "Erro ao inserir dados na coluna: '" & NOME_COLUNA & "' da tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirDadoTabela - mod_TableFunctions"
End Function

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

Function activateTable(ByVal NOME_TABELA As String) As Boolean
'
' Ativa a planilha da tabela para sele��o
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
' Desativa a planilha da tabela para sele��o
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
' Fun��es Privadas
' -------------------------------------------------------
'
' @TODO - Fun��es privadas devem retornar TRUE ou FALSE

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
            ' Ativa Planilhapara usar o DataBodyRange
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
    MsgBox "Erro ao inserir nova linha na tabela: '" & NOME_TABELA & "'", vbCritical, "Erro - inserirNovaLInha - mod_TableFunctions"
End Function

Private Function addSortFieldToTable(ByRef TABELA As ListObject, _
                                     ByRef NOME_TABELA As String, _
                                     ByRef NOME_CAMPO As String)
'
' Adiciona Filtro � uma tabela
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


