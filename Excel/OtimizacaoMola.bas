Option Explicit
'==============================================================================
' Otimizacao de mola de compressao (estampo/molde) com Excel Solver
' + atualizacao automatica do modelo SolidWorks (massa real + cor por forca)
'
' PRE-REQUISITOS (fazer uma vez so):
'   1) Excel > Arquivo > Opcoes > Suplementos > Gerenciar: Suplementos do Excel
'      > Ir... > marcar "Solver" > OK.
'   2) No editor VBA: Ferramentas > Referencias > marcar "SOLVER"
'      (se nao aparecer, o passo 1 nao foi feito ou o Excel precisa reiniciar).
'   3) No SolidWorks, a peca da mola deve ter as Equations com as variaveis
'      globais d_wire, D_mean, Na e L_free vinculadas a uma faixa da planilha
'      (Equations > Link Values to External File), apontando para este
'      mesmo arquivo .xlsx, na faixa nomeada "SW_LinkedValues".
'   4) Ajuste a constante CAMINHO_PECA abaixo para o caminho real do .SLDPRT.
'
' LAYOUT DE NOMES QUE A PLANILHA PRECISA TER (Nome Definido no Excel):
'   Variaveis do Solver (== mesmas celulas ligadas ao SolidWorks):
'     d_wire  -> diametro do arame (mm)
'     D_mean  -> diametro medio da espira (mm)
'     Na      -> numero de espiras ativas
'   (os tres nomes sao IDENTICOS aos das Global Variables do SolidWorks,
'    de proposito -- uma unica fonte de verdade, sem camada de traducao.)
'
'   Entradas de projeto:
'     in_Fwork     -> forca exigida na deflexao de trabalho (N)
'     in_swork     -> deflexao de trabalho (mm)
'     in_Lfree     -> comprimento livre disponivel no alojamento (mm)
'     in_Dhole     -> diametro do furo/alojamento (mm)
'     in_SF        -> fator de seguranca sobre a tensao admissivel
'     in_folga     -> folga minima entre espiras no solido (mm)
'
'   Propriedades do material (devem bater com o material atribuido no SW):
'     mat_rho      -> densidade (g/mm3)
'     mat_G        -> modulo de cisalhamento (MPa = N/mm2)
'     mat_tauadm   -> tensao admissivel de fadiga do material (MPa)
'
'   Celulas calculadas (formulas na planilha, no Excel mesmo, nao em VBA):
'     chk_C        -> D_mean / d_wire (indice da mola)
'     chk_Fcalc    -> forca calculada em s_work
'     chk_tau      -> tensao de cisalhamento calculada (com fator de Wahl)
'     chk_Dout     -> D_mean + d_wire
'     chk_Lsolid   -> Nt * d_wire   (Nt = Na + 2)
'     obj_Mass     -> massa calculada (rho * pi^2/4 * d_wire^2 * D_mean * Nt)
'
'   Limites auxiliares (facilita ler no SolverAdd):
'     lim_tau      -> = mat_tauadm / in_SF
'     lim_Lsolid   -> = in_Lfree - in_swork - in_folga
'     lim_Cmin, lim_Cmax -> 4 e 12 (faixa recomendada do indice da mola)
'     bnd_dwire_min, bnd_dwire_max, bnd_Dmean_min, bnd_Dmean_max,
'     bnd_Na_min, bnd_Na_max -> limites praticos das variaveis (evita o
'       Solver testar valores absurdos). Nomes com 'wire'/'mean' por extenso
'       de proposito -- o Excel nao diferencia maiusculas de minusculas em
'       nomes definidos, entao 'bnd_dmin' e 'bnd_Dmin' seriam O MESMO nome.
'
'   Saidas escritas por esta macro:
'     out_MassaReal   -> massa lida de volta do SolidWorks (g)
'     out_ClasseForca -> "Leve" / "Media" / "Pesada" / "ExtraPesada"
'==============================================================================

Public Const CAMINHO_PECA As String = "C:\CaminhoDoSeuProjeto\Mola.SLDPRT" ' AJUSTE AQUI

'------------------------------------------------------------------------------
' Ponto de entrada: roda o Solver e depois sincroniza com o SolidWorks
'------------------------------------------------------------------------------
Sub Main_OtimizarEAtualizar()
    If Not ConfigurarEResolverSolver() Then
        MsgBox "O Solver nao encontrou uma solucao viavel." & vbCrLf & _
               "Revise os limites das variaveis (bnd_*) ou as restricoes de projeto.", _
               vbExclamation, "Otimizacao da mola"
        Exit Sub
    End If

    AtualizarSolidWorksEClassificar

    MsgBox "Otimizacao concluida e modelo do SolidWorks atualizado." & vbCrLf & _
           "Massa (Excel): " & Round(Range("obj_Mass").Value, 3) & " g" & vbCrLf & _
           "Massa (SolidWorks): " & Round(Range("out_MassaReal").Value, 3) & " g" & vbCrLf & _
           "Classe de forca: " & Range("out_ClasseForca").Value, _
           vbInformation, "Otimizacao da mola"
End Sub

'------------------------------------------------------------------------------
' Monta o modelo do Solver e resolve. Retorna True se encontrou solucao.
'------------------------------------------------------------------------------
Function ConfigurarEResolverSolver() As Boolean
    On Error GoTo falha

    SolverReset

    SolverOk SetCell:=Range("obj_Mass").Address(External:=True), _
             MaxMinVal:=2, _                                           ' 2 = minimizar
             ValueOf:=0, _
             ByChange:=Range("d_wire,D_mean,Na").Address(External:=True), _
             Engine:=1, EngineDesc:="GRG Nonlinear"                     ' motor nao-linear (formulas com potencias)

    ' --- Restricoes de projeto ---
    SolverAdd CellRef:=Range("chk_Fcalc").Address(External:=True), Relation:=3, _
              FormulaText:=Range("in_Fwork").Address(External:=True)          ' Fcalc >= Fwork
    SolverAdd CellRef:=Range("chk_tau").Address(External:=True), Relation:=1, _
              FormulaText:=Range("lim_tau").Address(External:=True)           ' tau <= tau_adm/SF
    SolverAdd CellRef:=Range("chk_Dout").Address(External:=True), Relation:=1, _
              FormulaText:=Range("in_Dhole").Address(External:=True)          ' Dout <= Dhole
    SolverAdd CellRef:=Range("chk_Lsolid").Address(External:=True), Relation:=1, _
              FormulaText:=Range("lim_Lsolid").Address(External:=True)        ' Lsolid <= Lfree - swork - folga
    SolverAdd CellRef:=Range("chk_C").Address(External:=True), Relation:=3, _
              FormulaText:=Range("lim_Cmin").Address(External:=True)          ' C >= 4
    SolverAdd CellRef:=Range("chk_C").Address(External:=True), Relation:=1, _
              FormulaText:=Range("lim_Cmax").Address(External:=True)          ' C <= 12

    ' --- Limites praticos das variaveis (evita valores fisicamente absurdos) ---
    SolverAdd CellRef:=Range("d_wire").Address(External:=True), Relation:=3, _
              FormulaText:=Range("bnd_dwire_min").Address(External:=True)
    SolverAdd CellRef:=Range("d_wire").Address(External:=True), Relation:=1, _
              FormulaText:=Range("bnd_dwire_max").Address(External:=True)
    SolverAdd CellRef:=Range("D_mean").Address(External:=True), Relation:=3, _
              FormulaText:=Range("bnd_Dmean_min").Address(External:=True)
    SolverAdd CellRef:=Range("D_mean").Address(External:=True), Relation:=1, _
              FormulaText:=Range("bnd_Dmean_max").Address(External:=True)
    SolverAdd CellRef:=Range("Na").Address(External:=True), Relation:=3, _
              FormulaText:=Range("bnd_Na_min").Address(External:=True)
    SolverAdd CellRef:=Range("Na").Address(External:=True), Relation:=1, _
              FormulaText:=Range("bnd_Na_max").Address(External:=True)

    SolverOptions Precision:=0.000001, Convergence:=0.0001, MaxTime:=60, Iterations:=200

    ' UserFinish:=True aplica a solucao direto, sem abrir a caixa de dialogo
    ' de resultados do Solver. Na primeira vez, rode com False para conferir
    ' o resultado manualmente antes de confiar na automacao completa.
    Dim resultado As Long
    resultado = SolverSolve(UserFinish:=True)

    ' Codigos de retorno do Solver: 0 e 1 indicam que uma solucao foi
    ' encontrada dentro das condicoes de convergencia/celulas atendidas.
    ' Consulte a ajuda do Solver (F1 com o cursor em SolverSolve) para a
    ' lista completa, e confirme observando a caixa de dialogo na primeira
    ' execucao manual (UserFinish:=False).
    ConfigurarEResolverSolver = (resultado = 0 Or resultado = 1)
    Exit Function

falha:
    MsgBox "Erro ao configurar o Solver: " & Err.Description & vbCrLf & _
           "Confira se o suplemento Solver esta ativo e se a referencia " & _
           "'SOLVER' esta marcada em Ferramentas > Referencias no editor VBA.", _
           vbCritical, "Otimizacao da mola"
    ConfigurarEResolverSolver = False
End Function

'------------------------------------------------------------------------------
' Salva a planilha, forca o SolidWorks a reler os valores vinculados,
' le a massa real (considerando o material atribuido na peca) e aplica
' a cor conforme a classe de forca.
'------------------------------------------------------------------------------
Sub AtualizarSolidWorksEClassificar()
    Dim swApp As Object, swModel As Object, swExt As Object, swMass As Object
    Dim longStatus As Long, longWarnings As Long

    ' O SolidWorks so relê os valores vinculados se o arquivo estiver salvo em disco.
    ThisWorkbook.Save

    ' Tenta usar uma instancia do SolidWorks ja aberta; senao, abre uma nova.
    On Error Resume Next
    Set swApp = GetObject(, "SldWorks.Application")
    On Error GoTo 0
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
        swApp.Visible = True
    End If

    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        Set swModel = swApp.OpenDoc6(CAMINHO_PECA, 1, 0, "", longStatus, longWarnings) ' 1 = swDocPART
    ElseIf LCase(swModel.GetPathName) <> LCase(CAMINHO_PECA) Then
        Set swModel = swApp.OpenDoc6(CAMINHO_PECA, 1, 0, "", longStatus, longWarnings)
    End If

    If swModel Is Nothing Then
        MsgBox "Nao foi possivel abrir a peca da mola." & vbCrLf & _
               "Confira a constante CAMINHO_PECA no topo do modulo.", _
               vbCritical, "Otimizacao da mola"
        Exit Sub
    End If

    ' Forca a reconstrucao: e aqui que o SolidWorks vai reler os valores
    ' vinculados da planilha (Equations > Link Values to External File)
    ' e recalcular a geometria da mola.
    swModel.ForceRebuild3 False

    ' Le a massa real considerando o material atribuido na peca.
    Set swExt = swModel.Extension
    Set swMass = swExt.CreateMassProperty
    Range("out_MassaReal").Value = swMass.Mass * 1000   ' kg -> g (ajuste se sua planilha usar outra unidade)

    ' Classifica pela forca calculada e aplica a cor correspondente.
    Dim classe As String
    classe = ClassificarClasseDeForca(Range("chk_Fcalc").Value)
    Range("out_ClasseForca").Value = classe

    AplicarCorNoModelo swModel, classe
End Sub

'------------------------------------------------------------------------------
' Classifica a mola por faixa de forca (aproximacao didatica).
'
' ATENCAO: a classificacao oficial ISO 10243 compara a mola com o catalogo
' do fabricante PARA O MESMO diametro e comprimento livre -- nao existe um
' valor de forca absoluto que separe as classes universalmente. Os limites
' abaixo (lim_leve_max, lim_media_max, lim_pesada_max) sao definidos por
' voce na planilha, como uma aproximacao para fins didaticos do video.
' Em um projeto real de producao, confirme a classe no catalogo do
' fabricante escolhido.
'------------------------------------------------------------------------------
Function ClassificarClasseDeForca(ByVal Forca As Double) As String
    Select Case Forca
        Case Is < Range("lim_leve_max").Value:   ClassificarClasseDeForca = "Leve"
        Case Is < Range("lim_media_max").Value:  ClassificarClasseDeForca = "Media"
        Case Is < Range("lim_pesada_max").Value: ClassificarClasseDeForca = "Pesada"
        Case Else:                                ClassificarClasseDeForca = "ExtraPesada"
    End Select
End Function

'------------------------------------------------------------------------------
' Aplica a cor no modelo do SolidWorks conforme a classe de forca,
' seguindo a convencao ISO 10243 para molas de estampo:
'   Leve = verde | Media = azul | Pesada = vermelho | ExtraPesada = amarelo
'
' *** ATENCAO IMPORTANTE ***
' Nao foi possivel testar esta chamada dentro do SolidWorks. A forma mais
' segura de confirmar a linha exata para a SUA versao e:
'   1) No SolidWorks: Ferramentas > Macro > Gravar.
'   2) Mude a cor da peca manualmente uma vez (clique direito > Aparencia > Cor).
'   3) Pare a gravacao e abra o .swp gerado.
'   4) Copie a linha que muda a cor para dentro deste Sub, substituindo os
'      valores fixos de RGB por vColor(0), vColor(1), vColor(2).
' A chamada abaixo e a forma classica (legada) de setar uma cor simples de
' peca inteira e costuma funcionar, mas pode variar conforme a versao:
'------------------------------------------------------------------------------
Sub AplicarCorNoModelo(swModel As Object, ByVal classe As String)
    Dim vColor(2) As Double

    Select Case classe
        Case "Leve":        vColor(0) = 0#:    vColor(1) = 0.6:  vColor(2) = 0#     ' verde
        Case "Media":       vColor(0) = 0#:    vColor(1) = 0.3:  vColor(2) = 0.8    ' azul
        Case "Pesada":      vColor(0) = 0.8:   vColor(1) = 0#:   vColor(2) = 0#     ' vermelho
        Case "ExtraPesada": vColor(0) = 1#:    vColor(1) = 0.85: vColor(2) = 0#     ' amarelo
    End Select

    On Error GoTo semCor
    swModel.SetMaterialPropertyValues2 vColor, 1   ' 1 = swThisConfiguration
    swModel.ForceRebuild3 False
    Exit Sub

semCor:
    MsgBox "Nao foi possivel aplicar a cor automaticamente (" & Err.Description & ")." & vbCrLf & _
           "Grave uma macro no SolidWorks mudando a cor manualmente uma vez e " & _
           "substitua a linha 'swModel.SetMaterialPropertyValues2' por aquela " & _
           "gerada, dentro do Sub AplicarCorNoModelo.", vbExclamation, "Otimizacao da mola"
End Sub
