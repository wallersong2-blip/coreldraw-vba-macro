# README de teste rápido (CorelDRAW)

## Importação dos módulos no VBA (1 vez por arquivo `.cdr`)

1. No CorelDRAW, abra: **Ferramentas > Macros > Editor de Macros VBA**.
2. No editor VBA: **File > Import File...**.
3. Importe os módulos desta pasta, nesta ordem sugerida:
   - `src/modules/modConfig.bas`
   - `src/modules/modGeometry.bas`
   - `src/modules/modCorelAdapter.bas`
   - `src/modules/modPackingEngine.bas`
   - `src/modules/modPlacementFlow.bas`
4. Salve o projeto VBA.

## Modo 1 fase — `ExecutePlaceSelectedShapes`

1. Selecione **1 retângulo-alvo** + peças que deseja posicionar.
2. Execute a macro `ExecutePlaceSelectedShapes`.

## Modo 2 fases — `ExecutePlaceSelectedShapesTwoPhase`

1. Selecione **junto**: **retângulo-alvo + todas as peças** (fase 1 e fase 2).
2. Defina as peças da fase 2 com `Name` iniciando em `F2_`.
   - Ex.: `F2_NOME_B_01`, `F2_NOME_B_02`.
3. Execute a macro `ExecutePlaceSelectedShapesTwoPhase`.

## Regra de funcionamento no modo 2 fases

- Peças sem prefixo `F2_` → fase 1.
- Peças com prefixo `F2_` → fase 2.
- Após posicionar a fase 1, ela vira **obstáculo fixo** para o posicionamento da fase 2.

## Checklist de teste real

- O editor VBA importa os 5 módulos sem erro de compilação.
- Com seleção válida, a macro executa sem mensagem de erro de runtime.
- No modo 1 fase, as peças ficam dentro do retângulo-alvo sem sobreposição visível.
- No modo 2 fases, apenas shapes com `F2_` entram na fase 2.
- A fase 2 não invade peças já colocadas na fase 1.
- Rotações observadas continuam em `0°`/`180°`.
- Gap mínimo visual de `0,5 mm` é mantido entre peças.
