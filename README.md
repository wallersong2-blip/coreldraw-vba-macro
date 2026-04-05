# Macro VBA para encaixe de nomes no CorelDRAW

Primeira versão funcional (VBA clássico) para distribuir shapes selecionados dentro de um retângulo selecionado.

## O que esta versão faz

- identifica o **retângulo-alvo** na seleção (retângulo de maior área);
- trata os demais shapes selecionados como **peças**;
- testa rotação `0°` e `180°` para cada peça;
- aplica gap mínimo de `0,5 mm` por colisão de **bounding box inflada**;
- tenta posicionar o máximo possível dentro do retângulo;
- expõe ponto de entrada simples: `ExecutePlaceSelectedShapes`.

## Estrutura de módulos

- `src/modules/modConfig.bas`: constantes de gap, grade e pesos.
- `src/modules/modGeometry.bas`: tipos e funções geométricas (RectMM, interseções, move por centro).
- `src/modules/modPackingEngine.bas`: busca de candidatos em grade e validação de colisão.
- `src/modules/modCorelAdapter.bas`: acesso ao `ActiveSelectionRange` e identificação de alvo/peças.
- `src/modules/modPlacementFlow.bas`: fluxo principal para posicionamento.

## Limite desta etapa

Esta é uma base funcional simples para encaixe por bounding box. Melhorias de otimização fina (vazios complexos, score mais avançado e fase de sobras Nome B) entram na próxima iteração.
