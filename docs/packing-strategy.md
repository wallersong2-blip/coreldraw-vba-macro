# Estratégia de empacotamento (análise antes da versão final)

## 1) Objetivo de otimização

Queremos maximizar a ocupação útil do retângulo-alvo mantendo:

- sem sobreposição entre peças;
- folga mínima de **0,5 mm** entre quaisquer duas peças;
- possibilidade de rotação em **180°** por peça;
- priorização de um nome principal (Nome A) e preenchimento de sobras com Nome B.

## 2) Modelo geométrico recomendado

### 2.1 Representação de peça

Para cada nome convertido em curvas no CorelDRAW:

1. calcular bounding box local;
2. criar caixa inflada por `gap/2` em cada lado (com `gap = 0,5 mm`);
3. usar esta caixa inflada nas colisões rápidas;
4. opcionalmente validar com interseção mais precisa (outline) só no candidato final.

> Motivo: colisão por caixa inflada é muito mais rápida e já embute a distância mínima.

### 2.2 Obstáculos

- Obstáculos estáticos: formas já existentes no documento e bordas internas do retângulo-alvo.
- Obstáculos dinâmicos: peças aceitas durante o algoritmo.

Toda nova peça só entra se não colidir com ambos os grupos.

## 3) Estratégia de busca de posição

## 3.1 Grade adaptativa (coarse-to-fine)

Em vez de varrer só uma grade fixa, usar 2 estágios:

1. **Coarse**: passo maior (ex.: 1,0 mm) para localizar regiões promissoras rapidamente.
2. **Fine**: refinamento local (ex.: 0,2 mm) nos melhores candidatos.

Isso reduz tempo e melhora aproveitamento dos vazios.

### 3.2 Candidatos por rotação

Para cada posição `(x,y)` testada:

- testar peça em `0°`;
- testar peça em `180°`;
- guardar o melhor score viável.

Mesmo com mesmo bounding box, a forma real pode encaixar melhor em 180° por causa de recortes internos.

### 3.3 Ordem de inserção

Heurística sugerida:

- primeiro inserir as peças “mais difíceis” (maior área ou maior dimensão);
- depois peças menores para preencher lacunas.

No caso de nomes iguais repetidos, essa ordem vale quando houver escalas diferentes.

## 4) Função de score (qual candidato é melhor)

Score por candidato viável:

1. **Contato útil**: favorece peça mais próxima de obstáculos/bordas sem violar 0,5 mm.
2. **Penalidade de ilhas**: evita criar vazios estreitos não utilizáveis.
3. **Compactação**: prioriza manter cluster denso para liberar área contínua restante.

Um score simples inicial:

`score = w1*contato - w2*vazioResidual - w3*espalhamento`

com pesos calibráveis em `modConfig.bas`.

## 5) Fase dupla (Nome A e Nome B)

## 5.1 Fase 1: Nome A

- executar até limite de tentativas ou saturação de espaço;
- registrar cada peça aceita em lista de ocupação.

## 5.2 Fase 2: Nome B

- reutilizar exatamente a mesma lista de ocupação como obstáculo;
- testar apenas regiões remanescentes;
- parar quando não houver candidato viável.

## 6) Critérios de parada

Parar quando ocorrer um dos eventos:

- `N` iterações sem melhoria de ocupação;
- tempo máximo de execução;
- nenhum candidato válido no refinamento fine.

## 7) Métricas para validar melhoria real

Medir em cada execução:

- taxa de ocupação (% área ocupada / área disponível);
- menor distância entre peças (deve ser >= 0,5 mm);
- quantidade de peças Nome A e Nome B;
- tempo total de processamento.

Sem essas métricas, não há como comprovar otimização.

## 8) Plano de implementação incremental

1. Base geométrica + colisão por caixa inflada.
2. Inserção simples em grade + rotação 0/180.
3. Score de compactação.
4. Busca coarse-to-fine.
5. Fase dupla A/B com métricas em log.

Essa sequência reduz risco e permite validar ganho a cada etapa.
