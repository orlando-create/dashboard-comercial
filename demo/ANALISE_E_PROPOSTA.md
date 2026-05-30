# Dashboard Comercial 2026 — Análise + Proposta + DEMO

**Data:** 16/05/2026 · **Versão:** DEMO 1.0 · **Autor:** Claude (Cowork) · **Solicitante:** Zagatto

---

## 1. Diagnóstico da base atual

A planilha `forescast atualizado 2026.xlsx` é robusta mas mistura **dados transacionais, agregações e relatórios** numa única pasta de trabalho. Mapeei 15 abas, três famílias distintas:

| Família | Abas | Função | Estado |
|---|---|---|---|
| Transacional | `FORECAST 2026`, `Procurações Jan…Maio` | Origem dos dados | Boa, mas com inconsistências |
| Agregação manual | `JAN 26`, `Fev 26 `, `MAR 26` | Tabelas dinâmicas estáticas | Redundante — Power BI elimina |
| Relatório | `ONE PAGE JAN26…MAIO26`, `META` | Outputs visuais já compilados | Substituível pelo dashboard |

**Volume real:** 312 registros forecast · 15 hunters com meta · 268 procurações.

### Problemas que afetam a confiabilidade dos KPIs

1. **Inconsistência de domínios** (afeta segmentação e filtros):
   - `MES ASSINADO`: aparece `"ABRI"` no lugar de `"ABR"` em ~19 linhas.
   - `PRODUTO`: `"PREVIDENCIARIO "` (com espaço) e `"PREVIDENCIARIO"` contam como produtos diferentes.
   - `REGIÃO`: 76 cidades distintas, mas existem pares como `"FORTALEZA - CE"` vs `"FORTALEZA- CE"` (sem espaço) que duplicam.
   - `PARCEIRO`: variações de Somativa/Aragão (`"Somativa/ARAGAO"`, `"Somativa/ARAGAOALBANO & ALBANO…"`, `"ARAGÃO"`).
   - `HUNTER`: divergência entre abas — `"NATALIA"` na META vs `"NATHALIA"` no Forecast/Procurações; `"KATARINY"` na META vs `"KATERYNI"` em Procurações.

2. **Procurações fragmentadas em 5 abas** com layouts diferentes (Jan/Fev começam na coluna B, Mar/Abr/Mai começam na coluna A; ranges de dados começam em linhas diferentes — 2, 3, 4). Isso forçou o script atual a manter um dicionário fixo de offsets — qualquer alteração na planilha quebra a leitura.

3. **`ARTHUR` aparece no Forecast (2 negócios) sem meta na aba META** — o atingimento dele fica indefinido.

4. **Meta global de R$ 34 Mi não está na planilha** — está só no contexto do briefing. A soma das metas individuais por hunter dá **R$ 38,78 Mi**, então existe um "buffer" de R$ 4,78 Mi entre o agregado bottom-up dos hunters e o teto top-down da diretoria.

5. **Não há campo de data real** — só os campos `MES` (entrada) e `MES ASSINADO` (fechamento). Isso impossibilita análises diárias/semanais, cálculo correto de tempo médio em dias e análise de aging de pipeline.

6. **Não há identificador único do lead/conta**. Hoje a junção é por `EMPRESA` (string livre), o que é frágil (case, espaços, abreviações).

7. **`AGENDA` (reuniões NOVAS/REPETIDAS)** existe nas abas ONE PAGE mensais, mas **não está estruturada como dado transacional** — está só nos totalizadores. Logo, não dá pra cruzar com hunter, fechamento ou tempo médio.

### Recomendação imediata (sem virar projeto)

Antes de subir o novo dashboard em produção, criar um **dicionário de domínios** num pequeno script de normalização (padronizar Hunter, Produto, Região, Parceiro, Mês). Posso fazer essa etapa como item subsequente.

---

## 2. Modelagem de dados proposta (Power BI · esquema estrela)

O modelo atual lê tudo de uma "tabela larga" e calcula no Python. Para Power BI executivo, o ideal é um esquema estrela com **dimensões pequenas** e **tabelas-fato magras**:

```
                          ┌───────────────────┐
                          │   d_Calendario    │  (mês, trimestre, ano)
                          └─────────┬─────────┘
                                    │
        ┌──────────────┐    ┌───────┴────────┐    ┌──────────────┐
        │  d_Hunter    │────│  f_Forecast    │────│  d_Produto   │
        │  (id, nome,  │    │  (uma linha    │    │              │
        │  perfil,     │    │  por opt.)     │    └──────────────┘
        │  status)     │    │                │
        └──────────────┘    └───┬────────┬───┘    ┌──────────────┐
                                │        │        │ d_Status     │
        ┌──────────────┐        │        └────────│ (Base, EmAss,│
        │  d_Empresa   │────────┘                 │  Assinado,   │
        │ (id, CNPJ,   │                          │  Standby)    │
        │  cidade, UF) │                          └──────────────┘
        └──────────────┘
                │
                │              ┌────────────────┐
                └──────────────│  f_Procuracao  │
                               │  (uma linha    │
                               │  por proc.)    │
                               └────────────────┘

  + f_Meta (Hunter × Mês × ValorMeta)
  + d_Temperatura (Quente/Morno/Frio)
  + d_Contrato (Novo/Upsell)
```

**Tabelas-fato:**

- `f_Forecast` — uma linha por oportunidade, com FKs e métricas (Faturamento, Crédito, Honorários).
- `f_Procuracao` — uma linha por procuração, com FK Hunter, FK Empresa, Mês, Tipo (Novo/Renovação), Data Assinatura, Validade.
- `f_Meta` — uma linha por Hunter × Mês × ValorMeta (em vez de pivotada como hoje).

**Por que importa:** este modelo permite criar medidas DAX que respondem a *qualquer* filtro de dimensão (ex.: "meta atingida em SP por Hunter Dimas em produtos previdenciários no T1") sem alterar nada na fonte.

---

## 3. Medidas DAX essenciais

Estas são as medidas que sustentariam o dashboard. Estão escritas em DAX pronto para colar no Power BI — basta criar uma tabela "[Medidas]" e ir colando.

### 3.1 Receita & Conversão

```dax
Receita Fechada =
CALCULATE (
    SUM ( f_Forecast[Faturamento] ),
    f_Forecast[Status] IN { "ASSINADO", "EM ASSINATURA" }
)

Pipeline Total =
CALCULATE (
    SUM ( f_Forecast[Faturamento] ),
    f_Forecast[Status] IN { "BASE", "EM ASSINATURA" }
)

Pipeline Base =
CALCULATE ( SUM ( f_Forecast[Faturamento] ), f_Forecast[Status] = "BASE" )

Valor Standby =
CALCULATE ( SUM ( f_Forecast[Faturamento] ), f_Forecast[Status] = "STANDBY" )

Contas Fechadas =
CALCULATE (
    DISTINCTCOUNT ( f_Forecast[EmpresaID] ),
    f_Forecast[Status] IN { "ASSINADO", "EM ASSINATURA" }
)

Contas Entradas Mês =
CALCULATE (
    DISTINCTCOUNT ( f_Forecast[EmpresaID] ),
    USERELATIONSHIP ( d_Calendario[Mes], f_Forecast[MesEntrada] )
)

Conversão por Contas =
DIVIDE ( [Contas Fechadas], [Contas Entradas Mês] )

Conversão por Valor =
DIVIDE ( [Receita Fechada], [Pipeline Total] + [Receita Fechada] )
```

### 3.2 Regra de Competência (chave do briefing)

```dax
Fechado Dentro Competência =
CALCULATE (
    SUM ( f_Forecast[Faturamento] ),
    f_Forecast[Status] IN { "ASSINADO", "EM ASSINATURA" },
    f_Forecast[MesEntrada] = f_Forecast[MesAssinado]
)

Fechado Fora Competência =
CALCULATE (
    SUM ( f_Forecast[Faturamento] ),
    f_Forecast[Status] IN { "ASSINADO", "EM ASSINATURA" },
    f_Forecast[MesEntrada] <> f_Forecast[MesAssinado],
    NOT ISBLANK ( f_Forecast[MesAssinado] )
)

% Aderência Competência =
DIVIDE (
    [Fechado Dentro Competência],
    [Fechado Dentro Competência] + [Fechado Fora Competência]
)

Tempo Médio Fechamento (meses) =
AVERAGEX (
    FILTER (
        f_Forecast,
        f_Forecast[Status] IN { "ASSINADO", "EM ASSINATURA" }
            && NOT ISBLANK ( f_Forecast[MesAssinado] )
    ),
    MOD (
        RELATED ( d_Cal_Fech[NumMes] ) - RELATED ( d_Cal_Entr[NumMes] ) + 12,
        12
    )
)
```

### 3.3 Meta — regra 34 Mi

```dax
Meta Global = 34000000

Meta YTD =
CALCULATE (
    SUM ( f_Meta[ValorMeta] ),
    d_Calendario[Data] <= MAX ( d_Calendario[Data] )
)

Atingimento YTD = DIVIDE ( [Receita Fechada], [Meta YTD] )

Atingimento Global = DIVIDE ( [Receita Fechada], [Meta Global] )

Excedente Acima 34Mi =
VAR R = [Receita Fechada]
RETURN IF ( R > 34000000, R - 34000000, 0 )

Faltante para 34Mi =
VAR R = [Receita Fechada]
RETURN IF ( R >= 34000000, 0, 34000000 - R )

Run-Rate Necessário =
VAR mesesRestantes = 12 - MONTH ( TODAY () ) + 1
RETURN DIVIDE ( [Faltante para 34Mi], mesesRestantes )

Projeção Linear =
DIVIDE ( [Receita Fechada], MONTH ( TODAY () ) ) * 12

Pipeline Necessário =
DIVIDE ( [Faltante para 34Mi], [Conversão por Valor] )
```

### 3.4 Procurações

```dax
Procurações Total = COUNTROWS ( f_Procuracao )

Procurações Novas =
CALCULATE ( COUNTROWS ( f_Procuracao ), f_Procuracao[Tipo] = "NOVO" )

Procurações Renovadas =
CALCULATE ( COUNTROWS ( f_Procuracao ), f_Procuracao[Tipo] = "RENOVAÇÃO" )

% Renovação = DIVIDE ( [Procurações Renovadas], [Procurações Total] )

Eficiência (R$ por Procuração) = DIVIDE ( [Receita Fechada], [Procurações Total] )
```

### 3.5 Pipeline & Qualidade

```dax
Saúde do Funil =
DIVIDE (
    [Receita Fechada] + [Pipeline Base],
    [Receita Fechada] + [Pipeline Base] + [Valor Standby]
)

% Quente =
DIVIDE (
    CALCULATE ( [Pipeline Total], f_Forecast[Temperatura] = "QUENTE" ),
    [Pipeline Total]
)

Risco Standby Alto Valor =
CALCULATE (
    COUNTROWS ( f_Forecast ),
    f_Forecast[Status] = "STANDBY",
    f_Forecast[Faturamento] >= 100000
)
```

---

## 4. KPIs que você ainda não acompanha (sugeridos)

| KPI | Descrição | Por que importa |
|---|---|---|
| **% Aderência Competência** | Fechado dentro / (dentro+fora) | Mostra previsibilidade do funil. Acima de 70% = ciclo saudável |
| **Saúde do Funil** | (Fechado+Base) / Total | Sinal antecipado de erosão antes que apareça na meta |
| **Eficiência R$/Procuração** | Receita Fechada / Procs | Mede capacidade real do hunter, não só esforço |
| **Run-Rate Necessário** | (34Mi - YTD) / meses restantes | Tira a meta abstrata e dá um número operacional |
| **Pipeline Necessário** | Faltante / Conversão atual | Quanto pipeline precisa ENTRAR para bater a meta |
| **Velocidade do Pipeline** | (Pipeline atual / Tempo médio fechamento) | Forecast realista de receita em 90/180 dias |
| **Risco Standby Alto Valor** | # contas em standby ≥ R$ 100k | Foco para recuperação prioritária |
| **Receita por Cliente Único** | Receita / DISTINCT empresas | Mostra cross-sell real (várias soluções na mesma conta) |
| **Velocity Score por Hunter** | Receita × Conversão / Ciclo | Métrica composta que reflete eficiência completa |
| **NPS de Pipeline** | % Quente / % Frio | Qualidade do que está entrando — leading indicator |

---

## 5. Arquitetura visual entregue na DEMO

A versão DEMO está estruturada em **6 páginas navegáveis por tabs** (não mais One Page rolável). Cada uma com propósito claro:

### Página 1 · **Executiva** (foco diretoria)
- **3 Hero KPIs** acima da dobra: Receita Fechada YTD · Atingimento · Saúde do Funil — todos com barra de progresso visual contra a meta de R$ 34 Mi.
- 5 KPIs secundários: Conversão · Pipeline · Ticket Médio · Hunters Ativos · Procurações.
- Gráfico timeline duplo: Receita Fechada (barra) + Pipeline (barra clara) + Meta (linha verde).
- Ranking Top 8 hunters · Mix Novo/Upsell · Pipeline por Temperatura.
- **Insights automáticos** (texto auto-gerado): alerta se atingimento < 85%, se standby > 20% da receita, etc.

### Página 2 · **Conversão** (foco gerencial)
- KPIs: Pipeline do mês · Fechado dentro × fora da competência · Standby · Taxa de fechamento.
- **Funil visual** com 5 estágios (Pipeline → Ativas → Negociação → Em Assinatura → Assinadas).
- **Doughnut Dentro × Fora da Competência** (exatamente a regra do briefing).
- Tempo médio de fechamento mês a mês.
- Tabela por hunter com colunas: Entradas, Fechadas no mês, Fechadas fora, Standby, Conv. %, Receita, Ticket médio.

### Página 3 · **Hunters** (foco produtividade individual)
- 5 KPIs de destaque: Top Receita · Top Conversão · Top Procurações · Mais Eficiente · Em Atenção.
- **Ranking horizontal** Realizado vs Meta lado a lado.
- **Matriz de Performance** (scatter): eixo X = Receita, eixo Y = Conversão, tamanho do ponto = volume de procurações. Identifica quadrantes (estrelas, alta produtividade baixa conversão, etc.).
- **Cards individuais** com mini-gauge por hunter (1 card cada).
- Tabela completa com status (META BATIDA · NO RITMO · ATENÇÃO · CRÍTICO).

### Página 4 · **Procurações** (foco operacional)
- 4 cards limpos: Total · Novas · Renovações · % Renovação.
- Gráfico mensal Novas vs Renovações empilhado.
- Produção por hunter (barras horizontais).
- **Heatmap Hunter × Mês** (intensidade dourada = mais procurações).
- Tabela com % renovação por hunter.

### Página 5 · **Metas** (foco estratégico)
- **3 Hero KPIs** alinhados à regra do briefing: Meta 34 Mi · Realizado · Excedente (acima de 34 Mi).
- 4 KPIs operacionais: Faltante · Run-Rate Mensal Necessário · Projeção Linear · Pipeline Necessário.
- **Curva acumulada** Realizado vs reta linear da meta (visualiza se está acima/abaixo do ritmo).
- Ranking de atingimento por hunter (barras coloridas por faixa).

### Página 6 · **Pipeline** (foco saúde do funil)
- KPIs por temperatura: Total · Quente · Morno · Frio · Standby.
- Distribuição por Status · Doughnut por temperatura.
- Pipeline por UF · Mix por Produto.
- **Top 15 oportunidades em aberto** com empresa, hunter, status, temperatura.

### Padrões visuais aplicados
- Paleta: dark navy (`#070d1c → #0c1530`) + dourado executivo (`#f4cf7a`) + acentos pontuais (verde, vermelho, azul).
- Hierarquia: Hero KPI (24px+) → KPI normal (18-22px) → Card de gráfico (12-13px header).
- Gradientes suaves nos cards, sombras profundas, bordas finas para hierarquia.
- Pills coloridas para status (CRÍTICO/ATENÇÃO/NO RITMO/META BATIDA) em vez de só números.
- Indicador "DEMO · v1" no canto superior direito para diferenciar de produção.

---

## 6. Próximos passos sugeridos

| Prioridade | Ação | Esforço |
|---|---|---|
| Alta | Validar a DEMO e me dar feedback de o que ajustar antes de subir como nova produção | — |
| Alta | Normalizar a base (criar `data_quality.py` que padroniza Hunter/Produto/Região/Mês na origem) | 30 min |
| Alta | Adicionar campos de data real (Data Entrada, Data Fechamento) na planilha — não só MES | externo |
| Média | Estruturar a aba `AGENDA` como tabela transacional para entrar no modelo | externo |
| Média | Migrar de HTML para Power BI .pbix (mantendo HTML como vista pública via Netlify) | 4-6h |
| Baixa | Adicionar análise preditiva (regressão linear de meta vs tempo) | 2h |

---

## 7. O que abre na DEMO

Arquivo: `One_Page_Comercial_2026_DEMO.html` (~210 KB, autocontido, abre direto no navegador).

- **6 abas navegáveis** com transições suaves.
- **4 filtros globais** que afetam todas as páginas: Competência (mês de entrada), Hunter, UF, Tipo (Novo/Upsell).
- **Filtros persistem** ao trocar de aba.
- Dados embutidos da última leitura (16/05/2026 13:45).
- Mobile-friendly (até ~1280px o grid colapsa).

A produção atual (`One_Page_Comercial_2026.html`) **não foi tocada**.
