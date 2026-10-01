# Apuração de dados oficiais da Anatel — autenticação de origem de chamadas

Dados e rotina de apuração usados na monografia **"Autenticação de Origem de Chamadas Telefônicas no Brasil"**, de Tatianno Ferreira Alves, Especialização em Tecnologias e Sistemas de Informação da Universidade Federal do ABC (UFABC), 2026.

## Por que este repositório existe

O capítulo de metodologia da monografia promete que o leitor pode recalcular cada número derivado do trabalho. **Essa promessa não é cumprível a partir da origem.** Os painéis de dados da Anatel são rebaixados para períodos novos, de modo que quem os exportar hoje não obtém a exportação de agosto de 2026, e sim outra, com valores diferentes. A página do conjunto de contratos de interconexão registra, por exemplo, "última alteração em um arquivo" em 1º de setembro de 2026, posterior à extração aqui depositada.

Este repositório resolve isso guardando **os arquivos exatos que foram usados** e **a rotina que os processa**.

## Como executar

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl
.venv/bin/python apuracao.py
```

A rotina apura cada número, compara o resultado com os valores registrados na constante `ESPERADO` e acusa divergência. São **23 valores verificados**. A última execução sem divergência foi em 1º de outubro de 2026.

Os arquivos `.csv` usam apenas a biblioteca padrão. O `openpyxl` é necessário para as planilhas `.xlsx`.

## O que cada apuração sustenta

| Apuração | O que produz | Argumento que sustenta |
|---|---|---|
| Contagem de entidades, fixo e móvel | Grupos econômicos distintos por ano, de junho de 2021 a junho de 2026 | A superfície de origem se ampliou apenas onde a barreira de entrada é baixa. O fixo passou de 154 a 315 grupos enquanto perdia 38% da base; o móvel, que exige espectro leiloado, foi de 11 a 17 |
| Distribuição de acessos por grupo | Faixas de tamanho em junho de 2026, incluindo a cauda | Contagem de entidades, e não participação de mercado. Para o modelo de ameaça do *spoofing*, uma prestadora com 300 acessos origina sinalização com a mesma capacidade de uma com 5 milhões |
| Topologia da interconexão | Prestadoras, pares bilaterais e grau de concentração | O modelo de confiança do SS7 presume um conjunto pequeno e conhecido de operadoras. O grafo real tem centenas de nós |
| Tamanho médio do grupo | Contraste entre fixo e móvel na mesma data | Assimetria de cerca de 280 vezes entre o grupo móvel médio e o fixo médio |
| Cobertura móvel por moradores | Proporção coberta por 4G/5G | A rede de acesso móvel é predominantemente IP |
| Meio de acesso da telefonia fixa | Participação de cada meio no início e no fim da série | Migração para IP no último quilômetro |

## Procedência dos dados

Todos os arquivos em `dados/` são **redistribuídos sem modificação**. Apenas nomes de arquivo e de pasta foram normalizados, para remover espaços e extensões duplicadas, o que não altera o conteúdo.

### Contratos de interconexão

- **Órgão:** Agência Nacional de Telecomunicações (Anatel)
- **Conjunto:** Competição — Contratos de Interconexão
- **Origem:** https://dados.gov.br/dados/conjuntos-dados/contratos-de-interconexao
- **Licença declarada na página do conjunto:** Creative Commons Attribution (CC BY)
- **Área técnica responsável:** Gerência de Monitoramento das Relações Entre Prestadoras (CPRP/SCP)
- **Catalogação:** 23 de outubro de 2020. **Periodicidade de atualização:** não declarada
- **Extração aqui depositada:** 1º de agosto de 2026
- **Arquivos:** `contratos_interconexao.csv`, `contratos_mvno.csv`, `empresas_credenciadas_vigentes.csv` e o glossário de campos do próprio conjunto, renomeado para `glossario_campos.ods`

> O conjunto de origem contém ainda `contratos_compartilhamento.csv` e `contratos_ran_sharing.csv`, que esta apuração não usa e que por isso não foram depositados.

### Painéis de acessos e de cobertura

- **Órgão:** Agência Nacional de Telecomunicações (Anatel)
- **Origem:** https://informacoes.anatel.gov.br/paineis/ (painéis de acessos de telefonia fixa, de telefonia móvel e de cobertura móvel)
- **Extrações:** 1º e 7 de agosto de 2026, com período de referência de junho de cada ano
- **Condição de uso:** dados abertos nos termos do Decreto nº 8.777, de 11 de maio de 2016, que institui a Política de Dados Abertos do Poder Executivo federal. A Anatel define dados abertos como disponibilizados "sob licença aberta que permita sua livre utilização, consumo ou cruzamento, limitando-se o interessado a creditar a autoria ou a fonte" (https://www.gov.br/anatel/pt-br/dados/dados-abertos)

> Os painéis não trazem campo de licença por exportação, ao contrário da página do conjunto no `dados.gov.br`. Por isso a condição de uso está declarada pela política, e não por uma licença nomeada.

**Crédito da fonte, para quem reutilizar:** ANATEL — Agência Nacional de Telecomunicações, dados abertos e painéis de dados, extrações de 1º e 7 de agosto de 2026.

## Armadilhas dos dados, tratadas pela rotina

Cada uma está documentada no *docstring* da função correspondente.

- **Série de porte da prestadora, não usada.** Tem quebra de critério em janeiro de 2015 e ponto fora da curva em junho de 2021. A contagem de entidades é apurada dos rankings anuais, imunes a ambos.
- **Cobertura móvel, usar "moradores" e nunca "área".** Pela mesma base, a cobertura por área nacional é de 16,62%, porque a área não povoada da Amazônia domina o cálculo.
- **Meio de acesso não é protocolo de sinalização.** Fibra no último quilômetro torna a voz em IP provável, e não prova que a interconexão seja SIP.
- **Contratos protocolados não são interconexões em vigor.** O conjunto é acumulado histórico entre 2006 e 2025, não fotografia do presente. E há defasagem de publicação em 2025 e 2026, que não é interrupção real.
- **Não existe campo de tecnologia de interconexão** em nenhum arquivo do conjunto. Se um enlace é SIP ou SS7 não é informação pública.
- **Exportação do painel pode vir com o período errado.** A verificação de monotonicidade na série do fixo existe para pegar isso, e foi assim que se detectou, em 7 de agosto de 2026, uma exportação rotulada como 2022 que continha os dados de junho de 2026.
- **Validação do procedimento.** A Anatel publica pronto o topo do ranking, e não a cauda, que é o que o argumento usa. A rotina confere a própria agregação contra o agregado oficial: se reproduz o que a Agência publica, o mesmo método aplicado à cauda está validado.

## Licença

O código de `apuracao.py` e o texto deste README estão sob a licença MIT, em `LICENSE`.

**Os arquivos em `dados/` não são de autoria do autor deste repositório.** São dados públicos da Anatel, redistribuídos nas condições declaradas na seção de procedência.
