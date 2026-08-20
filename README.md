# Skill: Agravo de Instrumento Municipal

Skill para o Claude (Claude Code / Cowork) que gera minutas de **Agravo de Instrumento** em nome do Município do Rio de Janeiro, entregues em `.docx` já formatado no padrão das peças da Procuradoria.

A skill combina duas coisas: um conjunto de instruções que ensinam o modelo a estruturar e redigir a peça, e um script Node que converte o conteúdo aprovado em um documento Word com papel timbrado, cabeçalho, numeração de páginas e tipografia padronizada.

## O que a skill faz

Quando acionada, ela conduz a redação da peça na seguinte ordem:

1. Endereçamento ao Tribunal de Justiça ou ao Conselho Recursal
2. Identificação do processo originário
3. Qualificação das partes (agravante e agravado)
4. Abertura, com a fórmula de interposição do artigo 1.015 do CPC
5. Informações para cumprimento do artigo 1.016 do CPC
6. Razões do agravante, em seção própria: síntese da decisão agravada, demonstração do cabimento, razões para a reforma e, quando for o caso, pedido de efeito suspensivo ou de antecipação da tutela recursal
7. Pedidos
8. Parágrafo de intimações
9. Fechamento com data, nome, cargo e matrícula

O detalhamento de cada seção, incluindo as hipóteses de cabimento do artigo 1.015 e a taxatividade mitigada do Tema 988 do STJ, está em [`references/estrutura.md`](references/estrutura.md).

## Estrutura do repositório

```
SKILL.md                            instruções da skill (frontmatter + regras de redação)
references/estrutura.md             detalhamento da estrutura da peça
scripts/gerar_agravo_instrumento.js gerador do .docx
scripts/assets/image1.png           imagem do papel timbrado usada no cabeçalho
scripts/package.json                dependências do script (docx)
```

## Como instalar a skill

Copie a pasta inteira para o diretório de skills do seu ambiente Claude:

```bash
git clone https://github.com/simiao-cavalcante/agravo-de-instrumento.git
cp -R agravo-de-instrumento ~/.claude/skills/agravo-de-instrumento
```

A skill é reconhecida pelo arquivo `SKILL.md`, cujo frontmatter define o nome e os gatilhos de acionamento. Basta pedir ao Claude algo como "minute um agravo de instrumento nesse processo" para que ela entre em ação.

## Como usar o gerador isoladamente

O script também funciona fora do Claude. Instale a dependência e rode passando um JSON de entrada:

```bash
cd scripts
npm install
node gerar_agravo_instrumento.js entrada.json saida.docx
```

O JSON de entrada segue este formato:

```json
{
  "enderecamento": "EXMO. SR. DR. DESEMBARGADOR PRESIDENTE DO TRIBUNAL DE JUSTIÇA DO ESTADO DO RIO DE JANEIRO",
  "processo": "0000000-00.0000.0.00.0000",
  "tipo_recurso": "Agravo de Instrumento",
  "embargante": "MUNICÍPIO DO RIO DE JANEIRO",
  "embargado": "NOME DA PARTE CONTRÁRIA",
  "abertura": "O MUNICÍPIO DO RIO DE JANEIRO, inconformado com a r. decisão interlocutória...",
  "secoes": [
    { "tipo": "h1", "texto": "RAZÕES DO AGRAVANTE" },
    { "tipo": "corpo", "texto": "Texto do parágrafo." },
    { "tipo": "pedidos", "texto": "Ante o exposto, requer..." }
  ],
  "intimacoes": "Requer sejam as intimações...",
  "fechamento": {
    "data": "01 de janeiro de 2025",
    "nome": "Nome do Procurador",
    "cargo": "Procurador do Município do Rio de Janeiro",
    "matricula": "00/000.000-0"
  }
}
```

O campo `tipo` de cada seção aceita `h1`, `h2`, `corpo` e `pedidos`, que controlam o nível de título e o alinhamento aplicados no documento final.

## Requisitos

Node.js 18 ou superior e a biblioteca [`docx`](https://www.npmjs.com/package/docx), instalada via `npm install`.

## Licença

MIT. Veja [LICENSE](LICENSE).
