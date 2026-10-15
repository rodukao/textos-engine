# textos-engine

Motor de documentos `.docx` do [Textos](https://github.com/rodukao/textos): um fork do
[SuperDoc](https://github.com/superdoc/docx-editor) a partir da linha **1.x** (branch `v1`, commit
`66c22186c7`, equivalente ao pacote `superdoc@1.47.2`), distribuído sob a mesma licença,
**GNU AGPL-3.0** ([LICENSE](LICENSE)).

## Por que a linha 1.x

Versões posteriores do SuperDoc 2.x dependem de `@superdoc/docx-engine`, um pacote proprietário
cuja licença é incompatível com um editor de código aberto. Este fork parte da última linha
inteiramente AGPL e **não contém, não usa e não se baseia** naquele pacote.

## Regras deste fork

- A referência de fidelidade é o **Microsoft Word** e as especificações públicas
  (ECMA-376 / ISO 29500, MS-DOCX, MS-OI29500). O comportamento, o código e os resultados do
  pacote proprietário do SuperDoc 2.x **nunca** são usados como referência.
- Toda mudança é medida pela bancada do Textos (`bench/` no repositório do Textos), que compara
  com o Word a fidelidade ao salvar, a exibição e o desempenho.
- Antes de adicionar qualquer dependência, confira a licença de toda a árvore.

## Branches

- `textos` — branch principal do fork (padrão).
- `v1` — linha 1.x original do SuperDoc, sem alterações.

## Build

```bash
pnpm install
pnpm run build:superdoc                         # gera packages/superdoc/dist
pnpm --dir packages/superdoc pack               # gera o .tgz usado pelo Textos
```

Requer pnpm 10.25.0. No Windows, o passo `prepare` (hooks do git) falha por usar sintaxe de
bash; isso não afeta o build.
