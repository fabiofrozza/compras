---
name: atualizar-git
description: Verificar alterações nos arquivos e providenciar branches e commits no git
allowed-tools:
- Bash(git diff *)
- Bash(git status)
- Bash(git log *)
- Bash(git issue *)
- Bash(git add *)
- Bash(git commit *)
- Bash(git branch *)
- Bash(git checkout *)
- Bash(git merge)
- Bash(cd *)
- Bash(grep *)
---

# Git

## Branches

- **`main`**: Branch de produção. Cada merge aqui gera uma release.
- **`develop`**: Branch de integração. Todo desenvolvimento converge aqui antes de ir para `main`.
- **Branches de trabalho**: Criados a partir do `develop` e mergeados de volta nele (exceto `hotfix/`, criado a partir de `main`).

| Prefixo       | Uso                                      | Exemplo                          |
|----------------|------------------------------------------|----------------------------------|
| `feat/`     | Nova funcionalidade                      | `feat/auth-login`             |
| `fix/`         | Correção de bug                          | `fix/toast-duplicado`            |
| `hotfix/`      | Correção urgente em produção (a partir de `main`) | `hotfix/crash-on-startup` |
| `refactor/`    | Refatoração sem mudança de comportamento | `refactor/modularizar-scripts`   |
| `style/`       | Alterações visuais (CSS, layout, UI)     | `style/tab-row-column-layout`    |
| `docs/`        | Apenas documentação                      | `docs/atualizar-readme`          |
| `chore/`       | Manutenção, configs, dependências        | `chore/atualizar-deps`           |

**Nomeação de branches**: `prefixo/descricao-em-kebab-case` — máximo 3–4 palavras, sem acentos, tudo em minúsculas. Opcional: incluir número do issue (ex: `feat/42-auth-login`).

## Fluxo de trabalho

1. **Criar branch** a partir do `develop`: `git checkout develop && git checkout -b prefixo/descricao`
2. **Commitar no branch de trabalho**, nunca diretamente no `develop`.
3. **Mergear no `develop`** quando o trabalho estiver concluído: `git checkout develop && git merge prefixo/descricao`
4. **Deletar o branch de trabalho** após o merge: `git branch -d prefixo/descricao` — **obrigatório, nunca pular esta etapa.**
5. **Atualizar branches em andamento** com mudanças do `develop` quando necessário:
   - **Merge**: `git checkout feature/... && git merge develop` — preserva o histórico dos branches (cria um "balãozinho" no grafo).
   - **Rebase**: `git checkout feature/... && git rebase develop` — reescreve os commits sobre o `develop`, resultando em histórico linear e limpo. Não usar se o branch já foi compartilhado (pushed) com outros.

> **REGRA CRÍTICA:** O assistente de IA **NUNCA** deve mergear no `main` a não ser que o usuário **expressamente autorize** naquela conversa específica. Essa autorização não pode ser inferida, antecipada ou presumida — deve ser explícita e inequívoca.

## Commits

- Usar [Conventional Commits](https://www.conventionalcommits.org/): `tipo(escopo opcional): descrição curta`
- Tipos: `feat`, `fix`, `refactor`, `style`, `chore`, `perf`, `docs`, `test`, `ci`, `build`
- Mensagem em inglês, concisa, em letras minúsculas (exceto nomes próprios)
- Descrição no imperativo, sem ponto final, máximo ~72 caracteres na primeira linha
- Cada commit deve ser **atômico** (uma mudança lógica por commit)
- Não misturar mudanças de escopo diferente no mesmo commit (ex: refactor + style)
- Breaking changes: usar `tipo!: descrição` e incluir `BREAKING CHANGE:` no rodapé

## Versionamento (SemVer)

Tags seguem [Semantic Versioning](https://semver.org/): `vMAJOR.MINOR.PATCH`
- **MAJOR**: mudanças que quebram compatibilidade
- **MINOR**: novas funcionalidades retrocompatíveis
- **PATCH**: correções de bugs