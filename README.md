

# Links úteis:
https://learn.microsoft.com/pt-br/sharepoint/dev/spfx/set-up-your-development-environment

# WebPart Comunicados 

Lista PRECISA ter o nome Comunicados
Precisa ter as seguintes colunas:

Título: título comunicado
Descrição : várias linhas
StatusComunicado: Coluna do tipo opção - Ativo ou Inativo
CorCard : uma linha
Icone: uma linha
CorTexto: uma linha
Area: Opção


# WebPart Aniversariantes
Lista PRECISA ter o nome Aniversariantes


Título : nome do aniversariante
Data de Nascimento: coluna uma linha
Foto: coluna tipo foto
Email: uma linha
Funcao: uma linha

# WebPart Destaques do Mês

Lista PRECISA ter o nome Destaques do Mês

Título: Nome do colaborador
Conquista: várias linhas
Foto: coluna tipo foto
Area: uma linha

# WebPart Movimentação de Pessoas 

Lista PRECISA ter o nome Movimentação de Pessoas

Título: Nome do Colaborador
Foto: coluna tipo foto
Funcao: uma linha
Departamento: uma linha
TipoDestaque: coluna do tipo opção; precisa ter EXATAMENTE essas opções: Promoção, Admissão, Desligamento, Movimentação
Email: uma linha

# WebPart FAQ
Nome da lista tanto faz desde que identifique que é categoria e pergunta
Precisa de duas listas: Categoria e Perguntas

Lista de Categoria:
Nome da Categoria: coluna de Título
Sequence: tipo número
IsActive: Tipo sim ou não


Lista de Perguntas:
Sequencia: tipo número
Pergunta: coluna título
Resposta: várias linhas
Categoria: tipo pesquisa sendo o nome da categoria da lista de categoria
Visililidade: tipo sim ou não


# WebPart Instagram

Só precisa inserir nela o @ a ser mostrado

# WebPart do Calendário

Pega automático o calendário corporativo

# my-first-web-part

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
