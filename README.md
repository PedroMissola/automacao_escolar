# Motor de Automação e Auditoria de Dados (Data Pipeline)

![Node.js](https://img.shields.io/badge/Node.js-339933?style=for-the-badge&logo=nodedotjs&logoColor=white)
![SQLite](https://img.shields.io/badge/SQLite-07405E?style=for-the-badge&logo=sqlite&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Status](https://img.shields.io/badge/Status-Concluído-success?style=for-the-badge)

> Um sistema robusto desenvolvido em **Node.js** para processar, validar e auditar milhares de registros provenientes de planilhas brutas. O projeto atua como um funil inteligente que limpa os dados, corrige falhas humanas e padroniza as informações sem que o usuário final precise abandonar sua ferramenta de trabalho principal (Excel).

## 📋 Índice

- [O Problema e a Solução](#-o-problema-e-a-solução)
- [Tecnologias Utilizadas](#-tecnologias-utilizadas)
- [Como a Arquitetura Funciona](#-como-a-arquitetura-funciona)
- [Como Executar o Projeto](#-como-executar-o-projeto)

---

## O Problema e a Solução

**O Cenário**  
Uma base de dados bruta contendo respostas de estudantes e notas. Havia total falta de padronização (variações no nome das matérias), erros graves na digitação das notas (valores muito acima do limite) e grande dificuldade para auditar manualmente as atividades pendentes.

**A Solução**  
Um pipeline de dados local que roda nos bastidores, lê as planilhas originais, aplica regras de negócio via código e banco de dados, e exporta um novo arquivo organizado em abas (`PROCESSADAS`, `INCONSISTÊNCIAS` e `NÃO PROCESSADOS`), além de disparar e-mails para os responsáveis sobre atividades não realizadas.

---

## Tecnologias Utilizadas

- **[Node.js](https://nodejs.org/):** Orquestração da lógica principal de processamento.
- **[SQLite (better-sqlite3)](https://github.com/WiseLibs/better-sqlite3):** Banco de dados relacional local (em modo WAL) para controle de estado, rastreabilidade e alta performance.
- **[xlsx (SheetJS)](https://sheetjs.com/):** Biblioteca nativa para leitura, manipulação e escrita dos arquivos Excel.
- **[Resend API](https://resend.com/):** Integração para disparo automatizado de relatórios e alertas por e-mail.
- **Crypto (Nativo Node):** Geração de hashes de segurança (SHA1) para chaves únicas.

---

## Como a Arquitetura Funciona na Prática

1. **Prevenção de Concorrência (Locking)**
   O script utiliza a flag `'wx'` (`fs.openSync`) para criar um arquivo de trava (`.lock`) com o PID do processo. Isso garante que, em um ambiente de *cron jobs* ou execução múltipla, o script não rode duas vezes simultaneamente, evitando corrupção de dados.

2. **Classificação e Padronização Textual (Regex)**
   Uma função normaliza as strings (removendo acentos e espaços extras). Através de um mapeamento estruturado com Expressões Regulares, o sistema isola o texto principal, identifica a matéria (resolvendo ambiguidades com parênteses) e remonta o nome da atividade em um formato estrito (ex: `Lista 08 de Matemática Frente 2`).

3. **Validação Matemática Restrita**
   O código converte as entradas, removendo falhas de formatação. Se a nota ultrapassar o valor máximo permitido, um algoritmo de correção em cascata tenta ajustar o valor matematicamente (dividindo por 10 ou 100), assumindo erro na casa decimal. O ajuste é salvo com uma tag de observação. Erros irreparáveis são tipados como `INCONSISTÊNCIA` e isolados.

4. **Persistência Incremental e Idempotência**
   Para não reprocessar registros, é gerado um hash SHA1 combinando o ID da atividade e a matrícula. Os dados são inseridos em bloco no SQLite via `transaction` utilizando `ON CONFLICT(chave) DO NOTHING`. Isso torna o sistema *idempotente*: ele simplesmente ignora dados já lidos de forma performática.

5. **Exportação e Alertas Automáticos**
   O banco é consultado para gerar a planilha final, separando os registros validados dos erros. Simultaneamente, o sistema mapeia atividades não realizadas e agrupa por aluno, injetando os dados em templates HTML para envio imediato aos gestores utilizando a API da Resend.

---

## Como Executar o Projeto

### Pré-requisitos
- Node.js instalado (versão 16+ recomendada)
- Chave de API do [Resend](https://resend.com/) (para envio de e-mails)

### Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/PedroMissola/automacao_escolar.git
   ```

2. Instale as dependências:
   ```bash
   npm install
   ```

3. Configure as variáveis de ambiente (Crie um arquivo `.env` na raiz):
   ```env
    PLANILHA_ORIGEM=./notas.xlsx
    PLANILHA_SAIDA=./notas-processadas.xlsx
    BANCO_SQLITE=./notas.db
    RESEND_API_KEY=apidoresend
    EMAIL_REMETENTE=remetente@email.com
    EMAIL_DESTINATARIO=destinatario@email.com
   ```

4. Coloque sua planilha bruta na pasta de entrada e execute:
   ```bash
   npm start
   ```

---

<p align="center">
  <i>Desenvolvido com foco em Governança de Dados, Eficiência Operacional e Metodologias Ágeis.</i>
</p>
